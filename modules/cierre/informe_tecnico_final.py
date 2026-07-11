from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import streamlit as st

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

from config.constants import OUTPUT_DIR
from services.json_service import guardar_datos_json
from services.openai_service import generar_json_openai
from utils.nombres_archivo import safe_filename
from utils.textos import limpiar_texto
from utils.validaciones import validar_campos_obligatorios


VERSION_INFORME_TECNICO_FINAL = (
    "VERSION_APROBADA_TOC_ESTABLE_SIN_SEGFAULT"
)
CODIGO_FORMATO_INFORME = "GCDTP-F-023 V01"
NOMBRE_PLANTILLA_INFORME = "GCDTP-F-023_V01_Formato_Informe_Final.docx"

TIPOS_PROYECTO_INFORME = [
    "Software",
    "Diseño industrial",
    "Electrónica y automatización",
    "Inteligencia artificial",
    "Desarrollo de marca e identidad visual",
    "Prototipo agroindustrial",
    "Ficha técnica",
    "Otro",
]

CLASIFICACIONES_INFORMACION = [
    "Pública",
    "Pública Clasificada",
    "Pública Reservada",
]

METODOLOGIAS_DESARROLLO = [
    "Metodologías ágiles",
    "Scrum",
    "Kanban",
    "Modelo en cascada",
    "Modelo espiral",
    "Design Thinking",
    "Doble Diamante",
    "Diseño Centrado en el Usuario (DCU)",
    "Diseño para Manufactura y Ensamble (DFMA)",
    "Lean Startup",
    "Stage-Gate",
    "Desarrollo iterativo de prototipos",
    "Ingeniería de sistemas y modelo V",
    "CRISP-DM para proyectos de datos e inteligencia artificial",
    "DMAIC / Six Sigma",
    "Investigación aplicada y validación experimental",
    "TRIZ para solución inventiva de problemas",
    "Otra",
]

CLAVES_CONTENIDO = [
    "introduccion",
    "planteamiento_problema",
    "objetivo_general",
    "objetivos_especificos",
    "estado_arte_tecnica",
    "metodologia_desarrollo",
    "actividades_corregidas",
    "desarrollo_proyecto",
    "resultados_obtenidos",
    "analisis_viabilidad",
    "propiedad_transferencia",
    "impacto_proyecto",
    "conclusiones",
    "referencias_bibliograficas",
    "anexos",
]


# =====================================================
# UTILIDADES GENERALES
# =====================================================

def obtener_ruta_plantilla() -> Path:
    raiz_proyecto = Path(__file__).resolve().parents[2]

    candidatos = [
        raiz_proyecto / "resources" / NOMBRE_PLANTILLA_INFORME,
        raiz_proyecto / "recursos" / NOMBRE_PLANTILLA_INFORME,
        Path("resources") / NOMBRE_PLANTILLA_INFORME,
        Path("recursos") / NOMBRE_PLANTILLA_INFORME,
    ]

    for ruta in candidatos:
        if ruta.exists():
            return ruta

    raise FileNotFoundError(
        "No se encontró la plantilla oficial del Informe Final. "
        f"Guárdala como resources/{NOMBRE_PLANTILLA_INFORME}"
    )


def fecha_mes_anio_espanol(fecha: date) -> str:
    meses = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }
    return f"{meses[fecha.month]} {fecha.year}"


def dividir_lineas(texto: str) -> list[str]:
    texto = str(texto or "").replace(";", "\n")
    elementos = [
        limpiar_texto(linea.strip(" -•\t"))
        for linea in texto.splitlines()
        if limpiar_texto(linea.strip(" -•\t"))
    ]

    if len(elementos) <= 1 and "," in texto:
        elementos = [
            limpiar_texto(item)
            for item in texto.split(",")
            if limpiar_texto(item)
        ]

    return elementos


def dividir_actividades(texto: str) -> list[str]:
    """Convierte un solo campo de actividades en una lista limpia."""
    contenido = str(texto or "").replace(";", "\n")
    actividades: list[str] = []

    for linea in contenido.splitlines():
        linea = re.sub(
            r"^\s*(?:[-•*]|\d+[.)-]?)\s*",
            "",
            linea,
        )
        linea_limpia = limpiar_texto(linea)

        if linea_limpia:
            actividades.append(linea_limpia)

    if len(actividades) <= 1 and contenido.count(",") >= 2:
        actividades = [
            limpiar_texto(item)
            for item in contenido.split(",")
            if limpiar_texto(item)
        ]

    return actividades


def limpiar_parrafo(parrafo: Paragraph) -> None:
    for elemento in list(parrafo._p):
        if elemento.tag != qn("w:pPr"):
            parrafo._p.remove(elemento)


def escribir_parrafo(
    parrafo: Paragraph,
    texto: str,
    *,
    tamano: float = 11,
    negrita: bool = False,
    alineacion=None,
) -> None:
    limpiar_parrafo(parrafo)
    run = parrafo.add_run(str(texto or ""))
    run.bold = negrita
    run.font.name = "Arial"
    run.font.size = Pt(tamano)

    if alineacion is not None:
        parrafo.alignment = alineacion


def escribir_celda(celda, texto: object) -> None:
    celda.text = ""
    parrafo = celda.paragraphs[0]
    parrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = parrafo.add_run(str(texto or ""))
    run.font.name = "Arial"
    run.font.size = Pt(10)


def texto_elemento_xml(elemento) -> str:
    return "".join(
        nodo.text or ""
        for nodo in elemento.xpath('.//*[local-name()="t"]')
    ).strip()


def buscar_parrafo(
    documento: Document,
    texto_objetivo: str,
    *,
    coincidencia_exacta: bool = False,
) -> Paragraph:
    objetivo = " ".join(str(texto_objetivo).split()).casefold()

    for parrafo in documento.paragraphs:
        texto = " ".join(parrafo.text.split()).casefold()

        if coincidencia_exacta and texto == objetivo:
            return parrafo

        if not coincidencia_exacta and objetivo in texto:
            return parrafo

    raise ValueError(
        f"No se encontró el apartado '{texto_objetivo}' en la plantilla."
    )


def parrafo_siguiente(parrafo: Paragraph) -> Paragraph:
    elemento = parrafo._p.getnext()

    while elemento is not None:
        if elemento.tag == qn("w:p"):
            return Paragraph(elemento, parrafo._parent)
        elemento = elemento.getnext()

    raise ValueError(
        f"No se encontró un párrafo después de '{parrafo.text}'."
    )


def insertar_parrafo_despues(
    parrafo: Paragraph,
    texto: str = "",
    estilo: str | None = None,
) -> Paragraph:
    nuevo_p = OxmlElement("w:p")
    parrafo._p.addnext(nuevo_p)
    nuevo = Paragraph(nuevo_p, parrafo._parent)

    if estilo:
        try:
            nuevo.style = estilo
        except KeyError:
            pass

    if texto:
        escribir_parrafo(nuevo, texto)

    return nuevo


def eliminar_instrucciones_y_control_cambios(documento: Document) -> None:
    """Elimina desde INSTRUCCIONES hasta antes de las propiedades de sección."""
    cuerpo = documento._element.body
    eliminar = False

    for elemento in list(cuerpo):
        if elemento.tag == qn("w:sectPr"):
            continue

        texto = texto_elemento_xml(elemento)

        if (
            elemento.tag == qn("w:p")
            and texto.strip().casefold().startswith("instrucciones")
        ):
            eliminar = True

        if eliminar:
            cuerpo.remove(elemento)


def marcar_actualizacion_campos(documento: Document) -> None:
    settings = documento.settings.element
    update_fields = settings.find(qn("w:updateFields"))

    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)

    update_fields.set(qn("w:val"), "true")

    for campo in documento._element.xpath(
        './/*[local-name()="fldChar" and @w:fldCharType="begin"]'
    ):
        campo.set(qn("w:dirty"), "true")


def marcar_clasificacion(documento: Document, clasificacion: str) -> None:
    """
    Escribe una X dentro del cuadro de texto de la clasificación seleccionada,
    conservando los cuadros y la diagramación original de la portada.
    """
    tabla_portada = documento.tables[0]
    fila = tabla_portada.rows[5]

    for celda in fila.cells:
        nombre = " ".join(celda.text.split())
        seleccionada = nombre.casefold() == clasificacion.casefold()

        contenidos = celda._tc.xpath(
            './/*[local-name()="txbxContent"]'
        )

        for contenido in contenidos:
            parrafos = contenido.xpath('./*[local-name()="p"]')

            if not parrafos:
                nuevo_p = OxmlElement("w:p")
                contenido.append(nuevo_p)
                parrafos = [nuevo_p]

            for p_xml in parrafos:
                for hijo in list(p_xml):
                    if hijo.tag != qn("w:pPr"):
                        p_xml.remove(hijo)

                p_pr = p_xml.find(qn("w:pPr"))
                if p_pr is None:
                    p_pr = OxmlElement("w:pPr")
                    p_xml.insert(0, p_pr)

                jc = p_pr.find(qn("w:jc"))
                if jc is None:
                    jc = OxmlElement("w:jc")
                    p_pr.append(jc)
                jc.set(qn("w:val"), "center")

                if seleccionada:
                    run = OxmlElement("w:r")
                    r_pr = OxmlElement("w:rPr")

                    negrita = OxmlElement("w:b")
                    r_pr.append(negrita)

                    tamano = OxmlElement("w:sz")
                    tamano.set(qn("w:val"), "20")
                    r_pr.append(tamano)

                    run.append(r_pr)
                    texto = OxmlElement("w:t")
                    texto.text = "X"
                    run.append(texto)
                    p_xml.append(run)


def llenar_tabla_informacion_general(
    documento: Document,
    datos: dict,
) -> None:
    tabla = documento.tables[1]

    valores = [
        datos.get("nombre_talento", ""),
        datos.get("nombre_proyecto", ""),
        datos.get("codigo_proyecto", ""),
        datos.get("nombre_experto", ""),
        datos.get("linea_tecnologica", ""),
        datos.get("trl_inicial", ""),
        datos.get("trl_alcanzado", ""),
        datos.get("tecnoparque", ""),
        datos.get("fecha_entrega_texto", ""),
    ]

    for indice, valor in enumerate(valores):
        escribir_celda(tabla.rows[indice].cells[1], valor)


def formatear_lista_documento(
    elementos: list[str],
    *,
    prefijo: str = "•",
) -> str:
    return "\n".join(
        f"{prefijo} {elemento}"
        for elemento in elementos
        if limpiar_texto(elemento)
    )


def escribir_lista_en_parrafos(
    parrafo_inicial: Paragraph,
    elementos: list[str],
    *,
    prefijo: str = "•",
) -> Paragraph:
    """Escribe cada elemento en un párrafo independiente y retorna el último."""
    elementos_limpios = [
        limpiar_texto(str(elemento))
        for elemento in elementos
        if limpiar_texto(str(elemento))
    ]

    if not elementos_limpios:
        elementos_limpios = ["No se suministró información para este apartado."]

    actual = parrafo_inicial

    for indice, elemento in enumerate(elementos_limpios):
        if indice > 0:
            actual = insertar_parrafo_despues(
                actual,
                estilo="Normal",
            )

        escribir_parrafo(
            actual,
            f"{prefijo} {elemento}",
            tamano=11,
            alineacion=WD_ALIGN_PARAGRAPH.LEFT,
        )
        actual.paragraph_format.space_after = Pt(4)
        actual.paragraph_format.left_indent = Cm(0.35)
        actual.paragraph_format.first_line_indent = Cm(-0.35)

    return actual


def eliminar_parrafo(parrafo: Paragraph) -> None:
    elemento = parrafo._element
    elemento.getparent().remove(elemento)
    parrafo._p = None
    parrafo._element = None


def configurar_estilos_y_tabla_contenido(documento: Document) -> None:
    """
    Garantiza que los títulos del informe usen estilos de Word y que la tabla
    de contenido se actualice con la paginación real al abrir el archivo.
    """
    titulos_nivel_1 = [
        "Información general del proyecto",
        "Introducción",
        "Planteamiento del problema",
        "Objetivos",
        "5. Estado del arte y estado de la técnica",
        "6. Metodología de desarrollo",
        "7. Desarrollo del proyecto",
        "8. Resultados obtenidos",
        "9. Análisis de viabilidad",
        "10. Propiedad intelectual y transferencia tecnológica",
        "11. Impacto del proyecto",
        "12. Conclusiones",
        "13. Referencias bibliográficas",
        "14. Anexos",
    ]
    titulos_nivel_2 = [
        "4.1 Objetivo General",
        "4.2 Objetivos Específicos",
    ]

    for titulo in titulos_nivel_1:
        buscar_parrafo(
            documento,
            titulo,
            coincidencia_exacta=True,
        ).style = "Heading 1"

    for titulo in titulos_nivel_2:
        buscar_parrafo(
            documento,
            titulo,
            coincidencia_exacta=True,
        ).style = "Heading 2"

    # La plantilla ya contiene un campo TOC. Se limita a niveles 1 y 2,
    # se desbloquea y se marca como pendiente de actualización.
    for instr in documento._element.xpath(
        './/*[local-name()="instrText"]'
    ):
        if "TOC" in (instr.text or ""):
            instr.text = r' TOC \o "1-2" \h \z \u '

    # Elimina marcadores internos obsoletos del índice para que Word o
    # LibreOffice los reconstruyan según la ubicación actual de cada título.
    ids_toc: set[str] = set()

    for marcador in documento._element.xpath(
        './/*[local-name()="bookmarkStart"]'
    ):
        nombre = marcador.get(qn("w:name"), "")
        if nombre.startswith("_Toc"):
            identificador = marcador.get(qn("w:id"))
            if identificador is not None:
                ids_toc.add(identificador)
            marcador.getparent().remove(marcador)

    for marcador in documento._element.xpath(
        './/*[local-name()="bookmarkEnd"]'
    ):
        if marcador.get(qn("w:id")) in ids_toc:
            marcador.getparent().remove(marcador)

    for campo in documento._element.xpath(
        './/*[local-name()="fldChar"]'
    ):
        campo.attrib.pop(qn("w:fldLock"), None)
        if campo.get(qn("w:fldCharType")) == "begin":
            campo.set(qn("w:dirty"), "true")

    marcar_actualizacion_campos(documento)


def insertar_tabla_resultados(
    documento: Document,
    parrafo_ancla: Paragraph,
    actividades: list[str],
) -> None:
    actividades_limpias = [
        limpiar_texto(str(actividad))
        for actividad in actividades
        if limpiar_texto(str(actividad))
    ]

    if not actividades_limpias:
        actividades_limpias = [
            "No se registraron actividades ejecutadas para relacionar."
        ]

    tabla = documento.add_table(
        rows=1,
        cols=3,
    )
    tabla.style = "Table Grid"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False

    anchos = [Cm(1.2), Cm(11.3), Cm(4.6)]
    encabezados = [
        "N.°",
        "Descripción de la actividad ejecutada",
        "Evidencia",
    ]

    for indice, encabezado in enumerate(encabezados):
        celda = tabla.rows[0].cells[indice]
        celda.width = anchos[indice]
        celda.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        escribir_celda(celda, encabezado)
        for run in celda.paragraphs[0].runs:
            run.bold = True
        celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for numero, actividad in enumerate(actividades_limpias, start=1):
        fila = tabla.add_row()
        valores = [
            str(numero),
            actividad,
            "Agregar enlace: ______________________________",
        ]

        for indice, valor in enumerate(valores):
            celda = fila.cells[indice]
            celda.width = anchos[indice]
            celda.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            escribir_celda(celda, valor)

        fila.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    parrafo_ancla._p.addnext(tabla._tbl)


def serializar_datos_informe(datos: dict) -> dict:
    resultado = deepcopy(datos)

    fecha_entrega = resultado.get("fecha_entrega")
    if isinstance(fecha_entrega, date):
        resultado["fecha_entrega"] = fecha_entrega.strftime("%d/%m/%Y")

    return resultado


# =====================================================
# GENERACIÓN DE CONTENIDO
# =====================================================

def inferir_metodologia_base(datos: dict) -> str:
    """Infiere una metodología coherente sin solicitarla en el formulario."""
    texto_base = " ".join(
        [
            datos.get("descripcion_general_proyecto", ""),
            datos.get("entregables_proyecto_base", ""),
            datos.get("innovacion_proyecto_base", ""),
            actividades_en_texto(
                datos.get("actividades_ejecutadas_base", [])
            ),
        ]
    ).casefold()

    if any(
        termino in texto_base
        for termino in [
            "inteligencia artificial",
            "machine learning",
            "modelo predictivo",
            "analítica de datos",
            "clasificación",
            "dataset",
        ]
    ):
        return (
            "CRISP-DM, complementada con desarrollo iterativo y validación "
            "experimental"
        )

    if any(
        termino in texto_base
        for termino in [
            "software",
            "aplicación",
            "plataforma",
            "sistema web",
            "app",
            "interfaz",
        ]
    ):
        return (
            "Diseño Centrado en el Usuario, Design Thinking y desarrollo "
            "ágil iterativo"
        )

    if any(
        termino in texto_base
        for termino in [
            "electrónica",
            "sensor",
            "microcontrolador",
            "iot",
            "automatización",
            "circuito",
        ]
    ):
        return (
            "ingeniería de sistemas mediante modelo V, complementada con "
            "prototipado iterativo"
        )

    if any(
        termino in texto_base
        for termino in [
            "impresión 3d",
            "manufactura",
            "mecanismo",
            "carcasa",
            "pieza",
            "producto físico",
            "diseño industrial",
        ]
    ):
        return (
            "Design Thinking, desarrollo iterativo de prototipos y Diseño "
            "para Manufactura y Ensamble (DFMA)"
        )

    if any(
        termino in texto_base
        for termino in [
            "marca",
            "identidad visual",
            "experiencia de usuario",
            "servicio",
            "comunicación visual",
        ]
    ):
        return "Doble Diamante y Diseño Centrado en el Usuario"

    if any(
        termino in texto_base
        for termino in [
            "agroindustrial",
            "alimento",
            "formulación",
            "proceso productivo",
            "laboratorio",
        ]
    ):
        return (
            "investigación aplicada, validación experimental y mejora "
            "continua mediante DMAIC"
        )

    return "Design Thinking y desarrollo iterativo de prototipos"


def metodologias_en_texto(datos: dict) -> str:
    seleccionadas = datos.get("metodologias_seleccionadas", []) or []
    otra = limpiar_texto(datos.get("otra_metodologia", ""))

    metodologias = [
        limpiar_texto(str(item))
        for item in seleccionadas
        if limpiar_texto(str(item)) and str(item) != "Otra"
    ]

    if otra:
        metodologias.append(otra)

    if metodologias:
        return ", ".join(metodologias)

    metodologia_respaldo = limpiar_texto(
        datos.get("metodologia_inferida", "")
    )
    return metodologia_respaldo or inferir_metodologia_base(datos)


def actividades_en_texto(actividades: list[str]) -> str:
    actividades_limpias = [
        limpiar_texto(str(actividad))
        for actividad in actividades
        if limpiar_texto(str(actividad))
    ]

    return "\n".join(
        f"{indice}. {actividad}"
        for indice, actividad in enumerate(actividades_limpias, start=1)
    )


def corregir_actividad_basica(texto: str) -> str:
    actividad = limpiar_texto(texto)

    if not actividad:
        return ""

    actividad = actividad[0].upper() + actividad[1:]

    if actividad[-1] not in ".!?":
        actividad += "."

    return actividad


def obtener_api_key() -> str | None:
    api_key = os.getenv("OPENAI_API_KEY")

    if api_key:
        return api_key.strip()

    try:
        valor = st.secrets.get("OPENAI_API_KEY")
        return str(valor).strip() if valor else None
    except Exception:
        return None


def extraer_json_respuesta(texto: str) -> dict:
    contenido = str(texto or "").strip()
    contenido = re.sub(r"^```(?:json)?\s*", "", contenido, flags=re.I)
    contenido = re.sub(r"\s*```$", "", contenido)

    try:
        resultado = json.loads(contenido)
        return resultado if isinstance(resultado, dict) else {}
    except json.JSONDecodeError:
        coincidencia = re.search(r"\{.*\}", contenido, flags=re.S)

        if not coincidencia:
            return {}

        try:
            resultado = json.loads(coincidencia.group(0))
            return resultado if isinstance(resultado, dict) else {}
        except json.JSONDecodeError:
            return {}


def referentes_modo_prueba(datos: dict) -> list[dict]:
    """
    Referentes reales y generales para validar la estructura en modo prueba.
    En modo API se reemplazan por soluciones específicamente relacionadas.
    """
    descripcion = datos.get("descripcion_general_proyecto", "").casefold()

    if any(
        palabra in descripcion
        for palabra in ["inteligencia artificial", "modelo", "clasificación", "machine learning"]
    ):
        return [
            {
                "nombre": "Teachable Machine",
                "entidad": "Google Creative Lab",
                "anio": "s. f.",
                "descripcion": (
                    "Herramienta web que permite crear modelos de aprendizaje automático "
                    "con imágenes, sonidos y posturas sin requerir programación avanzada."
                ),
                "cita_corta": "(Google Creative Lab, s. f.)",
                "referencia_apa": (
                    "Google Creative Lab. (s. f.). Teachable Machine. "
                    "https://teachablemachine.withgoogle.com/"
                ),
                "url": "https://teachablemachine.withgoogle.com/",
            },
            {
                "nombre": "TensorFlow Lite",
                "entidad": "Google",
                "anio": "s. f.",
                "descripcion": (
                    "Conjunto de herramientas para ejecutar modelos de aprendizaje "
                    "automático en dispositivos móviles, embebidos y de borde."
                ),
                "cita_corta": "(Google, s. f.)",
                "referencia_apa": (
                    "Google. (s. f.). TensorFlow Lite. "
                    "https://www.tensorflow.org/lite"
                ),
                "url": "https://www.tensorflow.org/lite",
            },
        ]

    return [
        {
            "nombre": "MIT App Inventor",
            "entidad": "Massachusetts Institute of Technology",
            "anio": "s. f.",
            "descripcion": (
                "Entorno de desarrollo visual orientado a la creación de aplicaciones "
                "móviles mediante bloques y procesos iterativos de prototipado."
            ),
            "cita_corta": "(MIT App Inventor, s. f.)",
            "referencia_apa": (
                "MIT App Inventor. (s. f.). MIT App Inventor. "
                "https://appinventor.mit.edu/"
            ),
            "url": "https://appinventor.mit.edu/",
        },
        {
            "nombre": "Arduino",
            "entidad": "Arduino",
            "anio": "s. f.",
            "descripcion": (
                "Plataforma abierta de hardware y software utilizada para construir "
                "prototipos electrónicos interactivos y soluciones conectadas."
            ),
            "cita_corta": "(Arduino, s. f.)",
            "referencia_apa": (
                "Arduino. (s. f.). Arduino documentation. "
                "https://docs.arduino.cc/"
            ),
            "url": "https://docs.arduino.cc/",
        },
    ]


def investigar_referentes_reales(
    datos: dict,
    modelo_openai: str,
) -> list[dict]:
    """
    Busca exactamente dos proyectos o soluciones reales y devuelve referencias
    verificables. La búsqueda usa la herramienta web de la Responses API.
    """
    if OpenAI is None:
        raise RuntimeError(
            "La librería openai no está instalada y no es posible verificar "
            "los dos referentes reales del Estado del Arte."
        )

    api_key = obtener_api_key()

    if not api_key:
        raise RuntimeError(
            "No se encontró OPENAI_API_KEY para investigar los referentes reales."
        )

    client = OpenAI(api_key=api_key)

    prompt = f"""
Busca en fuentes oficiales, institucionales o académicas exactamente dos proyectos,
productos o soluciones reales comparables con el siguiente proyecto de base tecnológica.

Proyecto: {datos.get('nombre_proyecto', '')}
Descripción: {datos.get('descripcion_general_proyecto', '')}
Línea tecnológica: {datos.get('linea_tecnologica', '')}

Selecciona referentes realmente relacionados con el propósito, la tecnología o el
problema atendido. No inventes nombres, entidades, años, enlaces ni publicaciones.
Prefiere la página oficial del proyecto o una publicación académica primaria.

Responde exclusivamente con JSON válido:
{{
  "proyectos": [
    {{
      "nombre": "nombre oficial",
      "entidad": "entidad responsable",
      "anio": "año o s. f.",
      "descripcion": "síntesis verificable de 50 a 90 palabras",
      "cita_corta": "(Entidad, año)",
      "referencia_apa": "referencia completa en APA 7 con URL",
      "url": "https://..."
    }}
  ]
}}
"""

    respuesta = client.responses.create(
        model=modelo_openai,
        tools=[{"type": "web_search"}],
        input=prompt,
        temperature=0.1,
    )

    datos_respuesta = extraer_json_respuesta(
        getattr(respuesta, "output_text", "")
    )
    proyectos = datos_respuesta.get("proyectos", [])

    if not isinstance(proyectos, list):
        proyectos = []

    proyectos_validos: list[dict] = []

    for item in proyectos:
        if not isinstance(item, dict):
            continue

        campos = {
            clave: limpiar_texto(str(item.get(clave, "")))
            for clave in [
                "nombre",
                "entidad",
                "anio",
                "descripcion",
                "cita_corta",
                "referencia_apa",
                "url",
            ]
        }

        if (
            campos["nombre"]
            and campos["entidad"]
            and campos["descripcion"]
            and campos["referencia_apa"]
            and campos["url"].startswith(("http://", "https://"))
        ):
            proyectos_validos.append(campos)

    if len(proyectos_validos) < 2:
        raise RuntimeError(
            "La búsqueda no produjo dos referentes reales con fuentes verificables. "
            "Intenta nuevamente o revisa la descripción general del proyecto."
        )

    return proyectos_validos[:2]


def referencias_bibliograficas_proyecto(datos: dict) -> list[str]:
    referencias = [
        limpiar_texto(item.get("referencia_apa", ""))
        for item in datos.get("referentes_estado_arte", [])
        if limpiar_texto(item.get("referencia_apa", ""))
    ]

    anio = datos.get("fecha_entrega", date.today()).year
    referencias.append(
        (
            f"Red Tecnoparque SENA. ({anio}). Estado del Arte del proyecto "
            f"{datos.get('nombre_proyecto', '')}. Documento de planeación del proyecto."
        )
    )

    return referencias


def anexos_manuales() -> list[str]:
    return [
        "Espacio para insertar manualmente enlaces de evidencias y repositorios.",
        "Espacio para insertar manualmente imágenes, diagramas, planos o capturas.",
        "Espacio para relacionar manualmente otros soportes técnicos del proyecto.",
    ]


def recomendacion_continuidad_trl(trl_alcanzado: str) -> str:
    coincidencia = re.search(r"(\d+)", str(trl_alcanzado or ""))

    if not coincidencia:
        return (
            "Se recomienda definir una fase posterior de validación y continuidad "
            "de acuerdo con el nivel de madurez tecnológica comprobado."
        )

    nivel = int(coincidencia.group(1))

    if 1 <= nivel < 9:
        return (
            f"Al finalizar en TRL {nivel}, es viable formular un nuevo proyecto "
            f"orientado a avanzar hacia TRL {nivel + 1}, sin afirmar que dicho "
            "nivel haya sido alcanzado en la ejecución actual."
        )

    return (
        "Al finalizar en TRL 9, se recomienda concentrar la continuidad en adopción, "
        "escalamiento, sostenibilidad, transferencia y seguimiento del desempeño."
    )


def propiedad_intelectual_modo_prueba(datos: dict) -> str:
    texto_base = " ".join(
        [
            datos.get("descripcion_general_proyecto", ""),
            datos.get("entregables_proyecto_base", ""),
            datos.get("innovacion_proyecto_base", ""),
            actividades_en_texto(
                datos.get("actividades_ejecutadas_base", [])
            ),
        ]
    ).casefold()

    mecanismos: list[str] = []

    if any(
        palabra in texto_base
        for palabra in [
            "software", "aplicación", "plataforma", "código",
            "algoritmo", "base de datos", "sistema web",
        ]
    ):
        mecanismos.append(
            "registro de software ante la Dirección Nacional de Derecho de "
            "Autor, acompañado por la protección del código, la documentación "
            "y los contenidos originales mediante derecho de autor"
        )

    if any(
        palabra in texto_base
        for palabra in [
            "mejora funcional", "mecanismo", "dispositivo", "equipo",
            "prototipo físico", "máquina",
        ]
    ):
        mecanismos.append(
            "modelo de utilidad ante la Superintendencia de Industria y "
            "Comercio, siempre que la configuración funcional cumpla los "
            "requisitos de novedad y aplicación industrial"
        )

    if any(
        palabra in texto_base
        for palabra in [
            "invención", "nuevo principio", "solución técnica nueva",
            "procedimiento novedoso",
        ]
    ):
        mecanismos.append(
            "patente de invención ante la Superintendencia de Industria y "
            "Comercio, condicionada a una búsqueda de antecedentes y al "
            "cumplimiento de novedad, nivel inventivo y aplicación industrial"
        )

    if any(
        palabra in texto_base
        for palabra in [
            "apariencia", "forma externa", "carcasa", "diseño industrial",
            "configuración estética",
        ]
    ):
        mecanismos.append(
            "registro de diseño industrial ante la Superintendencia de "
            "Industria y Comercio para la apariencia externa original"
        )

    if any(
        palabra in texto_base
        for palabra in [
            "marca", "logo", "identidad visual", "nombre comercial",
            "signo distintivo",
        ]
    ):
        mecanismos.append(
            "registro de marca ante la Superintendencia de Industria y "
            "Comercio para los signos distintivos efectivamente utilizados"
        )

    if any(
        palabra in texto_base
        for palabra in [
            "fórmula", "receta", "know-how", "proceso confidencial",
            "secreto", "parámetro reservado",
        ]
    ):
        mecanismos.append(
            "secreto empresarial para la información técnica o comercial que "
            "se mantenga reservada mediante controles y acuerdos de confidencialidad"
        )

    if not mecanismos:
        mecanismos.append(
            "derecho de autor sobre la documentación técnica, planos, textos, "
            "gráficos y demás obras originales desarrolladas"
        )

    mecanismos = mecanismos[:2]

    return (
        "El análisis automático de la naturaleza de la solución, sus entregables, "
        "la innovación reportada y las actividades ejecutadas permite identificar "
        "como mecanismos pertinentes "
        + "; y ".join(mecanismos)
        + ". La selección definitiva requiere verificar la titularidad, la novedad, "
        "el nivel de divulgación y los antecedentes aplicables antes de presentar "
        "una solicitud. Para la transferencia tecnológica se recomienda organizar "
        "los activos, conservar evidencias de autoría, definir condiciones de uso y "
        "establecer acuerdos de licencia, colaboración o adopción con los aliados "
        "que participen en una fase posterior."
    )


def _complemento_apartado(clave: str, datos: dict) -> str:
    metodologia = metodologias_en_texto(datos)
    entregables = datos.get("entregables_proyecto_base", "")
    innovacion = datos.get("innovacion_proyecto_base", "")
    impacto = datos.get("impacto_proyecto_base", "")
    actividades = actividades_en_texto(
        datos.get("actividades_ejecutadas_base", [])
    )

    complementos = {
        "introduccion": (
            "El alcance del informe se limita a la experiencia efectivamente "
            "desarrollada y al nivel de madurez tecnológica declarado. La lectura "
            "del documento debe permitir reconocer el propósito, los actores, el "
            "entorno de aplicación y la relación entre la necesidad y la solución, "
            "sin anticipar detalles que corresponden a resultados, viabilidad o "
            "propiedad intelectual. Esta delimitación facilita una presentación "
            "ordenada del cierre y mantiene la trazabilidad de la información."
        ),
        "planteamiento_problema": (
            "La formulación del problema debe diferenciar causas, efectos y usuarios "
            "afectados, además de precisar la brecha que justificó la intervención. "
            "El análisis no convierte la solución en el centro del apartado; se "
            "concentra en las condiciones iniciales, las restricciones observadas y "
            "las consecuencias de mantener la situación sin una respuesta técnica "
            "adaptada al contexto."
        ),
        "estado_arte_tecnica": (
            "La comparación tecnológica considera el propósito, la arquitectura, la "
            "madurez, la accesibilidad y la capacidad de adaptación de las soluciones "
            "existentes. Los referentes se utilizan para identificar tendencias y "
            "diferencias, no para afirmar equivalencias completas. El aporte innovador "
            f"del proyecto se relaciona con {innovacion}, de acuerdo con la información "
            "suministrada y sin atribuir características que no hayan sido verificadas."
        ),
        "metodologia_desarrollo": (
            f"La metodología seleccionada fue {metodologia}. Su aplicación se interpreta "
            "como una secuencia de comprensión, definición, diseño, implementación, "
            "revisión y ajuste. La trazabilidad entre decisiones y actividades permite "
            "explicar cómo se transformaron los requerimientos en componentes y cómo "
            "se controlaron los cambios durante el desarrollo."
        ),
        "desarrollo_proyecto": (
            f"Las actividades reportadas fueron: {actividades} La descripción del "
            "desarrollo las organiza en una secuencia técnica, identifica relaciones "
            "entre ellas y explica su contribución a la construcción de la solución. "
            "Los entregables se presentan como productos del proceso y no como nuevas "
            "actividades, lo que evita duplicaciones y conserva claridad documental."
        ),
        "resultados_obtenidos": (
            f"Los entregables informados fueron: {entregables} La valoración de los "
            "resultados se realiza por su correspondencia con los objetivos, su función "
            "dentro de la solución y su aporte al nivel TRL alcanzado. La tabla de "
            "actividades deja disponibles espacios de evidencia para incorporar enlaces "
            "en Word sin exigir archivos durante el diligenciamiento."
        ),
        "analisis_viabilidad": (
            "La viabilidad se examina desde condiciones técnicas, operativas, económicas, "
            "normativas y de adopción. El análisis distingue capacidades disponibles de "
            "requisitos futuros, identifica dependencias y reconoce limitaciones que "
            "deben resolverse antes de un escalamiento. También considera mantenimiento, "
            "documentación, soporte, alianzas y recursos como factores de continuidad."
        ),
        "propiedad_transferencia": (
            "La recomendación de protección no equivale a un derecho concedido. Antes de "
            "tramitar cualquier mecanismo deben revisarse antecedentes, titularidad, "
            "divulgaciones previas y acuerdos entre participantes. La transferencia puede "
            "estructurarse mediante licencias, colaboración, cesión o adopción, según el "
            "activo consolidado y la estrategia de continuidad del proyecto."
        ),
        "impacto_proyecto": (
            f"El impacto reportado fue: {impacto} Su análisis debe identificar quiénes "
            "reciben el beneficio, mediante qué cambio concreto y bajo cuáles condiciones "
            "puede sostenerse. Se diferencian efectos observados de expectativas futuras y "
            "se evita presentar cifras o resultados que no hayan sido suministrados. La "
            "continuidad del impacto depende de adopción, soporte y seguimiento."
        ),
        "conclusiones": (
            "Las conclusiones articulan pertinencia, aprendizaje, entregables, innovación "
            "y limitaciones sin introducir hechos nuevos. El nivel TRL se utiliza como "
            "referencia para orientar una fase posterior, manteniendo explícito que el "
            "siguiente nivel constituye una meta de continuidad y no un resultado ya "
            "alcanzado dentro del proyecto que se cierra."
        ),
    }

    return complementos.get(clave, "")


def _ampliacion_adicional(clave: str, datos: dict) -> str:
    ampliaciones = {
        "introduccion": (
            "El cierre documental también permite ubicar el proyecto dentro del proceso "
            "de acompañamiento de TecnoParque, reconocer la participación del talento y "
            "del experto, y establecer el periodo al que corresponden los resultados. "
            "La introducción delimita el objeto del informe y aclara que las valoraciones "
            "posteriores se basan en la información efectivamente suministrada. De esta "
            "forma, el lector obtiene una visión inicial suficiente para comprender el "
            "sentido del proyecto, su relación con el entorno y la razón por la cual se "
            "documenta el proceso de desarrollo. La sección evita convertir el contexto "
            "en una repetición de actividades o productos, y reserva para cada numeral "
            "los elementos que le corresponden. Esta organización fortalece la coherencia "
            "del documento y facilita la revisión de la trazabilidad entre necesidad, "
            "propósito, ejecución, resultados e impacto."
        ),
        "planteamiento_problema": (
            "La situación inicial debe comprenderse a partir de las prácticas existentes, "
            "las limitaciones de acceso, desempeño, integración o disponibilidad y las "
            "dificultades que enfrentaban los usuarios. El apartado identifica por qué "
            "esas condiciones afectaban el proceso y qué consecuencias podían mantenerse "
            "si no se intervenían. También diferencia síntomas de causas para evitar una "
            "formulación superficial. La justificación tecnológica surge cuando la brecha "
            "no puede resolverse únicamente con ajustes menores y requiere diseño, "
            "integración o validación de una alternativa. La redacción conserva una relación "
            "directa con el contexto reportado, pero no atribuye magnitudes, frecuencias ni "
            "pérdidas que no hayan sido documentadas. Así, el problema funciona como base "
            "para valorar la pertinencia de los objetivos y la correspondencia de los "
            "resultados obtenidos."
        ),
        "estado_arte_tecnica": (
            "El análisis de referentes debe valorar el grado de madurez de las alternativas, "
            "las tecnologías empleadas, el tipo de usuario atendido y las condiciones de "
            "operación. La comparación no se limita a enumerar productos, sino que identifica "
            "tendencias de diseño, integración, automatización, interoperabilidad, usabilidad "
            "y sostenibilidad. Los dos casos reales se utilizan como evidencia de avances "
            "existentes y como punto de contraste para reconocer el aporte diferencial de "
            "la solución desarrollada. También se consideran limitaciones de adaptación, "
            "costos de adopción, dependencia de infraestructura y posibilidades de mejora, "
            "siempre que puedan deducirse razonablemente de las fuentes consultadas. Esta "
            "síntesis orienta la lectura técnica del informe, mientras la revisión extensa, "
            "las matrices comparativas y las demás referencias permanecen en el documento "
            "de planeación correspondiente."
        ),
        "metodologia_desarrollo": (
            "La metodología se evidencia en la manera de organizar decisiones, priorizar "
            "requerimientos y transformar observaciones en ajustes concretos. El proceso "
            "puede integrar exploración con usuarios, definición de criterios, construcción "
            "progresiva, integración de componentes y revisión del funcionamiento. Cada "
            "etapa conserva relación con una actividad reportada y con un resultado esperado, "
            "lo que permite explicar la evolución sin inventar ceremonias, roles o documentos "
            "que no hayan existido. La iteración se interpreta como un mecanismo para reducir "
            "incertidumbre, detectar incompatibilidades y mejorar la solución antes del cierre. "
            "La validación se presenta de acuerdo con las evidencias disponibles y el TRL, "
            "diferenciando pruebas realizadas de evaluaciones que deberán completarse en una "
            "fase posterior. Así, el enfoque metodológico ofrece trazabilidad y no una simple "
            "lista de conceptos."
        ),
        "desarrollo_proyecto": (
            "La narración del desarrollo explica la secuencia temporal y técnica de las "
            "acciones, señalando cómo los insumos de una actividad permitieron ejecutar la "
            "siguiente. Se describen los momentos de análisis, diseño, preparación, construcción, "
            "configuración, integración y revisión que resulten aplicables. Las decisiones "
            "se presentan por su efecto sobre el funcionamiento, la compatibilidad o la "
            "adaptación al contexto, sin atribuir especificaciones no reportadas. Cuando se "
            "produjeron ajustes, se explica su relación con hallazgos del proceso y con la "
            "necesidad de mejorar el prototipo. La sección diferencia claramente las acciones "
            "ejecutadas de los productos obtenidos y del impacto posterior. Esta estructura "
            "facilita reconocer la contribución de cada actividad y deja una base coherente "
            "para asociar evidencias en la tabla de resultados."
        ),
        "resultados_obtenidos": (
            "Los resultados se describen como productos verificables del proceso y no como "
            "expectativas. El análisis establece qué función cumple cada entregable, cómo se "
            "relaciona con los objetivos y qué parte de la necesidad permite atender. También "
            "reconoce el aporte de la innovación cuando esta se manifiesta en una mejora, una "
            "integración particular o una adaptación al entorno. La valoración respeta el TRL "
            "alcanzado y evita presentar como definitivo aquello que todavía requiere pruebas "
            "en condiciones más amplias. La tabla complementaria organiza las actividades y "
            "deja un espacio para enlaces de evidencia, de modo que fotografías, repositorios, "
            "actas, manuales o demostraciones puedan incorporarse posteriormente. Esta separación "
            "entre síntesis y soportes mantiene el informe legible y facilita su verificación."
        ),
        "analisis_viabilidad": (
            "La viabilidad técnica considera disponibilidad, compatibilidad, mantenibilidad y "
            "posibilidad de reproducir o escalar la solución. La dimensión operativa examina "
            "responsables, capacidades, rutinas y condiciones de uso. El componente económico "
            "identifica recursos que deberán estimarse sin inventar valores, mientras el análisis "
            "normativo reconoce obligaciones aplicables a datos, seguridad, propiedad intelectual, "
            "salud, ambiente o calidad, según la naturaleza del proyecto. La adopción depende de "
            "que los usuarios comprendan el funcionamiento y perciban beneficios suficientes. "
            "También se valoran sostenibilidad, soporte, disponibilidad de insumos, alianzas y "
            "oportunidades de mercado cuando sean pertinentes. El apartado concluye diferenciando "
            "condiciones favorables, limitaciones y acciones necesarias antes de una implementación "
            "más amplia."
        ),
        "propiedad_transferencia": (
            "El mecanismo sugerido debe corresponder a la naturaleza del activo y no a una lista "
            "general de posibilidades. Los componentes originales se analizan para distinguir "
            "obras protegibles por derecho de autor, desarrollos funcionales con potencial de "
            "propiedad industrial, signos distintivos y conocimiento reservado. La documentación "
            "de fechas, autores, versiones y aportes resulta esencial para sustentar titularidad. "
            "También deben evitarse divulgaciones que puedan afectar la novedad cuando se estudie "
            "una patente, un modelo de utilidad o un diseño industrial. La estrategia de transferencia "
            "puede involucrar licenciamiento, colaboración, adopción institucional, prestación de "
            "servicios o continuidad conjunta, siempre con reglas claras sobre uso, mantenimiento, "
            "confidencialidad y distribución de beneficios. La recomendación final permanece sujeta "
            "a una revisión jurídica y técnica específica."
        ),
        "impacto_proyecto": (
            "El impacto se interpreta a partir del cambio que la solución puede producir en usuarios, "
            "procesos u organizaciones. La dimensión tecnológica se relaciona con capacidades, acceso, "
            "integración o apropiación; la productiva, con mejoras en operación o toma de decisiones; "
            "la social, con beneficios para personas o comunidades; la económica, con oportunidades de "
            "eficiencia o generación de valor; y la ambiental, con reducción o mejor gestión de recursos, "
            "cuando aplique. La redacción distingue beneficios observados de efectos potenciales y evita "
            "cuantificaciones no respaldadas. También identifica condiciones para sostener el impacto, "
            "como capacitación, soporte, mantenimiento, actualización, participación de beneficiarios y "
            "seguimiento. Esta perspectiva permite valorar el aporte sin confundirlo con la simple entrega "
            "de un prototipo."
        ),
        "conclusiones": (
            "El cierre reconoce qué aspectos quedaron consolidados, cuáles requieren validación adicional "
            "y qué aprendizajes deben conservarse. La conclusión no repite el desarrollo completo; relaciona "
            "la necesidad, los objetivos, los entregables y el impacto para valorar la coherencia global. "
            "También identifica limitaciones técnicas, operativas o documentales que pueden orientar una "
            "fase posterior. La recomendación de continuidad se formula de acuerdo con el TRL alcanzado y "
            "presenta el siguiente nivel como una meta condicionada a nuevas actividades, recursos y evidencias. "
            "Se consideran además oportunidades de protección, transferencia, adopción o escalamiento, sin "
            "afirmar que ya se hayan materializado. De esta manera, el informe finaliza con una valoración "
            "realista del estado actual y con una orientación concreta para la evolución responsable del proyecto."
        ),
    }
    return ampliaciones.get(clave, "")


def ajustar_rango_palabras(
    texto: str,
    clave: str,
    datos: dict,
    minimo: int = 300,
    maximo: int = 340,
) -> str:
    resultado = limpiar_texto(texto)

    for fragmento in [
        _complemento_apartado(clave, datos),
        _ampliacion_adicional(clave, datos),
    ]:
        if len(resultado.split()) >= minimo:
            break
        fragmento_limpio = limpiar_texto(fragmento)
        if fragmento_limpio and fragmento_limpio.casefold() not in resultado.casefold():
            resultado = f"{resultado} {fragmento_limpio}".strip()

    if len(resultado.split()) < minimo:
        cierres_finales = {
            "introduccion": (
                "Con esta delimitación, el apartado ofrece una entrada suficiente para comprender "
                "el documento, ubicar el proyecto en su contexto institucional y reconocer el alcance "
                "del cierre sin confundirlo con una propuesta futura ni con una descripción exhaustiva "
                "de los productos obtenidos."
            ),
            "planteamiento_problema": (
                "Esta formulación permite usar el problema como criterio de evaluación del proyecto: "
                "los objetivos deben responder a la brecha identificada y los resultados deben mostrar "
                "una contribución verificable a su atención, aun cuando permanezcan necesidades de "
                "validación o ampliación."
            ),
            "estado_arte_tecnica": (
                "La síntesis resultante ofrece un marco suficiente para interpretar la posición del "
                "proyecto frente a soluciones existentes y para comprender por qué su adaptación al "
                "entorno puede representar una mejora pertinente, sin sustituir la revisión académica "
                "y técnica completa realizada en planeación."
            ),
            "metodologia_desarrollo": (
                "La explicación metodológica conserva correspondencia con lo ejecutado y evita presentar "
                "el proyecto como si hubiera aplicado procedimientos no reportados. Su valor está en mostrar "
                "cómo se ordenó el trabajo, cómo se revisaron los avances y cómo las decisiones permitieron "
                "consolidar los entregables."
            ),
            "desarrollo_proyecto": (
                "El resultado es una descripción continua del proceso, útil para que terceros comprendan "
                "la evolución del proyecto, identifiquen los momentos principales y relacionen cada actividad "
                "con la solución final sin perder claridad entre ejecución, resultados e impacto."
            ),
            "resultados_obtenidos": (
                "La valoración final mantiene una relación directa con los productos reportados y deja claro "
                "qué fue alcanzado dentro del proyecto. Cualquier desempeño adicional, adopción masiva o "
                "escalamiento deberá demostrarse en fases posteriores mediante evidencias específicas."
            ),
            "analisis_viabilidad": (
                "Con base en estas dimensiones, la continuidad puede planearse mediante una evaluación gradual "
                "de recursos, riesgos y condiciones de adopción. La decisión de implementar o escalar debe "
                "apoyarse en pruebas adicionales y en una estimación formal de costos, responsabilidades y "
                "requisitos aplicables."
            ),
            "propiedad_transferencia": (
                "Esta orientación permite priorizar acciones realistas de protección y transferencia, reducir "
                "riesgos de pérdida de novedad o autoría y preparar una estrategia compatible con la etapa de "
                "madurez del proyecto y con los intereses de los participantes."
            ),
            "impacto_proyecto": (
                "La valoración del impacto debe actualizarse cuando existan nuevas evidencias de uso, cobertura "
                "o desempeño. Por ahora, el informe conserva los efectos reportados y presenta las condiciones "
                "que permitirían mantener o ampliar los beneficios en una fase posterior."
            ),
            "conclusiones": (
                "Esta lectura de cierre permite tomar decisiones de continuidad sin sobreestimar el resultado. "
                "El proyecto queda documentado en su nivel real de avance y con una ruta de trabajo que puede "
                "convertirse en una nueva formulación orientada a validaciones más exigentes."
            ),
        }
        resultado = f"{resultado} {cierres_finales.get(clave, '')}".strip()

    if len(resultado.split()) < minimo:
        refuerzos_finales = {
            "introduccion": (
                "La información institucional, el TRL y la fecha de entrega completan esta "
                "ubicación inicial y permiten interpretar el documento como un cierre correspondiente "
                "a una etapa concreta del proceso tecnológico. El apartado queda así concentrado en "
                "presentar el proyecto y orientar la lectura de los numerales siguientes."
            ),
            "planteamiento_problema": (
                "La delimitación obtenida evita formulaciones demasiado amplias y permite reconocer "
                "una necesidad susceptible de ser atendida mediante un proyecto de base tecnológica. "
                "Esta precisión es indispensable para comprobar posteriormente la relación entre el "
                "problema, los objetivos formulados y los productos efectivamente alcanzados."
            ),
            "estado_arte_tecnica": (
                "La revisión también ayuda a reconocer oportunidades de interoperabilidad, accesibilidad, "
                "mantenimiento y escalamiento que pueden orientar nuevas decisiones. La comparación se "
                "mantiene sustentada en fuentes verificables y evita convertir tendencias generales en "
                "resultados propios del proyecto."
            ),
            "metodologia_desarrollo": (
                "La secuencia metodológica facilita además documentar responsables, insumos, criterios de "
                "revisión y decisiones de ajuste. Esta trazabilidad resulta útil para reproducir el proceso, "
                "identificar aprendizajes y preparar una eventual fase de continuidad con actividades y "
                "evidencias mejor delimitadas."
            ),
            "desarrollo_proyecto": (
                "La articulación entre actividades evidencia una progresión desde la definición hasta la "
                "consolidación de los productos. El relato técnico conserva el orden del trabajo y permite "
                "comprender qué decisiones fueron necesarias para integrar componentes, resolver dificultades "
                "y preparar la solución para su valoración final."
            ),
            "resultados_obtenidos": (
                "La síntesis permite reconocer el alcance real de los productos sin confundirlos con impactos "
                "futuros. Los espacios de evidencia facilitan completar la trazabilidad documental después de "
                "la generación del archivo, conservando flexibilidad para incorporar enlaces, repositorios, "
                "actas, fotografías o registros de prueba."
            ),
            "analisis_viabilidad": (
                "La conclusión de viabilidad debe entenderse como una valoración técnica preliminar basada en "
                "el estado actual del proyecto. Una decisión de adopción o inversión requerirá complementar "
                "esta información con pruebas de desempeño, costos detallados, revisión regulatoria, análisis "
                "de usuarios y definición de responsabilidades operativas. También será necesario establecer "
                "un plan de mantenimiento, soporte, actualización y gestión de riesgos que permita sostener "
                "el funcionamiento en condiciones reales durante una etapa posterior."
            ),
            "propiedad_transferencia": (
                "La identificación temprana de los activos facilita organizar expedientes, acuerdos y evidencias "
                "antes de una negociación o divulgación. También permite seleccionar aliados y mecanismos de "
                "transferencia compatibles con la estrategia de desarrollo, evitando comprometer derechos o "
                "información reservada sin una evaluación previa."
            ),
            "impacto_proyecto": (
                "La lectura de impacto conserva una relación directa con los beneficiarios y con el problema "
                "atendido. Su seguimiento posterior permitirá confirmar permanencia, ampliar la cobertura y "
                "establecer indicadores, siempre que se definan mecanismos de recolección de evidencias y una "
                "responsabilidad clara sobre la continuidad de la solución."
            ),
            "conclusiones": (
                "El informe deja así una base para decidir si conviene fortalecer el prototipo, ampliar las "
                "pruebas, preparar una estrategia de adopción o formular un nuevo proyecto. La recomendación "
                "debe revisarse junto con los recursos disponibles, los aliados y las condiciones necesarias "
                "para demostrar el siguiente nivel de madurez."
            ),
        }
        resultado = f"{resultado} {refuerzos_finales.get(clave, '')}".strip()

    palabras = resultado.split()
    if len(palabras) > maximo:
        resultado = " ".join(palabras[:maximo]).rstrip(" ,;:")
        if resultado[-1:] not in ".!?":
            resultado += "."

    return resultado


def asegurar_nota_estado_arte(texto: str) -> str:
    nota = (
        "El Estado del Arte completo se encuentra en la sección de documentos "
        "de planeación del proyecto, específicamente en el documento Estado del Arte."
    )

    if "documento estado del arte" not in texto.casefold():
        return f"{texto.strip()}\n\n{nota}"

    return texto.strip()


def contenido_modo_prueba(datos: dict) -> dict:
    metodologia = metodologias_en_texto(datos)
    actividades_corregidas = [
        corregir_actividad_basica(actividad)
        for actividad in datos.get("actividades_ejecutadas_base", [])
        if corregir_actividad_basica(actividad)
    ]
    referentes = datos.get("referentes_estado_arte", [])

    referentes_texto = " ".join(
        (
            f"{item.get('nombre', '')}, desarrollado por "
            f"{item.get('entidad', '')}, constituye un referente porque "
            f"{item.get('descripcion', '')} {item.get('cita_corta', '')}."
        )
        for item in referentes
    )

    descripcion = datos.get("descripcion_general_proyecto", "")
    entregables = datos.get("entregables_proyecto_base", "")
    innovacion = datos.get("innovacion_proyecto_base", "")
    impacto = datos.get("impacto_proyecto_base", "")

    base = {
        "introduccion": (
            f"El presente informe documenta el cierre técnico del proyecto "
            f"{datos.get('nombre_proyecto', '')}, desarrollado en "
            f"{datos.get('tecnoparque', '')}. La iniciativa se contextualiza a "
            f"partir de la siguiente información: {descripcion} El apartado presenta "
            "el propósito general, los actores, el entorno y el alcance alcanzado, "
            "sin desarrollar de manera anticipada el problema, la metodología ni los "
            "resultados. La organización del informe permite conservar trazabilidad "
            "entre la necesidad, los objetivos, las actividades y el nivel de madurez "
            "tecnológica declarado."
        ),
        "planteamiento_problema": (
            f"La problemática se deriva de las condiciones descritas para el proyecto: "
            f"{descripcion} El análisis identifica la brecha existente, los usuarios o "
            "procesos afectados, las causas que limitaron una respuesta adecuada y las "
            "consecuencias de mantener la situación inicial. La justificación se centra "
            "en la necesidad de disponer de una alternativa tecnológica pertinente y "
            "adaptada al contexto, sin convertir este apartado en una descripción de la "
            "solución o de los productos finales."
        ),
        "objetivo_general": (
            "Desarrollar y validar una solución de base tecnológica que responda a la "
            "necesidad identificada mediante actividades estructuradas de diseño, "
            "implementación, evaluación y documentación."
        ),
        "objetivos_especificos": [
            "Caracterizar la necesidad, los usuarios, los requerimientos y las condiciones de aplicación de la solución.",
            "Diseñar los componentes y criterios técnicos necesarios para materializar la propuesta tecnológica.",
            "Implementar e integrar la solución mediante actividades organizadas de construcción, configuración y ajuste.",
            "Validar los entregables obtenidos y establecer oportunidades de mejora, continuidad y escalamiento.",
        ],
        "estado_arte_tecnica": asegurar_nota_estado_arte(
            "Los avances tecnológicos del área muestran una tendencia hacia soluciones "
            "más modulares, interoperables, accesibles y orientadas al usuario. "
            + referentes_texto
            + f" El aporte innovador informado para el proyecto corresponde a {innovacion}. "
            "La comparación permite identificar mejoras de adaptación, integración o "
            "aplicación frente a alternativas existentes, sin afirmar equivalencias ni "
            "atribuir ventajas que no hayan sido verificadas."
        ),
        "metodologia_desarrollo": (
            f"A partir de la naturaleza de la solución y de las actividades reportadas, "
            f"se infiere la aplicación de {metodologia}. El enfoque organiza el proceso "
            "en momentos de comprensión de la necesidad, definición de requerimientos, "
            "diseño, implementación, revisión y ajuste. La metodología se articula con "
            "las actividades ejecutadas y permite explicar la secuencia real del trabajo "
            "sin exigir al usuario una selección metodológica adicional."
        ),
        "actividades_corregidas": actividades_corregidas,
        "desarrollo_proyecto": (
            f"El desarrollo se organizó con base en {metodologia}. Las actividades "
            "ejecutadas fueron: "
            + " ".join(actividades_corregidas)
            + " La secuencia describe cómo se pasó del análisis y el diseño a la "
            "construcción, integración, revisión y ajuste de la solución. Cada actividad "
            "se relaciona con una decisión o avance concreto, mientras los entregables "
            "se presentan como productos del proceso y no como actividades adicionales."
        ),
        "resultados_obtenidos": (
            f"Los entregables obtenidos fueron: {entregables} Estos resultados se "
            "analizan por su relación con los objetivos, las actividades ejecutadas, "
            f"el aporte innovador —{innovacion}— y el TRL alcanzado. La tabla de "
            "actividades incorpora una columna abierta para que el usuario agregue "
            "posteriormente enlaces de evidencia directamente en Word."
        ),
        "analisis_viabilidad": (
            "La viabilidad se examina considerando la estabilidad técnica, la operación, "
            "los recursos necesarios, la adopción, el mantenimiento, las obligaciones "
            "normativas y las oportunidades de sostenibilidad. El análisis diferencia "
            "las capacidades comprobadas de los requisitos que deben abordarse en una "
            "fase posterior y evita inventar costos, permisos, ventas o comportamientos "
            "de mercado no suministrados."
        ),
        "propiedad_transferencia": propiedad_intelectual_modo_prueba(datos),
        "impacto_proyecto": (
            f"El impacto informado fue: {impacto} El apartado organiza los beneficios "
            "según las dimensiones tecnológicas, sociales, económicas, ambientales o "
            "productivas que realmente resulten aplicables. Se identifica a los "
            "beneficiarios, el cambio generado y las condiciones necesarias para "
            "sostenerlo, sin atribuir cifras o efectos no reportados."
        ),
        "conclusiones": (
            "El cierre integra la pertinencia de la solución, la relación entre objetivos, "
            "actividades y entregables, el aporte innovador, los aprendizajes y las "
            "limitaciones identificadas. "
            + recomendacion_continuidad_trl(
                datos.get("trl_alcanzado", "")
            )
        ),
        "referencias_bibliograficas": referencias_bibliograficas_proyecto(datos),
        "anexos": anexos_manuales(),
    }

    for clave in [
        "introduccion",
        "planteamiento_problema",
        "estado_arte_tecnica",
        "metodologia_desarrollo",
        "desarrollo_proyecto",
        "resultados_obtenidos",
        "analisis_viabilidad",
        "propiedad_transferencia",
        "impacto_proyecto",
        "conclusiones",
    ]:
        base[clave] = ajustar_rango_palabras(
            base[clave],
            clave,
            datos,
        )

    return base


def normalizar_contenido(
    contenido: object,
    respaldo: dict,
    datos: dict,
) -> dict:
    if not isinstance(contenido, dict):
        contenido = {}

    resultado: dict = {}

    for clave in CLAVES_CONTENIDO:
        valor = contenido.get(clave)

        if clave in {
            "objetivos_especificos",
            "actividades_corregidas",
            "referencias_bibliograficas",
            "anexos",
        }:
            if isinstance(valor, str):
                valor = dividir_lineas(valor)

            if not isinstance(valor, list):
                valor = []

            valor = [
                limpiar_texto(str(item))
                for item in valor
                if limpiar_texto(str(item))
            ]

            if not valor:
                valor = respaldo[clave]

            resultado[clave] = valor
        else:
            texto_valor = limpiar_texto(str(valor or ""))
            resultado[clave] = texto_valor or respaldo[clave]

    resultado["objetivos_especificos"] = (
        resultado["objetivos_especificos"]
        + respaldo["objetivos_especificos"]
    )[:4]

    cantidad_actividades = len(
        datos.get("actividades_ejecutadas_base", [])
    )
    resultado["actividades_corregidas"] = resultado[
        "actividades_corregidas"
    ][:cantidad_actividades]

    if len(resultado["actividades_corregidas"]) < cantidad_actividades:
        faltantes = datos.get("actividades_ejecutadas_base", [])[
            len(resultado["actividades_corregidas"]):
        ]
        resultado["actividades_corregidas"].extend(
            corregir_actividad_basica(item)
            for item in faltantes
        )

    resultado["referencias_bibliograficas"] = (
        referencias_bibliograficas_proyecto(datos)
    )
    resultado["anexos"] = anexos_manuales()

    for clave in [
        "introduccion",
        "planteamiento_problema",
        "estado_arte_tecnica",
        "metodologia_desarrollo",
        "desarrollo_proyecto",
        "resultados_obtenidos",
        "analisis_viabilidad",
        "propiedad_transferencia",
        "impacto_proyecto",
        "conclusiones",
    ]:
        resultado[clave] = ajustar_rango_palabras(
            resultado[clave],
            clave,
            datos,
        )

    resultado["estado_arte_tecnica"] = asegurar_nota_estado_arte(
        resultado["estado_arte_tecnica"]
    )

    continuidad = recomendacion_continuidad_trl(
        datos.get("trl_alcanzado", "")
    )
    if continuidad.casefold() not in resultado["conclusiones"].casefold():
        resultado["conclusiones"] = (
            f"{resultado['conclusiones'].rstrip()} {continuidad}"
        )

    return resultado


def generar_contenido_con_ia(
    datos: dict,
    modelo_openai: str,
) -> dict:
    respaldo = contenido_modo_prueba(datos)
    metodologia = metodologias_en_texto(datos)
    actividades_originales = datos.get(
        "actividades_ejecutadas_base",
        [],
    )
    referentes = datos.get("referentes_estado_arte", [])

    reglas_comunes = """
Actúa como redactor técnico senior de proyectos de base tecnológica de la Red
TecnoParque SENA. Redacta un Informe Final institucional en español formal,
coherente, preciso y verificable.

REGLAS OBLIGATORIAS
- Utiliza únicamente los datos suministrados y los referentes web verificados.
- No inventes cifras, costos, pruebas, certificaciones, normas, patentes,
  registros, clientes, ventas, resultados ni referencias.
- Evita repetir la descripción general del proyecto entre apartados.
- Cada sección debe cumplir únicamente el propósito definido en el formato.
- Cada apartado narrativo debe contener entre 300 y 340 palabras.
- No uses introducciones genéricas ni frases de relleno.
- No uses markdown.
- Responde exclusivamente en JSON válido.
"""

    contexto = f"""
DATOS GENERALES
Talento: {datos.get('nombre_talento', '')}
Proyecto: {datos.get('nombre_proyecto', '')}
Código: {datos.get('codigo_proyecto', '')}
Experto: {datos.get('nombre_experto', '')}
Línea tecnológica: {datos.get('linea_tecnologica', '')}
TRL inicial: {datos.get('trl_inicial', '')}
TRL alcanzado: {datos.get('trl_alcanzado', '')}
TecnoParque: {datos.get('tecnoparque', '')}

DESCRIPCIÓN GENERAL DEL PROYECTO
{datos.get('descripcion_general_proyecto', '')}

ENTREGABLES OBTENIDOS
{datos.get('entregables_proyecto_base', '')}

INNOVACIÓN DEL PROYECTO
{datos.get('innovacion_proyecto_base', '')}

ACTIVIDADES EJECUTADAS
{actividades_en_texto(actividades_originales)}

IMPACTO DEL PROYECTO
{datos.get('impacto_proyecto_base', '')}

METODOLOGÍA SELECCIONADA
{metodologia}

DOS REFERENTES REALES VERIFICADOS PARA EL ESTADO DEL ARTE
{json.dumps(referentes, ensure_ascii=False, indent=2)}
"""

    instrucciones_bloque_1 = reglas_comunes + """
Genera los apartados 2, 3, 4, 5 y 6 y corrige las actividades.

REQUISITOS ESPECÍFICOS
- Introducción: contexto, propósito, entorno, actores y alcance. No desarrolles
  el problema, la metodología, los resultados ni el impacto.
- Planteamiento del problema: necesidad, causas, consecuencias, usuarios
  afectados y justificación. No describas extensamente la solución.
- Objetivo general: una sola oración con verbo en infinitivo.
- Objetivos específicos: exactamente cuatro, cada uno con verbo en infinitivo.
- Estado del arte: resume avances del área e integra exactamente los dos
  referentes reales con sus citas cortas. Explica el aporte innovador y finaliza
  indicando que el Estado del Arte completo está en los documentos de
  planeación, específicamente en el documento Estado del Arte.
- Metodología: determina automáticamente el enfoque estandarizado más coherente
  a partir de la naturaleza del proyecto y las actividades. Puedes emplear
  Design Thinking, DCU, Doble Diamante, metodologías ágiles, Scrum, Kanban,
  CRISP-DM, modelo V, DFMA, desarrollo iterativo de prototipos, DMAIC o
  investigación aplicada, pero menciona solo la combinación pertinente.
- Actividades corregidas: conserva exactamente la cantidad y el sentido de las
  actividades suministradas. Corrige ortografía y mejora la redacción. Cada
  actividad debe ser breve, técnica y clara; no agregues entregables, estados,
  observaciones ni evidencias.
"""

    entrada_bloque_1 = contexto + """

ESTRUCTURA JSON OBLIGATORIA
{
  "introduccion": "300 a 340 palabras",
  "planteamiento_problema": "300 a 340 palabras",
  "objetivo_general": "una oración",
  "objetivos_especificos": ["objetivo 1", "objetivo 2", "objetivo 3", "objetivo 4"],
  "estado_arte_tecnica": "300 a 340 palabras con dos citas",
  "metodologia_desarrollo": "300 a 340 palabras",
  "actividades_corregidas": ["actividad corregida 1", "actividad corregida 2"]
}
"""

    bloque_1 = generar_json_openai(
        instrucciones=instrucciones_bloque_1,
        entrada=entrada_bloque_1,
        modelo=modelo_openai,
        temperature=0.15,
    )

    bloque_1_normalizado = normalizar_contenido(
        bloque_1,
        respaldo,
        datos,
    )
    actividades_corregidas = bloque_1_normalizado[
        "actividades_corregidas"
    ]

    instrucciones_bloque_2 = reglas_comunes + """
Genera los apartados 7, 8, 9, 10, 11 y 12.

REQUISITOS ESPECÍFICOS
- Desarrollo: organiza la ejecución y articula la metodología seleccionada con las
  actividades corregidas. No repitas la introducción.
- Resultados: utiliza los entregables como fuente principal y explica su relación
  con objetivos, actividades, innovación y TRL. No enumeres las actividades,
  porque aparecerán en una tabla independiente.
- Viabilidad: analiza solo los aspectos técnicos, operativos, económicos,
  normativos, de adopción, sostenibilidad y escalabilidad pertinentes.
- Propiedad intelectual y transferencia: determina automáticamente los
  mecanismos realmente pertinentes en Colombia mediante el análisis de la
  descripción, objetivos, innovación, actividades y entregables. Limítate a
  registro de software, derecho de autor, modelo de utilidad, patente de
  invención, diseño industrial, secreto empresarial o registro de marca. No
  afirmes que exista un derecho concedido.
- Impacto: desarrolla el impacto suministrado y solo las dimensiones aplicables,
  sin repetir resultados ni introducción.
- Conclusiones: integra pertinencia, entregables, innovación, limitaciones y
  aprendizajes. Incluye exactamente la recomendación TRL suministrada y no
  afirmes que el nivel siguiente ya fue alcanzado.
"""

    entrada_bloque_2 = contexto + f"""

ACTIVIDADES CORREGIDAS
{actividades_en_texto(actividades_corregidas)}

RECOMENDACIÓN OBLIGATORIA DE CONTINUIDAD TRL
{recomendacion_continuidad_trl(datos.get('trl_alcanzado', ''))}

ESTRUCTURA JSON OBLIGATORIA
{{
  "desarrollo_proyecto": "300 a 340 palabras",
  "resultados_obtenidos": "300 a 340 palabras",
  "analisis_viabilidad": "300 a 340 palabras",
  "propiedad_transferencia": "300 a 340 palabras",
  "impacto_proyecto": "300 a 340 palabras",
  "conclusiones": "300 a 340 palabras"
}}
"""

    bloque_2 = generar_json_openai(
        instrucciones=instrucciones_bloque_2,
        entrada=entrada_bloque_2,
        modelo=modelo_openai,
        temperature=0.15,
    )

    contenido: dict = {}
    if isinstance(bloque_1, dict):
        contenido.update(bloque_1)
    if isinstance(bloque_2, dict):
        contenido.update(bloque_2)

    contenido["actividades_corregidas"] = actividades_corregidas
    contenido["referencias_bibliograficas"] = (
        referencias_bibliograficas_proyecto(datos)
    )
    contenido["anexos"] = anexos_manuales()

    return normalizar_contenido(
        contenido,
        respaldo,
        datos,
    )


# =====================================================
# GENERACIÓN DEL DOCUMENTO OFICIAL
# =====================================================

def actualizar_toc_con_libreoffice(ruta_docx: Path) -> bool:
    """
    Actualiza la tabla de contenido en un proceso completamente aislado.

    LibreOffice y UNO se ejecutan fuera del proceso de Streamlit para evitar
    fallos nativos. El documento se guarda mediante ``store()``, se cierra de
    forma ordenada y luego se termina el escritorio de LibreOffice.
    """
    ejecutable = shutil.which("libreoffice") or shutil.which("soffice")

    if not ejecutable or not ruta_docx.exists():
        return False

    python_sistema = (
        "/usr/bin/python3"
        if Path("/usr/bin/python3").exists()
        else (shutil.which("python3") or sys.executable)
    )

    import socket

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as socket_prueba:
        socket_prueba.bind(("127.0.0.1", 0))
        puerto = socket_prueba.getsockname()[1]

    carpeta_temporal = Path(
        tempfile.mkdtemp(prefix="softdocutecno_toc_estable_")
    )
    perfil = carpeta_temporal / "perfil_libreoffice"
    script_actualizacion = carpeta_temporal / "actualizar_toc_estable.py"

    script_actualizacion.write_text(
        r"""
import subprocess
import sys
import time
from pathlib import Path

sys.path.append("/usr/lib/python3/dist-packages")

try:
    import uno
    from com.sun.star.beans import PropertyValue
except Exception:
    raise SystemExit(2)

ruta = Path(sys.argv[1]).resolve()
ejecutable = sys.argv[2]
puerto = int(sys.argv[3])
perfil = Path(sys.argv[4]).resolve()
perfil.mkdir(parents=True, exist_ok=True)


def propiedad(nombre, valor):
    item = PropertyValue()
    item.Name = nombre
    item.Value = valor
    return item


proceso = subprocess.Popen(
    [
        ejecutable,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--norestore",
        "--nolockcheck",
        f"-env:UserInstallation=file://{perfil.as_posix()}",
        (
            "--accept=socket,host=127.0.0.1,"
            f"port={puerto};urp;StarOffice.ComponentContext"
        ),
    ],
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)

documento = None
escritorio = None

try:
    contexto_local = uno.getComponentContext()
    resolvedor = contexto_local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        contexto_local,
    )

    contexto = None

    for _ in range(100):
        try:
            contexto = resolvedor.resolve(
                (
                    "uno:socket,host=127.0.0.1,"
                    f"port={puerto};urp;StarOffice.ComponentContext"
                )
            )
            break
        except Exception:
            time.sleep(0.2)

    if contexto is None:
        raise SystemExit(3)

    escritorio = contexto.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        contexto,
    )

    documento = escritorio.loadComponentFromURL(
        uno.systemPathToFileUrl(str(ruta)),
        "_blank",
        0,
        (
            propiedad("Hidden", True),
            propiedad("ReadOnly", False),
            propiedad("UpdateDocMode", 3),
        ),
    )

    if documento is None:
        raise SystemExit(4)

    try:
        documento.updateLinks()
    except Exception:
        pass

    try:
        documento.getTextFields().refresh()
    except Exception:
        pass

    try:
        documento.calculateAll()
    except Exception:
        pass

    indices = documento.getDocumentIndexes()

    for _ in range(2):
        for numero in range(indices.getCount()):
            indices.getByIndex(numero).update()

        try:
            documento.getTextFields().refresh()
        except Exception:
            pass

        try:
            documento.calculateAll()
        except Exception:
            pass

        time.sleep(0.5)

    documento.store()
    documento.close(True)
    documento = None

    try:
        escritorio.terminate()
    except Exception:
        pass

    escritorio = None

finally:
    if documento is not None:
        try:
            documento.close(True)
        except Exception:
            pass

    if escritorio is not None:
        try:
            escritorio.terminate()
        except Exception:
            pass

    try:
        proceso.wait(timeout=15)
    except Exception:
        proceso.terminate()

        try:
            proceso.wait(timeout=5)
        except Exception:
            proceso.kill()
""",
        encoding="utf-8",
    )

    entorno = os.environ.copy()
    entorno.update(
        {
            "HOME": str(carpeta_temporal),
            "TMPDIR": str(carpeta_temporal),
            "SAL_USE_VCLPLUGIN": "svp",
            "QT_QPA_PLATFORM": "offscreen",
            "OMP_NUM_THREADS": "1",
            "MALLOC_ARENA_MAX": "2",
        }
    )

    proceso_worker = None

    try:
        proceso_worker = subprocess.Popen(
            [
                python_sistema,
                str(script_actualizacion),
                str(ruta_docx),
                ejecutable,
                str(puerto),
                str(perfil),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=entorno,
            start_new_session=True,
        )

        try:
            codigo_salida = proceso_worker.wait(timeout=100)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(
                    os.getpgid(proceso_worker.pid),
                    signal.SIGTERM,
                )
            except Exception:
                proceso_worker.terminate()

            try:
                codigo_salida = proceso_worker.wait(timeout=8)
            except Exception:
                try:
                    os.killpg(
                        os.getpgid(proceso_worker.pid),
                        signal.SIGKILL,
                    )
                except Exception:
                    proceso_worker.kill()

                codigo_salida = proceso_worker.wait(timeout=5)

        return codigo_salida == 0 and ruta_docx.exists()

    except Exception:
        return False

    finally:
        if proceso_worker is not None and proceso_worker.poll() is None:
            try:
                os.killpg(
                    os.getpgid(proceso_worker.pid),
                    signal.SIGKILL,
                )
            except Exception:
                proceso_worker.kill()

        shutil.rmtree(carpeta_temporal, ignore_errors=True)



def generar_docx_informe_tecnico_final(datos: dict) -> str:
    plantilla = obtener_ruta_plantilla()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    codigo_archivo = safe_filename(
        datos.get("codigo_proyecto", "proyecto")
    )
    nombre_archivo = f"Informe Final {codigo_archivo}.docx"
    ruta_salida = OUTPUT_DIR / nombre_archivo

    documento = Document(str(plantilla))

    eliminar_instrucciones_y_control_cambios(documento)
    configurar_estilos_y_tabla_contenido(documento)
    marcar_clasificacion(
        documento,
        datos.get("clasificacion_informacion", "Pública"),
    )
    llenar_tabla_informacion_general(documento, datos)

    fecha_portada = fecha_mes_anio_espanol(datos["fecha_entrega"])

    for parrafo in documento.paragraphs:
        if "junio 2026" in parrafo.text.casefold():
            escribir_parrafo(
                parrafo,
                fecha_portada,
                tamano=12,
                negrita=True,
                alineacion=WD_ALIGN_PARAGRAPH.CENTER,
            )
            break

    contenido = datos["contenido_informe"]

    mapa_reemplazos = [
        ("Introducción", contenido["introduccion"]),
        (
            "Planteamiento del problema",
            contenido["planteamiento_problema"],
        ),
        ("4.1 Objetivo General", contenido["objetivo_general"]),
        (
            "5. Estado del arte y estado de la técnica",
            contenido["estado_arte_tecnica"],
        ),
        (
            "6. Metodología de desarrollo",
            contenido["metodologia_desarrollo"],
        ),
        (
            "7. Desarrollo del proyecto",
            contenido["desarrollo_proyecto"],
        ),
        (
            "8. Resultados obtenidos",
            contenido["resultados_obtenidos"],
        ),
        (
            "9. Análisis de viabilidad",
            contenido["analisis_viabilidad"],
        ),
        (
            "10. Propiedad intelectual y transferencia tecnológica",
            contenido["propiedad_transferencia"],
        ),
        (
            "11. Impacto del proyecto",
            contenido["impacto_proyecto"],
        ),
        ("12. Conclusiones", contenido["conclusiones"]),
    ]

    destinos: dict[str, Paragraph] = {}

    for titulo, texto_apartado in mapa_reemplazos:
        encabezado = buscar_parrafo(documento, titulo)
        destino = parrafo_siguiente(encabezado)
        destino.style = "Normal"
        escribir_parrafo(
            destino,
            texto_apartado,
            tamano=11,
            alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY,
        )
        destinos[titulo] = destino

    encabezado_objetivos_especificos = buscar_parrafo(
        documento,
        "4.2 Objetivos Específicos",
    )
    destino_objetivos = parrafo_siguiente(
        encabezado_objetivos_especificos
    )
    escribir_lista_en_parrafos(
        destino_objetivos,
        contenido["objetivos_especificos"],
    )

    # Elimina completamente el texto introductorio genérico de Objetivos.
    encabezado_objetivos = buscar_parrafo(
        documento,
        "Objetivos",
        coincidencia_exacta=True,
    )
    introduccion_objetivos = parrafo_siguiente(encabezado_objetivos)
    eliminar_parrafo(introduccion_objetivos)

    # Tabla institucional de resultados basada únicamente en actividades.
    insertar_tabla_resultados(
        documento,
        destinos["8. Resultados obtenidos"],
        contenido["actividades_corregidas"],
    )

    encabezado_referencias = buscar_parrafo(
        documento,
        "13. Referencias bibliográficas",
    )
    destino_referencias = parrafo_siguiente(encabezado_referencias)
    escribir_lista_en_parrafos(
        destino_referencias,
        contenido["referencias_bibliograficas"],
    )

    encabezado_anexos = buscar_parrafo(
        documento,
        "14. Anexos",
        coincidencia_exacta=True,
    )
    # La plantilla trae un salto manual antes de Anexos. Se elimina para
    # evitar una página en blanco. Anexos continuará en la página disponible.
    parrafo_previo_xml = encabezado_anexos._p.getprevious()
    if parrafo_previo_xml is not None:
        for salto in parrafo_previo_xml.xpath(
            './/*[local-name()="br"]'
        ):
            padre = salto.getparent()
            if padre is not None:
                padre.remove(salto)

    encabezado_anexos.paragraph_format.page_break_before = False
    destino_anexos = parrafo_siguiente(encabezado_anexos)
    escribir_lista_en_parrafos(
        destino_anexos,
        contenido["anexos"],
    )

    documento.core_properties.title = (
        f"Informe Final {datos.get('codigo_proyecto', '')}"
    )
    documento.core_properties.subject = CODIGO_FORMATO_INFORME
    documento.core_properties.keywords = (
        "SENA, TecnoParque, Informe Final, GCDTP-F-023 V01"
    )

    configurar_estilos_y_tabla_contenido(documento)

    # Limpieza final del salto heredado antes de Anexos, después de aplicar estilos.
    encabezado_anexos_final = buscar_parrafo(
        documento,
        "14. Anexos",
        coincidencia_exacta=True,
    )
    encabezado_anexos_final.paragraph_format.page_break_before = False
    previo_final = encabezado_anexos_final._p.getprevious()
    if previo_final is not None:
        for salto in previo_final.xpath('.//*[local-name()="br"]'):
            padre = salto.getparent()
            if padre is not None:
                padre.remove(salto)

    documento.save(str(ruta_salida))

    actualizar_toc_con_libreoffice(ruta_salida)

    datos_json = serializar_datos_informe(datos)
    datos_json["ruta_docx"] = str(ruta_salida)

    guardar_datos_json(
        datos_json,
        nombre_archivo="datos_informe_tecnico_final.json",
    )

    return str(ruta_salida)


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_informe_tecnico_final(
    modo_prueba: bool = True,
    modelo_openai: str = "gpt-4.1-mini",
) -> None:
    st.markdown("---")
    st.subheader("Informe Técnico Final")
    st.caption(
        f"{CODIGO_FORMATO_INFORME} · "
        f"{VERSION_INFORME_TECNICO_FINAL}"
    )

    st.info(
        "El formulario solicita los datos institucionales, la descripción general, "
        "los entregables, la innovación, la metodología utilizada, las actividades "
        "desarrolladas y el impacto. Las actividades se escriben en un solo campo "
        "y se articulan con la metodología para generar el desarrollo del proyecto."
    )

    datos_key = "datos_informe_tecnico_final_formulario_minimo"
    ruta_key = "ruta_docx_informe_tecnico_final_formulario_minimo"

    if datos_key not in st.session_state:
        st.session_state[datos_key] = None

    if ruta_key not in st.session_state:
        st.session_state[ruta_key] = None

    with st.container():
        st.markdown("## 1. Información general del proyecto")

        col_1, col_2 = st.columns(2)

        with col_1:
            clasificacion_informacion = st.selectbox(
                "Clasificación de la información",
                options=CLASIFICACIONES_INFORMACION,
            )

            nombre_talento = st.text_input(
                "Talento que realiza el informe",
            )

            nombre_proyecto = st.text_area(
                "Nombre del Proyecto de Base Tecnológica",
                height=90,
            )

            codigo_proyecto = st.text_input(
                "Código de la idea",
                placeholder="Ejemplo: P2026-143440-18724",
            )

        with col_2:
            nombre_experto = st.text_input(
                "Experto del proyecto",
            )

            linea_tecnologica = st.text_input(
                "Línea tecnológica",
            )

            trl_inicial = st.selectbox(
                "TRL inicial",
                options=[f"TRL {numero}" for numero in range(1, 10)],
                index=4,
            )

            trl_alcanzado = st.selectbox(
                "TRL alcanzado",
                options=[f"TRL {numero}" for numero in range(1, 10)],
                index=5,
            )

            tecnoparque = st.text_input(
                "TecnoParque donde se desarrolló el proyecto",
                value="TecnoParque Nodo Angostura",
            )

            fecha_entrega = st.date_input(
                "Fecha de entrega",
                value=date.today(),
            )

        st.markdown("## Información técnica mínima")

        descripcion_general_proyecto = st.text_area(
            "Descripción general del proyecto",
            placeholder=(
                "Describe el origen, la necesidad, los usuarios, el propósito, "
                "los componentes, las tecnologías, el funcionamiento y el "
                "contexto de aplicación."
            ),
            height=340,
        )

        entregables_proyecto_base = st.text_area(
            "Entregables obtenidos",
            placeholder=(
                "Describe los productos, prototipos, componentes, documentos, "
                "sistemas o desarrollos realmente obtenidos al cierre."
            ),
            height=210,
        )

        innovacion_proyecto_base = st.text_area(
            "Innovación del proyecto",
            placeholder=(
                "Explica el elemento diferencial, la mejora frente a alternativas "
                "existentes y el aporte técnico o funcional de la solución."
            ),
            height=190,
        )

        metodologias_seleccionadas = st.multiselect(
            "Metodología utilizada",
            options=METODOLOGIAS_DESARROLLO,
            help=(
                "Selecciona una o varias metodologías aplicadas durante el proyecto. "
                "Cuando elijas Otra, se habilitará un campo para escribirla."
            ),
            key="informe_final_metodologias_v1",
        )

        otra_metodologia = ""

        if "Otra" in metodologias_seleccionadas:
            otra_metodologia = st.text_input(
                "Escribe la metodología utilizada",
                placeholder=(
                    "Ejemplo: metodología propia de diseño y validación por etapas."
                ),
                key="informe_final_otra_metodologia_v1",
            )

        st.markdown("### Actividades desarrolladas")

        actividades_ejecutadas_texto = st.text_area(
            "Descripción de las actividades ejecutadas",
            placeholder=(
                "Escribe una actividad por línea. También puedes usar numeración, "
                "viñetas o separar las actividades con punto y coma.\n\n"
                "Ejemplo:\n"
                "1. Levantamiento de requerimientos con los usuarios.\n"
                "2. Diseño de la arquitectura de la solución.\n"
                "3. Desarrollo e integración de los componentes.\n"
                "4. Pruebas, ajustes y validación del prototipo."
            ),
            height=230,
            key="informe_final_actividades_texto_v1",
            help=(
                "El sistema corregirá la ortografía, mejorará la redacción y "
                "articulará estas actividades con la metodología seleccionada."
            ),
        )

        impacto_proyecto_base = st.text_area(
            "Impacto del proyecto",
            placeholder=(
                "Describe los beneficios, beneficiarios y efectos tecnológicos, "
                "sociales, económicos, ambientales o productivos identificados."
            ),
            height=220,
        )

        generar_contenido = st.button(
            (
                "Generar contenido en modo prueba"
                if modo_prueba
                else "Generar Informe Final con la API de OpenAI"
            )
        )

    if generar_contenido:
        campos_obligatorios = {
            "Talento que realiza el informe": nombre_talento,
            "Nombre del proyecto": nombre_proyecto,
            "Código de la idea": codigo_proyecto,
            "Experto del proyecto": nombre_experto,
            "Línea tecnológica": linea_tecnologica,
            "Descripción general del proyecto": descripcion_general_proyecto,
            "Entregables obtenidos": entregables_proyecto_base,
            "Innovación del proyecto": innovacion_proyecto_base,
            "Impacto del proyecto": impacto_proyecto_base,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        metodologias_validas = [
            item
            for item in metodologias_seleccionadas
            if item != "Otra"
        ]

        if not metodologias_validas and not limpiar_texto(otra_metodologia):
            st.error(
                "Selecciona al menos una metodología o utiliza la opción Otra "
                "para escribirla manualmente."
            )
            st.stop()

        if (
            "Otra" in metodologias_seleccionadas
            and not limpiar_texto(otra_metodologia)
        ):
            st.error(
                "Escribe la metodología utilizada en el campo habilitado."
            )
            st.stop()

        actividades_validas = dividir_actividades(
            actividades_ejecutadas_texto
        )

        if not actividades_validas:
            st.error(
                "Escribe al menos una actividad desarrollada."
            )
            st.stop()

        datos_base = {
            "tipo_documento": "Informe Final",
            "codigo_formato": CODIGO_FORMATO_INFORME,
            "clasificacion_informacion": limpiar_texto(
                clasificacion_informacion
            ),
            "nombre_talento": limpiar_texto(nombre_talento),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_experto": limpiar_texto(nombre_experto),
            "linea_tecnologica": limpiar_texto(linea_tecnologica),
            "trl_inicial": limpiar_texto(trl_inicial),
            "trl_alcanzado": limpiar_texto(trl_alcanzado),
            "tecnoparque": limpiar_texto(tecnoparque),
            "fecha_entrega": fecha_entrega,
            "fecha_entrega_texto": fecha_entrega.strftime("%d/%m/%Y"),
            "descripcion_general_proyecto": limpiar_texto(
                descripcion_general_proyecto
            ),
            "entregables_proyecto_base": limpiar_texto(
                entregables_proyecto_base
            ),
            "innovacion_proyecto_base": limpiar_texto(
                innovacion_proyecto_base
            ),
            "metodologias_seleccionadas": metodologias_seleccionadas,
            "otra_metodologia": limpiar_texto(otra_metodologia),
            "actividades_ejecutadas_base": actividades_validas,
            "impacto_proyecto_base": limpiar_texto(
                impacto_proyecto_base
            ),
            "modo_generacion": (
                "Prueba local"
                if modo_prueba
                else "API de OpenAI con búsqueda web"
            ),
            "version": VERSION_INFORME_TECNICO_FINAL,
        }
        datos_base["metodologia_inferida"] = (
            metodologias_en_texto(datos_base)
        )

        with st.spinner(
            "Investigando dos referentes reales y generando los apartados "
            "del informe entre 300 y 340 palabras."
        ):
            try:
                if modo_prueba:
                    datos_base["referentes_estado_arte"] = (
                        referentes_modo_prueba(datos_base)
                    )
                    contenido = contenido_modo_prueba(datos_base)
                else:
                    datos_base["referentes_estado_arte"] = (
                        investigar_referentes_reales(
                            datos_base,
                            modelo_openai,
                        )
                    )
                    contenido = generar_contenido_con_ia(
                        datos_base,
                        modelo_openai,
                    )
            except Exception as error:
                st.error(
                    f"No se pudo generar el contenido del informe: {error}"
                )
                st.stop()

        datos_base["contenido_informe"] = contenido
        st.session_state[datos_key] = datos_base
        st.session_state[ruta_key] = None

        st.success(
            "Contenido generado. Puedes revisarlo antes de crear el Word oficial."
        )

    datos = st.session_state.get(datos_key)

    if datos:
        contenido = datos["contenido_informe"]

        st.success(
            "Todos los apartados del informe fueron generados automáticamente. "
            "No es necesario diligenciar información adicional por sección."
        )

        st.markdown("### Resumen de generación")
        st.write(
            "**Metodología utilizada:**",
            metodologias_en_texto(datos),
        )
        st.write(
            "**Actividades procesadas:**",
            len(contenido.get("actividades_corregidas", [])),
        )
        st.caption(
            "El desarrollo del proyecto fue redactado articulando la metodología "
            "seleccionada con las actividades registradas."
        )

        st.markdown("### Vista previa de la tabla de resultados")
        tabla_resultados = [
            {
                "N.°": indice,
                "Descripción de la actividad ejecutada": actividad,
                "Evidencia": "Agregar enlace en Word",
            }
            for indice, actividad in enumerate(
                contenido["actividades_corregidas"],
                start=1,
            )
        ]
        st.dataframe(
            tabla_resultados,
            use_container_width=True,
            hide_index=True,
        )

        col_json, col_docx = st.columns(2)

        with col_json:
            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(
                    serializar_datos_informe(datos),
                    ensure_ascii=False,
                    indent=4,
                ),
                file_name="datos_informe_tecnico_final.json",
                mime="application/json",
            )

        with col_docx:
            if st.button(
                "📄 Generar Word oficial GCDTP-F-023 V01",
                key="generar_docx_informe_tecnico_final_minimo",
            ):
                try:
                    ruta_docx = generar_docx_informe_tecnico_final(
                        datos
                    )
                    st.session_state[ruta_key] = ruta_docx
                    st.success(
                        "Documento Word generado correctamente."
                    )
                except Exception as error:
                    st.error(
                        f"No se pudo generar el documento Word: {error}"
                    )

        ruta_docx = st.session_state.get(ruta_key)

        if ruta_docx and Path(ruta_docx).exists():
            with open(ruta_docx, "rb") as archivo_docx:
                st.download_button(
                    label="⬇️ Descargar Informe Final oficial",
                    data=archivo_docx,
                    file_name=Path(ruta_docx).name,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),
                )
