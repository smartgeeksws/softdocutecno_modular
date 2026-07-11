from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
import json
from pathlib import Path
import tempfile

import streamlit as st

from docx import Document
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
    "VERSION_MODULAR_INFORME_FINAL_GCDTP_F_023_V01"
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

CLAVES_CONTENIDO = [
    "introduccion",
    "planteamiento_problema",
    "objetivo_general",
    "objetivos_especificos",
    "estado_arte_tecnica",
    "metodologia_desarrollo",
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


def guardar_archivos_anexos(
    archivos_subidos,
    codigo_proyecto: str,
) -> list[dict]:
    if not archivos_subidos:
        return []

    carpeta = (
        Path(tempfile.gettempdir())
        / "softdocutecno_informe_final"
        / safe_filename(codigo_proyecto)
    )
    carpeta.mkdir(parents=True, exist_ok=True)

    anexos: list[dict] = []

    for indice, archivo in enumerate(archivos_subidos, start=1):
        nombre_seguro = safe_filename(Path(archivo.name).stem)
        extension = Path(archivo.name).suffix.lower()
        ruta = carpeta / f"{indice:02d}_{nombre_seguro}{extension}"
        ruta.write_bytes(archivo.getbuffer())

        anexos.append(
            {
                "nombre_original": archivo.name,
                "ruta": str(ruta),
                "extension": extension,
                "es_imagen": extension in {
                    ".png",
                    ".jpg",
                    ".jpeg",
                    ".webp",
                },
            }
        )

    return anexos


def serializar_datos_informe(datos: dict) -> dict:
    resultado = deepcopy(datos)

    fecha_entrega = resultado.get("fecha_entrega")
    if isinstance(fecha_entrega, date):
        resultado["fecha_entrega"] = fecha_entrega.strftime("%d/%m/%Y")

    return resultado


# =====================================================
# GENERACIÓN DE CONTENIDO
# =====================================================

def contenido_modo_prueba(datos: dict) -> dict:
    especificos = dividir_lineas(
        datos.get("objetivos_especificos_base", "")
    )

    if len(especificos) < 4:
        especificos = [
            "Identificar los requerimientos técnicos, funcionales y operativos de la solución.",
            "Diseñar los componentes y la arquitectura necesaria para el desarrollo del prototipo.",
            "Implementar e integrar los componentes definidos mediante un proceso iterativo de construcción.",
            "Validar el funcionamiento de la solución y documentar los resultados obtenidos.",
        ]

    referencias = dividir_lineas(
        datos.get("referencias_suministradas", "")
    )

    anexos = dividir_lineas(datos.get("descripcion_anexos", ""))

    return {
        "introduccion": (
            f"El presente informe final documenta el desarrollo del proyecto de base "
            f"tecnológica {datos.get('nombre_proyecto', '')}, identificado con el código "
            f"{datos.get('codigo_proyecto', '')}. La iniciativa se desarrolló en "
            f"{datos.get('tecnoparque', '')}, con el acompañamiento del experto "
            f"{datos.get('nombre_experto', '')}. El contexto, propósito y entorno del "
            f"proyecto fueron los siguientes: {datos.get('contexto_proposito', '')} "
            f"El trabajo se orientó a obtener una solución verificable y coherente con "
            f"la línea tecnológica {datos.get('linea_tecnologica', '')}, partiendo de "
            f"un nivel inicial {datos.get('trl_inicial', '')} y alcanzando el nivel "
            f"{datos.get('trl_alcanzado', '')}."
        ),
        "planteamiento_problema": (
            f"La problemática o necesidad que dio origen al proyecto se describe así: "
            f"{datos.get('problema_necesidad', '')} Las causas, consecuencias y el "
            f"impacto reportado fueron: {datos.get('causas_consecuencias', '')} "
            f"Esta situación justificó la construcción de una solución tecnológica "
            f"orientada a reducir las limitaciones identificadas y generar valor para "
            f"los usuarios o beneficiarios definidos."
        ),
        "objetivo_general": (
            limpiar_texto(datos.get("objetivo_general_base", ""))
            or (
                "Desarrollar una solución de base tecnológica que responda a la "
                "problemática identificada mediante actividades de diseño, "
                "implementación, prototipado y validación."
            )
        ),
        "objetivos_especificos": especificos[:4],
        "estado_arte_tecnica": (
            f"El análisis del estado del arte y de la técnica consideró los siguientes "
            f"referentes, tendencias, tecnologías y soluciones existentes: "
            f"{datos.get('referentes_estado_arte', '')} La propuesta desarrollada se "
            f"diferencia o aporta valor frente a las alternativas disponibles por: "
            f"{datos.get('diferenciador_innovacion', '')} Las referencias suministradas "
            f"deben verificarse y mantenerse asociadas con las afirmaciones técnicas "
            f"incluidas en este apartado."
        ),
        "metodologia_desarrollo": (
            f"El proyecto se desarrolló mediante el siguiente enfoque metodológico, "
            f"fases, actividades y herramientas: {datos.get('metodologia_fases', '')} "
            f"El proceso de validación, prototipado y ajuste se ejecutó de la siguiente "
            f"manera: {datos.get('proceso_validacion', '')} La metodología permitió "
            f"mantener trazabilidad entre la necesidad, las decisiones técnicas, los "
            f"resultados y los ajustes realizados."
        ),
        "desarrollo_proyecto": (
            f"Durante la ejecución se realizaron las siguientes actividades de diseño, "
            f"construcción, configuración e integración: "
            f"{datos.get('actividades_desarrollo', '')} Los componentes y decisiones "
            f"técnicas principales fueron: {datos.get('componentes_decisiones', '')} "
            f"El desarrollo avanzó de manera iterativa, incorporando revisiones y "
            f"ajustes conforme a los resultados parciales obtenidos."
        ),
        "resultados_obtenidos": (
            f"Los productos, entregables, prototipos y desarrollos alcanzados fueron: "
            f"{datos.get('resultados_entregables', '')} Las pruebas, validaciones y "
            f"evidencias reportadas fueron: {datos.get('pruebas_evidencias', '')} "
            f"Estos resultados contribuyen al cumplimiento de los objetivos y permiten "
            f"demostrar el avance técnico alcanzado por la solución."
        ),
        "analisis_viabilidad": (
            f"La viabilidad técnica, operativa, económica, normativa y de mercado se "
            f"analizó con base en la siguiente información: "
            f"{datos.get('viabilidad_multidimensional', '')} Las condiciones para su "
            f"adopción, las limitaciones y las oportunidades de continuidad, "
            f"escalabilidad o sostenibilidad son: "
            f"{datos.get('limitaciones_escalabilidad', '')}"
        ),
        "propiedad_transferencia": (
            f"Los activos de propiedad intelectual generados o con potencial de "
            f"protección corresponden a: {datos.get('activos_propiedad_intelectual', '')} "
            f"Las oportunidades y estrategias de transferencia, adopción, "
            f"comercialización o apropiación se plantean así: "
            f"{datos.get('estrategia_transferencia', '')}"
        ),
        "impacto_proyecto": (
            f"Los beneficios y efectos tecnológicos, sociales, económicos, ambientales "
            f"o productivos identificados son: {datos.get('impactos_beneficios', '')} "
            f"Los principales usuarios o beneficiarios y la forma en que reciben valor "
            f"se describen de la siguiente manera: {datos.get('beneficiarios_valor', '')}"
        ),
        "conclusiones": (
            f"El proyecto permitió consolidar una solución pertinente frente a la "
            f"problemática identificada y alcanzar resultados coherentes con el nivel "
            f"{datos.get('trl_alcanzado', '')}. Las principales conclusiones, lecciones "
            f"aprendidas y oportunidades de mejora, continuidad o escalamiento son: "
            f"{datos.get('conclusiones_futuro', '')}"
        ),
        "referencias_bibliograficas": (
            referencias
            or [
                "No se suministraron referencias bibliográficas verificables. "
                "Este apartado debe completarse antes de aprobar el informe definitivo."
            ]
        ),
        "anexos": (
            anexos
            or [
                "Evidencias técnicas y documentales relacionadas con el desarrollo "
                "y la validación del proyecto."
            ]
        ),
    }


def normalizar_contenido(
    contenido: object,
    respaldo: dict,
) -> dict:
    if not isinstance(contenido, dict):
        return respaldo

    resultado: dict = {}

    for clave in CLAVES_CONTENIDO:
        valor = contenido.get(clave)

        if clave in {
            "objetivos_especificos",
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
            texto = limpiar_texto(str(valor or ""))
            resultado[clave] = texto or respaldo[clave]

    objetivos = resultado["objetivos_especificos"]
    objetivos_respaldo = respaldo["objetivos_especificos"]

    for objetivo in objetivos_respaldo:
        if len(objetivos) >= 4:
            break
        if objetivo not in objetivos:
            objetivos.append(objetivo)

    resultado["objetivos_especificos"] = objetivos[:4]

    return resultado


def generar_contenido_con_ia(
    datos: dict,
    modelo_openai: str,
) -> dict:
    respaldo = contenido_modo_prueba(datos)

    instrucciones = """
Actúa como redactor técnico senior de proyectos de base tecnológica de la Red
TecnoParque SENA. Genera el contenido de un Informe Final institucional.

Reglas obligatorias:
- Redacta en español formal, técnico, claro y verificable.
- Usa exclusivamente la información suministrada.
- No inventes nombres, cifras, pruebas, resultados, normas, referencias, patentes,
  aceptaciones, costos ni impactos.
- No afirmes que una prueba fue exitosa si el usuario no lo indicó.
- No inventes referencias bibliográficas. Utiliza solo las suministradas.
- El objetivo general debe iniciar con un verbo en infinitivo.
- Genera exactamente cuatro objetivos específicos, cada uno con verbo en infinitivo.
- Evita repetir el nombre del proyecto en todos los párrafos.
- El texto debe corresponder a un informe de cierre, no a una propuesta futura.
- Responde únicamente en JSON válido, sin markdown.

Extensión orientativa:
- Introducción: 180 a 260 palabras.
- Planteamiento del problema: 180 a 280 palabras.
- Estado del arte y estado de la técnica: 250 a 400 palabras.
- Metodología de desarrollo: 250 a 400 palabras.
- Desarrollo del proyecto: 300 a 500 palabras.
- Resultados obtenidos: 220 a 350 palabras.
- Análisis de viabilidad: 250 a 400 palabras.
- Propiedad intelectual y transferencia: 180 a 300 palabras.
- Impacto: 180 a 300 palabras.
- Conclusiones: 160 a 260 palabras.
"""

    entrada = f"""
INFORMACIÓN GENERAL
Tipo de proyecto: {datos.get('tipo_proyecto', '')}
Clasificación: {datos.get('clasificacion_informacion', '')}
Talento: {datos.get('nombre_talento', '')}
Nombre del proyecto: {datos.get('nombre_proyecto', '')}
Código: {datos.get('codigo_proyecto', '')}
Experto: {datos.get('nombre_experto', '')}
Línea tecnológica: {datos.get('linea_tecnologica', '')}
TRL inicial: {datos.get('trl_inicial', '')}
TRL alcanzado: {datos.get('trl_alcanzado', '')}
TecnoParque: {datos.get('tecnoparque', '')}

CONTEXTO Y PROBLEMA
Contexto, origen, propósito y entorno:
{datos.get('contexto_proposito', '')}

Problema o necesidad:
{datos.get('problema_necesidad', '')}

Causas, consecuencias e impacto:
{datos.get('causas_consecuencias', '')}

OBJETIVOS BASE
Objetivo general:
{datos.get('objetivo_general_base', '')}

Objetivos específicos:
{datos.get('objetivos_especificos_base', '')}

ESTADO DEL ARTE Y DIFERENCIACIÓN
Referentes, tendencias, tecnologías y soluciones existentes:
{datos.get('referentes_estado_arte', '')}

Diferenciador, innovación o mejora:
{datos.get('diferenciador_innovacion', '')}

METODOLOGÍA
Enfoque, fases, actividades y herramientas:
{datos.get('metodologia_fases', '')}

Prototipado, pruebas, validación y ajustes:
{datos.get('proceso_validacion', '')}

DESARROLLO Y RESULTADOS
Actividades de diseño, construcción, configuración e integración:
{datos.get('actividades_desarrollo', '')}

Componentes y decisiones técnicas:
{datos.get('componentes_decisiones', '')}

Resultados, productos, entregables o prototipos:
{datos.get('resultados_entregables', '')}

Pruebas, validaciones y evidencias:
{datos.get('pruebas_evidencias', '')}

VIABILIDAD
Aspectos técnicos, operativos, económicos, normativos y de mercado:
{datos.get('viabilidad_multidimensional', '')}

Condiciones, limitaciones, sostenibilidad y escalabilidad:
{datos.get('limitaciones_escalabilidad', '')}

PROPIEDAD INTELECTUAL Y TRANSFERENCIA
Activos o resultados con potencial de protección:
{datos.get('activos_propiedad_intelectual', '')}

Estrategias de transferencia, adopción, comercialización o apropiación:
{datos.get('estrategia_transferencia', '')}

IMPACTO
Beneficios y efectos tecnológicos, sociales, económicos, ambientales o productivos:
{datos.get('impactos_beneficios', '')}

Usuarios, beneficiarios y valor generado:
{datos.get('beneficiarios_valor', '')}

CONCLUSIONES
Conclusiones, lecciones, mejoras, continuidad y escalamiento:
{datos.get('conclusiones_futuro', '')}

REFERENCIAS SUMINISTRADAS
{datos.get('referencias_suministradas', '')}

ANEXOS DESCRITOS
{datos.get('descripcion_anexos', '')}

ESTRUCTURA JSON OBLIGATORIA
{{
  "introduccion": "texto",
  "planteamiento_problema": "texto",
  "objetivo_general": "texto",
  "objetivos_especificos": ["objetivo 1", "objetivo 2", "objetivo 3", "objetivo 4"],
  "estado_arte_tecnica": "texto",
  "metodologia_desarrollo": "texto",
  "desarrollo_proyecto": "texto",
  "resultados_obtenidos": "texto",
  "analisis_viabilidad": "texto",
  "propiedad_transferencia": "texto",
  "impacto_proyecto": "texto",
  "conclusiones": "texto",
  "referencias_bibliograficas": ["referencia 1", "referencia 2"],
  "anexos": ["anexo 1", "anexo 2"]
}}
"""

    contenido = generar_json_openai(
        instrucciones=instrucciones,
        entrada=entrada,
        modelo=modelo_openai,
        temperature=0.2,
    )

    return normalizar_contenido(contenido, respaldo)


# =====================================================
# GENERACIÓN DEL DOCUMENTO OFICIAL
# =====================================================

def agregar_anexos_visuales(
    documento: Document,
    parrafo_anexos: Paragraph,
    archivos_anexos: list[dict],
) -> None:
    ancla = parrafo_anexos

    for indice, archivo in enumerate(archivos_anexos, start=1):
        titulo = insertar_parrafo_despues(
            ancla,
            f"Anexo {indice}. {archivo.get('nombre_original', '')}",
            estilo="Heading 2",
        )
        ancla = titulo

        if archivo.get("es_imagen") and Path(
            archivo.get("ruta", "")
        ).exists():
            parrafo_imagen = insertar_parrafo_despues(ancla)
            parrafo_imagen.alignment = WD_ALIGN_PARAGRAPH.CENTER

            try:
                parrafo_imagen.add_run().add_picture(
                    archivo["ruta"],
                    width=Cm(14.5),
                )
            except Exception:
                escribir_parrafo(
                    parrafo_imagen,
                    (
                        "No fue posible insertar la imagen. "
                        f"Archivo: {archivo.get('nombre_original', '')}"
                    ),
                )

            ancla = parrafo_imagen
        else:
            detalle = insertar_parrafo_despues(
                ancla,
                (
                    "Archivo complementario relacionado: "
                    f"{archivo.get('nombre_original', '')}"
                ),
            )
            ancla = detalle


def generar_docx_informe_tecnico_final(datos: dict) -> str:
    plantilla = obtener_ruta_plantilla()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Informe_Final_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    )
    ruta_salida = OUTPUT_DIR / nombre_archivo

    documento = Document(str(plantilla))

    eliminar_instrucciones_y_control_cambios(documento)
    marcar_actualizacion_campos(documento)
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
        (
            "Introducción",
            contenido["introduccion"],
        ),
        (
            "Planteamiento del problema",
            contenido["planteamiento_problema"],
        ),
        (
            "4.1 Objetivo General",
            contenido["objetivo_general"],
        ),
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
        (
            "12. Conclusiones",
            contenido["conclusiones"],
        ),
    ]

    for titulo, texto_apartado in mapa_reemplazos:
        encabezado = buscar_parrafo(documento, titulo)
        destino = parrafo_siguiente(encabezado)
        escribir_parrafo(destino, texto_apartado, tamano=11)

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

    encabezado_referencias = buscar_parrafo(
        documento,
        "13. Referencias bibliográficas",
    )
    destino_referencias = parrafo_siguiente(encabezado_referencias)
    escribir_lista_en_parrafos(
        destino_referencias,
        contenido["referencias_bibliograficas"],
    )

    encabezado_anexos = buscar_parrafo(documento, "14. Anexos")
    destino_anexos = parrafo_siguiente(encabezado_anexos)
    parrafo_anexos = escribir_lista_en_parrafos(
        destino_anexos,
        contenido["anexos"],
    )

    # El texto instructivo general del numeral 4 no debe quedar en el informe final.
    encabezado_objetivos = buscar_parrafo(
        documento,
        "Objetivos",
        coincidencia_exacta=True,
    )
    introduccion_objetivos = parrafo_siguiente(encabezado_objetivos)
    escribir_parrafo(
        introduccion_objetivos,
        (
            "Los objetivos formulados orientaron el desarrollo y sirvieron "
            "como referencia para evaluar el cumplimiento y alcance de la "
            "solución tecnológica."
        ),
        tamano=11,
    )

    agregar_anexos_visuales(
        documento,
        parrafo_anexos,
        datos.get("archivos_anexos", []),
    )

    documento.core_properties.title = (
        f"Informe Final - {datos.get('nombre_proyecto', '')}"
    )
    documento.core_properties.subject = CODIGO_FORMATO_INFORME
    documento.core_properties.keywords = (
        "SENA, TecnoParque, Informe Final, GCDTP-F-023 V01"
    )

    documento.save(str(ruta_salida))

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
        "El documento se genera directamente sobre la plantilla oficial "
        "GCDTP-F-023 V01. Conserva el logo, la portada, la tabla de "
        "contenido, el encabezado, el pie de página y el código del formato. "
        "Las páginas de instrucciones y control de cambios se eliminan "
        "automáticamente del archivo final."
    )

    if "datos_informe_tecnico_final_generado" not in st.session_state:
        st.session_state.datos_informe_tecnico_final_generado = None

    if "ruta_docx_informe_tecnico_final_generado" not in st.session_state:
        st.session_state.ruta_docx_informe_tecnico_final_generado = None

    with st.form("form_informe_tecnico_final_gcdtp_023"):
        st.markdown("## 1. Información general del proyecto")

        col_1, col_2 = st.columns(2)

        with col_1:
            tipo_proyecto = st.selectbox(
                "Tipo de proyecto",
                options=TIPOS_PROYECTO_INFORME,
            )

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

        st.markdown("## 2 y 3. Introducción y planteamiento del problema")

        contexto_proposito = st.text_area(
            "Contexto, origen, propósito y entorno del proyecto",
            placeholder=(
                "Indica cómo surgió la iniciativa, qué busca desarrollar, "
                "quién la promovió, dónde se aplicó y cuál fue su contexto."
            ),
            height=150,
        )

        problema_necesidad = st.text_area(
            "Problemática o necesidad identificada",
            placeholder=(
                "Describe la situación concreta que se buscó resolver y "
                "quiénes resultaban afectados."
            ),
            height=130,
        )

        causas_consecuencias = st.text_area(
            "Causas, consecuencias e impacto del problema",
            placeholder=(
                "Incluye causas principales, efectos, limitaciones actuales "
                "e impacto técnico, social, productivo o económico."
            ),
            height=130,
        )

        st.markdown("## 4. Objetivos")

        objetivo_general_base = st.text_area(
            "Objetivo general",
            placeholder="Debe iniciar con un verbo en infinitivo.",
            height=90,
        )

        objetivos_especificos_base = st.text_area(
            "Objetivos específicos",
            placeholder=(
                "Escribe cuatro objetivos, uno por línea. "
                "Cada objetivo debe iniciar con un verbo en infinitivo."
            ),
            height=150,
        )

        st.markdown("## 5. Estado del arte y estado de la técnica")

        referentes_estado_arte = st.text_area(
            "Investigaciones, tecnologías, productos, referentes y tendencias",
            placeholder=(
                "Describe las soluciones existentes y los referentes "
                "consultados. Incluye nombres y datos verificables."
            ),
            height=150,
        )

        diferenciador_innovacion = st.text_area(
            "Valor diferencial, innovación o mejora frente a las alternativas",
            height=120,
        )

        st.markdown("## 6. Metodología de desarrollo")

        metodologia_fases = st.text_area(
            "Enfoque, fases, actividades y herramientas utilizadas",
            placeholder=(
                "Ejemplo: Design Thinking, desarrollo iterativo, diseño CAD, "
                "programación, fabricación, integración y documentación."
            ),
            height=160,
        )

        proceso_validacion = st.text_area(
            "Prototipado, pruebas, validación, retroalimentación y ajustes",
            height=140,
        )

        st.markdown("## 7 y 8. Desarrollo y resultados")

        actividades_desarrollo = st.text_area(
            "Actividades ejecutadas durante el desarrollo",
            placeholder=(
                "Describe diseño, construcción, configuración, integración, "
                "puesta en funcionamiento y avances por etapas."
            ),
            height=170,
        )

        componentes_decisiones = st.text_area(
            "Componentes, tecnologías y decisiones técnicas adoptadas",
            height=140,
        )

        resultados_entregables = st.text_area(
            "Resultados, productos, entregables y prototipos obtenidos",
            height=150,
        )

        pruebas_evidencias = st.text_area(
            "Pruebas, validaciones y evidencias disponibles",
            placeholder=(
                "No afirmes resultados no comprobados. Indica qué se probó, "
                "cómo se verificó y qué evidencias existen."
            ),
            height=150,
        )

        st.markdown("## 9. Análisis de viabilidad")

        viabilidad_multidimensional = st.text_area(
            "Viabilidad técnica, operativa, económica, normativa y de mercado",
            height=160,
        )

        limitaciones_escalabilidad = st.text_area(
            "Condiciones de adopción, limitaciones, sostenibilidad y escalabilidad",
            height=140,
        )

        st.markdown(
            "## 10. Propiedad intelectual y transferencia tecnológica"
        )

        activos_propiedad_intelectual = st.text_area(
            "Activos generados o con potencial de protección",
            placeholder=(
                "Software, código fuente, diseños, modelos, prototipos, "
                "marcas, obras, invenciones u otros resultados."
            ),
            height=130,
        )

        estrategia_transferencia = st.text_area(
            "Estrategia de transferencia, adopción, comercialización o apropiación",
            height=130,
        )

        st.markdown("## 11 y 12. Impacto y conclusiones")

        impactos_beneficios = st.text_area(
            "Beneficios e impactos tecnológicos, sociales, económicos, ambientales o productivos",
            height=150,
        )

        beneficiarios_valor = st.text_area(
            "Usuarios o beneficiarios y valor generado",
            height=120,
        )

        conclusiones_futuro = st.text_area(
            "Conclusiones, lecciones aprendidas, oportunidades de mejora, continuidad y escalamiento",
            height=150,
        )

        st.markdown("## 13 y 14. Referencias y anexos")

        referencias_suministradas = st.text_area(
            "Referencias bibliográficas verificables",
            placeholder=(
                "Escribe una referencia por línea en formato APA, IEEE u otro "
                "estilo reconocido. No se inventarán referencias."
            ),
            height=160,
        )

        descripcion_anexos = st.text_area(
            "Descripción de los anexos",
            placeholder=(
                "Escribe un anexo por línea: fotografías, diagramas, planos, "
                "manuales, pruebas, actas, código fuente o evidencias."
            ),
            height=130,
        )

        archivos_anexos_upload = st.file_uploader(
            "Archivos de anexos y evidencias",
            type=[
                "png",
                "jpg",
                "jpeg",
                "webp",
                "pdf",
                "docx",
                "xlsx",
                "csv",
                "zip",
            ],
            accept_multiple_files=True,
            help=(
                "Las imágenes se insertan en el documento. Los demás archivos "
                "se relacionan por nombre en la sección de anexos."
            ),
        )

        generar_contenido = st.form_submit_button(
            (
                "Generar contenido en modo prueba"
                if modo_prueba
                else "Generar contenido del Informe Final con IA"
            )
        )

    if generar_contenido:
        campos_obligatorios = {
            "Talento que realiza el informe": nombre_talento,
            "Nombre del proyecto": nombre_proyecto,
            "Código de la idea": codigo_proyecto,
            "Experto del proyecto": nombre_experto,
            "Línea tecnológica": linea_tecnologica,
            "Contexto y propósito": contexto_proposito,
            "Problema o necesidad": problema_necesidad,
            "Objetivo general": objetivo_general_base,
            "Objetivos específicos": objetivos_especificos_base,
            "Metodología": metodologia_fases,
            "Actividades de desarrollo": actividades_desarrollo,
            "Resultados y entregables": resultados_entregables,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        try:
            archivos_anexos = guardar_archivos_anexos(
                archivos_anexos_upload,
                codigo_proyecto,
            )
        except Exception as error:
            st.error(f"No se pudieron procesar los anexos: {error}")
            st.stop()

        datos_base = {
            "tipo_documento": "Informe Final",
            "codigo_formato": CODIGO_FORMATO_INFORME,
            "tipo_proyecto": limpiar_texto(tipo_proyecto),
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
            "contexto_proposito": limpiar_texto(contexto_proposito),
            "problema_necesidad": limpiar_texto(problema_necesidad),
            "causas_consecuencias": limpiar_texto(
                causas_consecuencias
            ),
            "objetivo_general_base": limpiar_texto(
                objetivo_general_base
            ),
            "objetivos_especificos_base": limpiar_texto(
                objetivos_especificos_base
            ),
            "referentes_estado_arte": limpiar_texto(
                referentes_estado_arte
            ),
            "diferenciador_innovacion": limpiar_texto(
                diferenciador_innovacion
            ),
            "metodologia_fases": limpiar_texto(metodologia_fases),
            "proceso_validacion": limpiar_texto(proceso_validacion),
            "actividades_desarrollo": limpiar_texto(
                actividades_desarrollo
            ),
            "componentes_decisiones": limpiar_texto(
                componentes_decisiones
            ),
            "resultados_entregables": limpiar_texto(
                resultados_entregables
            ),
            "pruebas_evidencias": limpiar_texto(
                pruebas_evidencias
            ),
            "viabilidad_multidimensional": limpiar_texto(
                viabilidad_multidimensional
            ),
            "limitaciones_escalabilidad": limpiar_texto(
                limitaciones_escalabilidad
            ),
            "activos_propiedad_intelectual": limpiar_texto(
                activos_propiedad_intelectual
            ),
            "estrategia_transferencia": limpiar_texto(
                estrategia_transferencia
            ),
            "impactos_beneficios": limpiar_texto(
                impactos_beneficios
            ),
            "beneficiarios_valor": limpiar_texto(
                beneficiarios_valor
            ),
            "conclusiones_futuro": limpiar_texto(
                conclusiones_futuro
            ),
            "referencias_suministradas": limpiar_texto(
                referencias_suministradas
            ),
            "descripcion_anexos": limpiar_texto(
                descripcion_anexos
            ),
            "archivos_anexos": archivos_anexos,
            "modo_generacion": (
                "Prueba local"
                if modo_prueba
                else "ChatGPT API"
            ),
            "version": VERSION_INFORME_TECNICO_FINAL,
        }

        with st.spinner(
            "Generando los apartados del Informe Final."
        ):
            try:
                if modo_prueba:
                    contenido = contenido_modo_prueba(datos_base)
                else:
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

        st.session_state.datos_informe_tecnico_final_generado = (
            datos_base
        )
        st.session_state.ruta_docx_informe_tecnico_final_generado = (
            None
        )

        st.success(
            "Contenido generado. Revísalo y edítalo antes de crear "
            "el documento oficial."
        )

    datos = st.session_state.get(
        "datos_informe_tecnico_final_generado"
    )

    if datos:
        contenido = datos["contenido_informe"]

        st.markdown("## Revisión y edición del contenido")
        st.caption(
            "Los cambios deben guardarse antes de generar el archivo Word."
        )

        with st.form("form_revisar_informe_tecnico_final"):
            introduccion_editada = st.text_area(
                "2. Introducción",
                value=contenido["introduccion"],
                height=220,
            )

            problema_editado = st.text_area(
                "3. Planteamiento del problema",
                value=contenido["planteamiento_problema"],
                height=220,
            )

            objetivo_general_editado = st.text_area(
                "4.1 Objetivo General",
                value=contenido["objetivo_general"],
                height=90,
            )

            objetivos_editados = st.text_area(
                "4.2 Objetivos Específicos",
                value="\n".join(
                    contenido["objetivos_especificos"]
                ),
                height=150,
            )

            estado_arte_editado = st.text_area(
                "5. Estado del arte y estado de la técnica",
                value=contenido["estado_arte_tecnica"],
                height=260,
            )

            metodologia_editada = st.text_area(
                "6. Metodología de desarrollo",
                value=contenido["metodologia_desarrollo"],
                height=260,
            )

            desarrollo_editado = st.text_area(
                "7. Desarrollo del proyecto",
                value=contenido["desarrollo_proyecto"],
                height=300,
            )

            resultados_editados = st.text_area(
                "8. Resultados obtenidos",
                value=contenido["resultados_obtenidos"],
                height=240,
            )

            viabilidad_editada = st.text_area(
                "9. Análisis de viabilidad",
                value=contenido["analisis_viabilidad"],
                height=260,
            )

            propiedad_editada = st.text_area(
                "10. Propiedad intelectual y transferencia tecnológica",
                value=contenido["propiedad_transferencia"],
                height=220,
            )

            impacto_editado = st.text_area(
                "11. Impacto del proyecto",
                value=contenido["impacto_proyecto"],
                height=220,
            )

            conclusiones_editadas = st.text_area(
                "12. Conclusiones",
                value=contenido["conclusiones"],
                height=220,
            )

            referencias_editadas = st.text_area(
                "13. Referencias bibliográficas",
                value="\n".join(
                    contenido["referencias_bibliograficas"]
                ),
                height=180,
            )

            anexos_editados = st.text_area(
                "14. Anexos",
                value="\n".join(contenido["anexos"]),
                height=160,
            )

            guardar_edicion = st.form_submit_button(
                "💾 Guardar cambios del informe"
            )

        if guardar_edicion:
            datos["contenido_informe"] = {
                "introduccion": limpiar_texto(
                    introduccion_editada
                ),
                "planteamiento_problema": limpiar_texto(
                    problema_editado
                ),
                "objetivo_general": limpiar_texto(
                    objetivo_general_editado
                ),
                "objetivos_especificos": dividir_lineas(
                    objetivos_editados
                )[:4],
                "estado_arte_tecnica": limpiar_texto(
                    estado_arte_editado
                ),
                "metodologia_desarrollo": limpiar_texto(
                    metodologia_editada
                ),
                "desarrollo_proyecto": limpiar_texto(
                    desarrollo_editado
                ),
                "resultados_obtenidos": limpiar_texto(
                    resultados_editados
                ),
                "analisis_viabilidad": limpiar_texto(
                    viabilidad_editada
                ),
                "propiedad_transferencia": limpiar_texto(
                    propiedad_editada
                ),
                "impacto_proyecto": limpiar_texto(
                    impacto_editado
                ),
                "conclusiones": limpiar_texto(
                    conclusiones_editadas
                ),
                "referencias_bibliograficas": dividir_lineas(
                    referencias_editadas
                ),
                "anexos": dividir_lineas(anexos_editados),
            }

            st.session_state.datos_informe_tecnico_final_generado = (
                datos
            )
            st.session_state.ruta_docx_informe_tecnico_final_generado = (
                None
            )

            st.success("Cambios guardados correctamente.")
            st.rerun()

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
                key="generar_docx_informe_tecnico_final",
            ):
                try:
                    ruta_docx = generar_docx_informe_tecnico_final(
                        datos
                    )
                    st.session_state.ruta_docx_informe_tecnico_final_generado = (
                        ruta_docx
                    )
                    st.success(
                        "Documento Word generado correctamente."
                    )
                except Exception as error:
                    st.error(
                        f"No se pudo generar el documento Word: {error}"
                    )

        ruta_docx = st.session_state.get(
            "ruta_docx_informe_tecnico_final_generado"
        )

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