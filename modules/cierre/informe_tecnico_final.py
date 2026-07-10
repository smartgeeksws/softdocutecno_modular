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
    "VERSION_MODULAR_INFORME_FINAL_GCDTP_F_023_V01_FORMULARIO_SIMPLIFICADO_400"
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
    "Design Thinking",
    "Doble Diamante",
    "Diseño Centrado en el Usuario (DCU)",
    "Metodologías ágiles",
    "Scrum",
    "Kanban",
    "Lean Startup",
    "Stage-Gate",
    "Desarrollo iterativo de prototipos",
    "Ingeniería de sistemas y modelo V",
    "Ciclo de vida en cascada",
    "CRISP-DM para proyectos de datos e inteligencia artificial",
    "DMAIC / Six Sigma",
    "Diseño para Manufactura y Ensamble (DFMA)",
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

def metodologias_en_texto(datos: dict) -> str:
    seleccionadas = datos.get("metodologias_seleccionadas", []) or []
    otras = limpiar_texto(datos.get("otra_metodologia", ""))

    metodologias = [
        limpiar_texto(str(item))
        for item in seleccionadas
        if limpiar_texto(str(item)) and str(item) != "Otra"
    ]

    if otras:
        metodologias.append(otras)

    return ", ".join(metodologias)


def referencias_internas_proyecto(datos: dict) -> list[str]:
    anio = datos.get("fecha_entrega", date.today()).year
    nombre = datos.get("nombre_proyecto", "el proyecto")

    return [
        (
            f"Red Tecnoparque SENA. ({anio}). Estado del Arte del proyecto "
            f"{nombre}. Documento de planeación del proyecto."
        ),
        (
            "Las fuentes técnicas, académicas y bibliográficas completas "
            "consultadas para el análisis del área tecnológica se encuentran "
            "relacionadas en el documento Estado del Arte de la fase de planeación."
        ),
    ]


def anexos_desde_archivos(datos: dict) -> list[str]:
    archivos = datos.get("archivos_anexos", []) or []

    if archivos:
        return [
            f"Archivo de soporte: {archivo.get('nombre_original', '')}"
            for archivo in archivos
        ]

    return [
        "No se adjuntaron archivos de soporte al momento de generar el informe."
    ]


def _parrafos_complementarios(clave: str, datos: dict) -> list[str]:
    descripcion = datos.get("descripcion_general_proyecto", "")
    metodologias = metodologias_en_texto(datos)
    resultados = datos.get("resultados_obtenidos_base", "")
    impacto = datos.get("impacto_proyecto_base", "")

    nombres = {
        "introduccion": "la introducción",
        "planteamiento_problema": "el planteamiento del problema",
        "estado_arte_tecnica": "el estado del arte y de la técnica",
        "metodologia_desarrollo": "la metodología de desarrollo",
        "desarrollo_proyecto": "el desarrollo del proyecto",
        "resultados_obtenidos": "los resultados obtenidos",
        "analisis_viabilidad": "el análisis de viabilidad",
        "propiedad_transferencia": (
            "la propiedad intelectual y la transferencia tecnológica"
        ),
        "impacto_proyecto": "el impacto del proyecto",
        "conclusiones": "las conclusiones",
    }
    nombre_apartado = nombres.get(clave, "el apartado")

    return [
        (
            f"Para consolidar {nombre_apartado}, la información general suministrada "
            f"se analiza como una descripción integral del proceso: {descripcion} "
            "Este análisis relaciona la necesidad atendida, las decisiones adoptadas, "
            "los recursos disponibles y el alcance real del trabajo, sin incorporar "
            "cifras, validaciones o hechos que no hayan sido reportados."
        ),
        (
            f"El enfoque metodológico declarado —{metodologias or 'no especificado'}— "
            "permite organizar la experiencia en etapas comprensibles, mantener la "
            "trazabilidad de las actividades y explicar la relación entre el problema, "
            "el diseño de la solución, la implementación y la revisión de los avances. "
            "La redacción distingue entre resultados comprobados y oportunidades que "
            "todavía requieren validación adicional."
        ),
        (
            f"Los resultados informados fueron los siguientes: {resultados} "
            "Estos elementos se interpretan en función de los objetivos, el nivel de "
            "madurez tecnológica y las condiciones de aplicación descritas. El análisis "
            "evita atribuir desempeños, aprobaciones o beneficios cuantificados que no "
            "se encuentren expresamente respaldados por la información entregada."
        ),
        (
            f"El impacto reportado se resume así: {impacto} A partir de este insumo se "
            "establece la contribución del proyecto para sus usuarios y para el entorno "
            "de aplicación, considerando también las posibilidades de continuidad, "
            "apropiación, sostenibilidad y escalamiento. Las afirmaciones se presentan "
            "como resultados observados o como análisis técnico, según corresponda."
        ),
    ]


def asegurar_minimo_palabras(
    texto: str,
    clave: str,
    datos: dict,
    minimo: int = 410,
) -> str:
    resultado = limpiar_texto(texto)
    complementos = _parrafos_complementarios(clave, datos)
    indice = 0

    while len(resultado.split()) < minimo:
        resultado = f"{resultado}\n\n{complementos[indice % len(complementos)]}"
        indice += 1

    return resultado.strip()


def asegurar_nota_estado_arte(texto: str) -> str:
    nota = (
        "El Estado del Arte completo puede consultarse en la documentación "
        "de planeación del proyecto, específicamente en el documento Estado del Arte."
    )

    if "documento estado del arte" not in texto.casefold():
        return f"{texto.strip()}\n\n{nota}"

    return texto.strip()


def contenido_modo_prueba(datos: dict) -> dict:
    descripcion = datos.get("descripcion_general_proyecto", "")
    metodologias = metodologias_en_texto(datos)
    resultados = datos.get("resultados_obtenidos_base", "")
    impacto = datos.get("impacto_proyecto_base", "")

    base = {
        "introduccion": (
            f"El presente informe final documenta el proyecto de base tecnológica "
            f"{datos.get('nombre_proyecto', '')}, identificado con el código "
            f"{datos.get('codigo_proyecto', '')}. La iniciativa se desarrolló en "
            f"{datos.get('tecnoparque', '')}, con acompañamiento del experto "
            f"{datos.get('nombre_experto', '')}. La descripción general suministrada "
            f"establece el siguiente contexto: {descripcion} El informe organiza esta "
            "información para explicar la necesidad atendida, el proceso de desarrollo, "
            "los resultados alcanzados y la contribución de la solución."
        ),
        "planteamiento_problema": (
            f"La problemática se interpreta a partir de la descripción general del "
            f"proyecto: {descripcion} El análisis identifica la brecha existente entre "
            "la situación inicial y la condición que se esperaba alcanzar mediante la "
            "solución tecnológica. También considera a los usuarios involucrados, las "
            "limitaciones del contexto y las consecuencias de mantener el problema sin "
            "una intervención estructurada."
        ),
        "objetivo_general": (
            "Desarrollar y validar una solución de base tecnológica que responda a la "
            "necesidad identificada, mediante un proceso metodológico de diseño, "
            "implementación, evaluación y documentación de resultados."
        ),
        "objetivos_especificos": [
            "Caracterizar la necesidad, los usuarios, los requerimientos y las condiciones de aplicación de la solución.",
            "Diseñar los componentes, procesos y criterios técnicos necesarios para materializar la propuesta tecnológica.",
            "Implementar e integrar la solución mediante las metodologías y herramientas seleccionadas para el proyecto.",
            "Validar los resultados obtenidos, documentar las evidencias y establecer oportunidades de mejora y continuidad.",
        ],
        "estado_arte_tecnica": (
            f"El área tecnológica relacionada con el proyecto presenta una evolución "
            "orientada a soluciones más integradas, accesibles, modulares y centradas "
            "en las necesidades de los usuarios. A partir de la descripción del proyecto "
            f"—{descripcion}— se reconocen avances asociados con la digitalización, el "
            "prototipado iterativo, la interoperabilidad, la automatización, el análisis "
            "de información y la validación temprana de soluciones. Este apartado ofrece "
            "una síntesis general y no sustituye la revisión documental completa."
        ),
        "metodologia_desarrollo": (
            f"El desarrollo se estructuró con base en las metodologías seleccionadas: "
            f"{metodologias}. Estas se utilizaron para organizar la comprensión de la "
            "necesidad, la definición de requerimientos, el diseño de alternativas, la "
            "construcción del prototipo, la revisión de avances y el ajuste de la solución. "
            "La metodología se describe como un proceso trazable y adaptable al tipo de proyecto."
        ),
        "desarrollo_proyecto": (
            f"La ejecución se reconstruye a partir de la descripción general: {descripcion} "
            f"y del enfoque metodológico seleccionado: {metodologias}. El proceso se "
            "organizó en actividades de análisis, diseño, preparación de recursos, "
            "implementación, integración, revisión y documentación, procurando que cada "
            "decisión guardara relación con la necesidad y con los resultados esperados."
        ),
        "resultados_obtenidos": (
            f"Los resultados reportados por el usuario fueron: {resultados} Este apartado "
            "los presenta como productos, avances, prototipos, desarrollos, validaciones "
            "o evidencias de acuerdo con su naturaleza, y explica su relación con el "
            "cumplimiento de los objetivos y con la problemática que originó el proyecto."
        ),
        "analisis_viabilidad": (
            f"La viabilidad se analiza a partir de la descripción del proyecto, los "
            f"resultados reportados y el impacto esperado. La información base es: "
            f"{descripcion} Resultados: {resultados} Impacto: {impacto} Se consideran "
            "las condiciones técnicas, operativas, económicas, normativas y de adopción, "
            "sin afirmar costos, permisos o comportamientos de mercado no suministrados."
        ),
        "propiedad_transferencia": (
            f"A partir de la naturaleza del proyecto descrita en {descripcion}, se analizan "
            "los posibles activos de propiedad intelectual, tales como software, diseños, "
            "modelos, documentación, prototipos, contenidos o conocimiento técnico. Este "
            "análisis identifica potenciales mecanismos de protección y transferencia, "
            "sin afirmar que exista un registro, licencia o derecho concedido."
        ),
        "impacto_proyecto": (
            f"El impacto reportado por el usuario fue: {impacto} Este apartado organiza "
            "dicha información en dimensiones tecnológicas, sociales, económicas, "
            "ambientales o productivas, según resulte aplicable, y explica la forma en "
            "que los resultados aportan valor a usuarios, beneficiarios y procesos del entorno."
        ),
        "conclusiones": (
            f"El cierre del proyecto integra la descripción general —{descripcion}—, los "
            f"resultados obtenidos —{resultados}— y el impacto identificado —{impacto}—. "
            "Las conclusiones valoran la pertinencia de la solución, el aprendizaje del "
            "proceso, el nivel de madurez alcanzado y las oportunidades de mejora, "
            "continuidad, validación adicional o escalamiento."
        ),
        "referencias_bibliograficas": referencias_internas_proyecto(datos),
        "anexos": anexos_desde_archivos(datos),
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
        base[clave] = asegurar_minimo_palabras(base[clave], clave, datos)

    base["estado_arte_tecnica"] = asegurar_nota_estado_arte(
        base["estado_arte_tecnica"]
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

    objetivos = resultado["objetivos_especificos"]
    objetivos_respaldo = respaldo["objetivos_especificos"]

    for objetivo in objetivos_respaldo:
        if len(objetivos) >= 4:
            break
        if objetivo not in objetivos:
            objetivos.append(objetivo)

    resultado["objetivos_especificos"] = objetivos[:4]
    resultado["referencias_bibliograficas"] = referencias_internas_proyecto(datos)
    resultado["anexos"] = anexos_desde_archivos(datos)

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
        resultado[clave] = asegurar_minimo_palabras(
            resultado[clave],
            clave,
            datos,
        )

    resultado["estado_arte_tecnica"] = asegurar_nota_estado_arte(
        resultado["estado_arte_tecnica"]
    )

    return resultado


def generar_contenido_con_ia(
    datos: dict,
    modelo_openai: str,
) -> dict:
    respaldo = contenido_modo_prueba(datos)
    metodologias = metodologias_en_texto(datos)

    reglas_comunes = """
Actúa como redactor técnico senior de proyectos de base tecnológica de la Red
TecnoParque SENA. Redacta un Informe Final institucional en español formal,
coherente, técnico y verificable.

REGLAS OBLIGATORIAS
- Usa exclusivamente la descripción general, las metodologías seleccionadas,
  los resultados reportados, el impacto informado y los datos generales.
- No inventes nombres, cifras, costos, porcentajes, pruebas, validaciones,
  certificaciones, normas, patentes, registros, clientes, ventas ni referencias.
- Cuando una condición no haya sido comprobada, utiliza lenguaje analítico o
  condicional, sin presentarla como un hecho.
- Redacta como informe de cierre, no como propuesta futura.
- Cada apartado narrativo solicitado debe contener entre 430 y 600 palabras.
- Evita repetir literalmente la misma información entre apartados.
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

METODOLOGÍAS UTILIZADAS
{metodologias}

RESULTADOS OBTENIDOS REPORTADOS
{datos.get('resultados_obtenidos_base', '')}

IMPACTO DEL PROYECTO REPORTADO
{datos.get('impacto_proyecto_base', '')}
"""

    instrucciones_bloque_1 = reglas_comunes + """
Genera los apartados 2, 3, 4, 5 y 6.

REQUISITOS ESPECÍFICOS
- Introducción: contexto, propósito, entorno, actores y alcance.
- Planteamiento del problema: necesidad, causas, consecuencias, usuarios
  afectados y justificación, derivándolo de la descripción general.
- Objetivo general: una oración con verbo en infinitivo.
- Objetivos específicos: exactamente cuatro, cada uno con verbo en infinitivo.
- Estado del arte y estado de la técnica: presenta únicamente un resumen de los
  principales avances tecnológicos del área del proyecto. No inventes autores ni
  referencias. Incluye al final esta frase exacta: "El Estado del Arte completo
  puede consultarse en la documentación de planeación del proyecto,
  específicamente en el documento Estado del Arte."
- Metodología: explica de manera coherente cómo se aplicaron las metodologías
  seleccionadas al análisis, diseño, desarrollo, prototipado, validación y ajuste.
"""

    entrada_bloque_1 = contexto + """

ESTRUCTURA JSON OBLIGATORIA
{
  "introduccion": "430 a 600 palabras",
  "planteamiento_problema": "430 a 600 palabras",
  "objetivo_general": "una oración",
  "objetivos_especificos": ["objetivo 1", "objetivo 2", "objetivo 3", "objetivo 4"],
  "estado_arte_tecnica": "430 a 600 palabras",
  "metodologia_desarrollo": "430 a 600 palabras"
}
"""

    instrucciones_bloque_2 = reglas_comunes + """
Genera los apartados 7, 8, 9, 10, 11 y 12.

REQUISITOS ESPECÍFICOS
- Desarrollo: reconstruye las actividades, fases, componentes y decisiones
  técnicas a partir de la descripción y de las metodologías seleccionadas.
- Resultados: utiliza como fuente principal los resultados reportados; explica
  su relación con los objetivos sin inventar pruebas o desempeños.
- Viabilidad: analiza dimensiones técnica, operativa, económica, normativa y
  de adopción; usa lenguaje condicional cuando falten datos comprobables.
- Propiedad intelectual y transferencia: identifica activos potenciales y
  estrategias posibles; no afirmes que existen registros concedidos.
- Impacto: utiliza como fuente principal el impacto reportado y organízalo en
  dimensiones aplicables al proyecto.
- Conclusiones: integra pertinencia, resultados, TRL, aprendizajes, mejoras,
  continuidad y escalamiento sin introducir hechos nuevos.
"""

    entrada_bloque_2 = contexto + """

ESTRUCTURA JSON OBLIGATORIA
{
  "desarrollo_proyecto": "430 a 600 palabras",
  "resultados_obtenidos": "430 a 600 palabras",
  "analisis_viabilidad": "430 a 600 palabras",
  "propiedad_transferencia": "430 a 600 palabras",
  "impacto_proyecto": "430 a 600 palabras",
  "conclusiones": "430 a 600 palabras"
}
"""

    bloque_1 = generar_json_openai(
        instrucciones=instrucciones_bloque_1,
        entrada=entrada_bloque_1,
        modelo=modelo_openai,
        temperature=0.2,
    )

    bloque_2 = generar_json_openai(
        instrucciones=instrucciones_bloque_2,
        entrada=entrada_bloque_2,
        modelo=modelo_openai,
        temperature=0.2,
    )

    contenido: dict = {}
    if isinstance(bloque_1, dict):
        contenido.update(bloque_1)
    if isinstance(bloque_2, dict):
        contenido.update(bloque_2)

    contenido["referencias_bibliograficas"] = referencias_internas_proyecto(datos)
    contenido["anexos"] = anexos_desde_archivos(datos)

    return normalizar_contenido(contenido, respaldo, datos)


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

    # =========================================================
# PROPIEDADES INTERNAS DEL DOCUMENTO WORD
# Se usa el código del proyecto para evitar errores por más de 255 caracteres
# =========================================================
# =========================================================
# PROPIEDADES INTERNAS DEL DOCUMENTO WORD
# Se usa el código del proyecto para evitar errores por más de 255 caracteres
# =========================================================

    def limpiar_texto_propiedad_word(texto, limite=250):
        texto = str(texto or "").replace("\n", " ").replace("\r", " ").strip()
        if len(texto) > limite:
            return texto[:limite - 3] + "..."
        return texto


    codigo_proyecto = (
        datos.get("codigo_proyecto")
        or datos.get("codigo")
        or datos.get("codigo_tecnoparque")
        or datos.get("codigo_sennova")
        or datos.get("id_proyecto")
        or "SIN_CODIGO"
    )

    codigo_proyecto = str(codigo_proyecto).strip()

    if not codigo_proyecto:
        codigo_proyecto = "SIN_CODIGO"


    documento.core_properties.title = limpiar_texto_propiedad_word(
        f"Informe Final - {codigo_proyecto}"
    )

    documento.core_properties.subject = "Informe Final Tecnoparque"
    documento.core_properties.author = "Tecnoparque"
    documento.core_properties.keywords = "SENA, TecnoParque, Informe Final, GCDTP-F-023 V01"

    

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
        "El documento se genera sobre la plantilla oficial GCDTP-F-023 V01. "
        "El formulario solicita únicamente los datos institucionales y cuatro "
        "insumos de contenido: descripción general, metodología, resultados e "
        "impacto. La API redacta los demás apartados con más de 400 palabras. "
        "Las páginas de instrucciones y control de cambios se eliminan del Word final."
    )

    if "datos_informe_tecnico_final_generado" not in st.session_state:
        st.session_state.datos_informe_tecnico_final_generado = None

    if "ruta_docx_informe_tecnico_final_generado" not in st.session_state:
        st.session_state.ruta_docx_informe_tecnico_final_generado = None

    with st.form("form_informe_tecnico_final_gcdtp_023_simplificado"):
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

        st.markdown("## Información para generar el contenido")

        descripcion_general_proyecto = st.text_area(
            "Descripción general del proyecto",
            placeholder=(
                "Describe ampliamente el origen del proyecto, la necesidad, los usuarios, "
                "el propósito, los componentes, tecnologías utilizadas, actividades "
                "realizadas, funcionamiento de la solución y contexto de aplicación."
            ),
            height=340,
        )

        metodologias_seleccionadas = st.multiselect(
            "Metodología o metodologías utilizadas",
            options=METODOLOGIAS_DESARROLLO,
            help=(
                "Puedes seleccionar varias metodologías. Para una metodología no incluida, "
                "selecciona Otra y escríbela en el campo siguiente."
            ),
        )

        otra_metodologia = st.text_input(
            "Otra metodología utilizada",
            placeholder=(
                "Escribe aquí una metodología adicional o una combinación propia."
            ),
        )

        resultados_obtenidos_base = st.text_area(
            "Resultados obtenidos",
            placeholder=(
                "Indica los prototipos, productos, desarrollos, entregables, pruebas, "
                "validaciones y logros realmente obtenidos. No incluyas resultados esperados."
            ),
            height=220,
        )

        impacto_proyecto_base = st.text_area(
            "Impacto del proyecto",
            placeholder=(
                "Describe los beneficios y efectos tecnológicos, sociales, económicos, "
                "ambientales o productivos, los beneficiarios y el valor generado."
            ),
            height=220,
        )

        archivos_anexos_upload = st.file_uploader(
            "Anexos y evidencias opcionales",
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
                "Las imágenes se insertan en el documento. Los demás archivos se "
                "relacionan por nombre en la sección de anexos."
            ),
        )

        generar_contenido = st.form_submit_button(
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
            "Resultados obtenidos": resultados_obtenidos_base,
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
                "Selecciona al menos una metodología o escribe una en el campo Otra metodología."
            )
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
            "metodologias_seleccionadas": metodologias_seleccionadas,
            "otra_metodologia": limpiar_texto(otra_metodologia),
            "resultados_obtenidos_base": limpiar_texto(
                resultados_obtenidos_base
            ),
            "impacto_proyecto_base": limpiar_texto(
                impacto_proyecto_base
            ),
            "archivos_anexos": archivos_anexos,
            "modo_generacion": (
                "Prueba local"
                if modo_prueba
                else "API de OpenAI"
            ),
            "version": VERSION_INFORME_TECNICO_FINAL,
        }

        with st.spinner(
            "Generando los apartados del Informe Final. Este proceso puede tardar "
            "porque cada sección narrativa tendrá más de 400 palabras."
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

        st.session_state.datos_informe_tecnico_final_generado = datos_base
        st.session_state.ruta_docx_informe_tecnico_final_generado = None

        st.success(
            "Contenido generado. Revisa y edita los apartados antes de crear el Word oficial."
        )

    datos = st.session_state.get(
        "datos_informe_tecnico_final_generado"
    )

    if datos:
        contenido = datos["contenido_informe"]

        st.markdown("## Revisión y edición del contenido")
        st.caption(
            "Los apartados narrativos se generan con un mínimo de 400 palabras. "
            "Guarda cualquier cambio antes de crear el archivo Word."
        )

        with st.form("form_revisar_informe_tecnico_final"):
            introduccion_editada = st.text_area(
                "2. Introducción",
                value=contenido["introduccion"],
                height=360,
            )

            problema_editado = st.text_area(
                "3. Planteamiento del problema",
                value=contenido["planteamiento_problema"],
                height=360,
            )

            objetivo_general_editado = st.text_area(
                "4.1 Objetivo General",
                value=contenido["objetivo_general"],
                height=100,
            )

            objetivos_editados = st.text_area(
                "4.2 Objetivos Específicos",
                value="\n".join(contenido["objetivos_especificos"]),
                height=170,
            )

            estado_arte_editado = st.text_area(
                "5. Estado del arte y estado de la técnica",
                value=contenido["estado_arte_tecnica"],
                height=380,
            )

            metodologia_editada = st.text_area(
                "6. Metodología de desarrollo",
                value=contenido["metodologia_desarrollo"],
                height=380,
            )

            desarrollo_editado = st.text_area(
                "7. Desarrollo del proyecto",
                value=contenido["desarrollo_proyecto"],
                height=400,
            )

            resultados_editados = st.text_area(
                "8. Resultados obtenidos",
                value=contenido["resultados_obtenidos"],
                height=380,
            )

            viabilidad_editada = st.text_area(
                "9. Análisis de viabilidad",
                value=contenido["analisis_viabilidad"],
                height=380,
            )

            propiedad_editada = st.text_area(
                "10. Propiedad intelectual y transferencia tecnológica",
                value=contenido["propiedad_transferencia"],
                height=380,
            )

            impacto_editado = st.text_area(
                "11. Impacto del proyecto",
                value=contenido["impacto_proyecto"],
                height=380,
            )

            conclusiones_editadas = st.text_area(
                "12. Conclusiones",
                value=contenido["conclusiones"],
                height=380,
            )

            referencias_editadas = st.text_area(
                "13. Referencias bibliográficas",
                value="\n".join(contenido["referencias_bibliograficas"]),
                height=150,
            )

            anexos_editados = st.text_area(
                "14. Anexos",
                value="\n".join(contenido["anexos"]),
                height=150,
            )

            guardar_edicion = st.form_submit_button(
                "💾 Guardar cambios del informe"
            )

        if guardar_edicion:
            datos["contenido_informe"] = {
                "introduccion": limpiar_texto(introduccion_editada),
                "planteamiento_problema": limpiar_texto(problema_editado),
                "objetivo_general": limpiar_texto(objetivo_general_editado),
                "objetivos_especificos": dividir_lineas(objetivos_editados)[:4],
                "estado_arte_tecnica": asegurar_nota_estado_arte(
                    limpiar_texto(estado_arte_editado)
                ),
                "metodologia_desarrollo": limpiar_texto(metodologia_editada),
                "desarrollo_proyecto": limpiar_texto(desarrollo_editado),
                "resultados_obtenidos": limpiar_texto(resultados_editados),
                "analisis_viabilidad": limpiar_texto(viabilidad_editada),
                "propiedad_transferencia": limpiar_texto(propiedad_editada),
                "impacto_proyecto": limpiar_texto(impacto_editado),
                "conclusiones": limpiar_texto(conclusiones_editadas),
                "referencias_bibliograficas": dividir_lineas(
                    referencias_editadas
                ),
                "anexos": dividir_lineas(anexos_editados),
            }

            st.session_state.datos_informe_tecnico_final_generado = datos
            st.session_state.ruta_docx_informe_tecnico_final_generado = None

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
                    ruta_docx = generar_docx_informe_tecnico_final(datos)
                    st.session_state.ruta_docx_informe_tecnico_final_generado = ruta_docx
                    st.success("Documento Word generado correctamente.")
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

