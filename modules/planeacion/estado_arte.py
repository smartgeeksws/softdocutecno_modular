from __future__ import annotations

from datetime import date, datetime
from html import escape
from pathlib import Path
import json
import os
import re

import streamlit as st

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from config.constants import OUTPUT_DIR, RUTA_LOGO_SENA
from services.json_service import guardar_datos_json
from services.openai_service import generar_json_openai
from utils.nombres_archivo import safe_filename
from utils.textos import limpiar_texto
from utils.validaciones import validar_campos_obligatorios


VERSION_ESTADO_ARTE = "VERSION_MODULAR_ESTADO_ARTE_ACADEMICO_APA7_VALIDADO"
FORMATO_ESTADO_ARTE = "ESTADO DEL ARTE - RED TECNOPARQUE SENA"


# =====================================================
# UTILIDADES
# =====================================================

def limpiar_lista_tecnologias(texto: str) -> list[str]:
    """Convierte texto separado por comas, punto y coma o saltos en una lista."""
    if not texto:
        return []

    tecnologias = [str(texto)]

    for separador in ["\n", ";", ","]:
        nuevas: list[str] = []
        for item in tecnologias:
            nuevas.extend(item.split(separador))
        tecnologias = nuevas

    resultado: list[str] = []
    vistos: set[str] = set()

    for tecnologia in tecnologias:
        tecnologia_limpia = limpiar_texto(tecnologia)

        if not tecnologia_limpia:
            continue

        clave = tecnologia_limpia.casefold()
        if clave not in vistos:
            vistos.add(clave)
            resultado.append(tecnologia_limpia)

    return resultado


def obtener_api_key() -> str | None:
    api_key = os.getenv("OPENAI_API_KEY")
    if api_key:
        return api_key

    try:
        valor = st.secrets.get("OPENAI_API_KEY")
        return str(valor).strip() if valor else None
    except Exception:
        return None


def obtener_ruta_logo_tecnoparque() -> str | None:
    candidatos = [
        Path(RUTA_LOGO_SENA),
        Path("resources/logo_tecnoparque.png"),
        Path("resources/logo_tecnoparque.jpg"),
        Path("resources/logo_sena.png"),
        Path("recursos/logo_tecnoparque.png"),
        Path("recursos/logo_sena.png"),
    ]

    for ruta in candidatos:
        if ruta.exists():
            return str(ruta)

    return None


def normalizar_enlace(enlace: str) -> str:
    enlace = limpiar_texto(enlace)

    if not enlace:
        return ""

    if enlace.startswith(("http://", "https://")):
        return enlace

    if enlace.startswith("www."):
        return f"https://{enlace}"

    return enlace


def _texto_seguro(valor: object) -> str:
    return escape(str(valor or "").strip())


def _parrafo(
    texto: object,
    estilo: ParagraphStyle,
    permitir_enlace: bool = False,
) -> Paragraph:
    texto_limpio = str(texto or "").strip()

    if permitir_enlace:
        enlace = normalizar_enlace(texto_limpio)
        if enlace.startswith(("http://", "https://")):
            contenido = (
                f'<link href="{escape(enlace, quote=True)}" '
                f'color="#1A5FB4">{escape(enlace)}</link>'
            )
            return Paragraph(contenido, estilo)

    return Paragraph(escape(texto_limpio).replace("\n", "<br/>"), estilo)


def _serializar_datos_estado_arte(datos: dict) -> dict:
    serializado = dict(datos)
    fecha_documento = serializado.get("fecha_documento")

    if isinstance(fecha_documento, date):
        serializado["fecha_documento"] = fecha_documento.strftime("%d/%m/%Y")

    return serializado


# =====================================================
# CONTENIDO EN MODO PRUEBA
# =====================================================

def generar_estado_arte_modo_prueba(
    nombre_proyecto: str,
    codigo_proyecto: str,
    descripcion_proyecto: str,
    tecnologias_previstas: list[str],
) -> dict:
    tecnologias_base = tecnologias_previstas or [
        "Aplicaciones web",
        "Bases de datos",
        "Analítica de información",
        "Internet de las cosas",
    ]

    antecedentes = (
        f"El proyecto identificado con el código {codigo_proyecto} se orienta al desarrollo "
        f"de una solución de base tecnológica relacionada con la iniciativa denominada "
        f"{nombre_proyecto}. La descripción suministrada señala como punto de partida la "
        f"siguiente necesidad y oportunidad de innovación: {descripcion_proyecto} "
        "Para establecer sus antecedentes es necesario considerar la evolución reciente de "
        "las soluciones digitales, los procesos de automatización, el diseño centrado en el "
        "usuario y la integración de datos como elementos que permiten transformar una idea "
        "en un prototipo verificable. En proyectos de esta naturaleza, el Estado del Arte "
        "cumple la función de identificar enfoques ya utilizados, reconocer limitaciones "
        "técnicas y evitar que el proceso de diseño repita soluciones sin una justificación "
        "suficiente. La revisión debe comparar productos, investigaciones, plataformas y "
        "experiencias que respondan a necesidades semejantes, diferenciando los desarrollos "
        "académicos de las soluciones comerciales y de las iniciativas institucionales. "
        "También debe revisar las condiciones del entorno de aplicación, los perfiles de "
        "usuario, la disponibilidad de infraestructura y la posibilidad de validar la "
        "solución en escenarios reales. Desde la perspectiva metodológica, los antecedentes "
        "deben relacionarse con procesos iterativos de análisis, diseño, prototipado, prueba "
        "y ajuste. Esta secuencia permite documentar decisiones y establecer criterios de "
        "éxito asociados con funcionalidad, usabilidad, seguridad, interoperabilidad, "
        "mantenibilidad y sostenibilidad. Las tecnologías previstas deben evaluarse no solo "
        "por su novedad, sino por su pertinencia para resolver la necesidad descrita, su "
        "madurez, sus costos de implementación y la capacidad del equipo para integrarlas. "
        "El análisis también debe considerar aspectos éticos, protección de datos, propiedad "
        "intelectual, accesibilidad y uso responsable de la información cuando sean "
        "aplicables. En el contexto de la Red Tecnoparque SENA, este ejercicio facilita la "
        "formulación técnica del proyecto, fortalece la trazabilidad de sus decisiones y "
        "orienta la construcción de un prototipo que pueda ser sometido a pruebas. Por ello, "
        "el documento debe actualizarse con fuentes verificables y referencias académicas "
        "reales antes de ser utilizado como soporte definitivo del proyecto."
    )

    proyectos_similares = [
        {
            "numero": 1,
            "nombre": "Plataforma de innovación tecnológica aplicada",
            "enlace": "https://scholar.google.com/",
            "descripcion_breve": (
                "Referente de demostración para revisar proyectos académicos que integran "
                "componentes digitales, gestión de información y validación con usuarios."
            ),
            "referencia_apa": (
                "Referencia de demostración. Debe reemplazarse por una fuente académica "
                "verificada antes de utilizar el documento."
            ),
        },
        {
            "numero": 2,
            "nombre": "Sistema de apoyo a procesos mediante tecnologías emergentes",
            "enlace": "https://www.sciencedirect.com/",
            "descripcion_breve": (
                "Ejemplo de búsqueda para identificar sistemas que emplean automatización, "
                "analítica o servicios digitales en contextos productivos y sociales."
            ),
            "referencia_apa": (
                "Referencia de demostración. Consultar una publicación verificable en la "
                "base de datos indicada."
            ),
        },
        {
            "numero": 3,
            "nombre": "Prototipo tecnológico con diseño centrado en el usuario",
            "enlace": "https://ieeexplore.ieee.org/",
            "descripcion_breve": (
                "Referente orientado al análisis de prototipos que incorporan pruebas de "
                "usabilidad, iteración y validación funcional."
            ),
            "referencia_apa": (
                "Referencia de demostración. Verificar autor, año, título, publicación y DOI."
            ),
        },
        {
            "numero": 4,
            "nombre": "Solución interoperable para captura y visualización de datos",
            "enlace": "https://dl.acm.org/",
            "descripcion_breve": (
                "Ejemplo para comparar arquitecturas, integración de datos, interfaces y "
                "mecanismos de acceso a la información."
            ),
            "referencia_apa": (
                "Referencia de demostración. Reemplazar por un artículo pertinente al proyecto."
            ),
        },
        {
            "numero": 5,
            "nombre": "Iniciativa de transformación digital para organizaciones",
            "enlace": "https://repositorio.sena.edu.co/",
            "descripcion_breve": (
                "Punto de consulta para localizar experiencias nacionales, proyectos "
                "formativos y documentos técnicos relacionados con innovación aplicada."
            ),
            "referencia_apa": (
                "Referencia de demostración. Completar con los datos de una fuente real."
            ),
        },
    ]

    tecnologias_relevantes = [
        {
            "tecnologia": tecnologia,
            "analisis": (
                f"{tecnologia} se considera relevante porque puede aportar capacidades para "
                "el diseño, implementación, integración, prueba o escalamiento de la solución. "
                "Su selección debe sustentarse en requerimientos verificables, compatibilidad "
                "con los demás componentes y condiciones reales de uso."
            ),
            "cita_apa": (
                "Referencia técnica de demostración; debe verificarse y reemplazarse por una "
                "fuente académica o documentación oficial."
            ),
        }
        for tecnologia in tecnologias_base
    ]

    tecnologias_emergentes = [
        {
            "tecnologia": "Inteligencia artificial generativa",
            "analisis": (
                "Puede apoyar la creación, clasificación o síntesis de contenidos y asistir "
                "procesos de análisis, siempre que se incorporen validación humana, control "
                "de calidad y protección de datos."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
        {
            "tecnologia": "Internet de las cosas",
            "analisis": (
                "Permite conectar sensores, dispositivos y servicios para recopilar datos, "
                "automatizar acciones y supervisar variables en escenarios físicos."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
        {
            "tecnologia": "Analítica de datos",
            "analisis": (
                "Facilita la interpretación de registros, identificación de patrones y "
                "construcción de indicadores para apoyar decisiones técnicas."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
        {
            "tecnologia": "Modelado y simulación digital",
            "analisis": (
                "Permite evaluar componentes o procesos antes de su implementación física, "
                "reduciendo riesgos y mejorando la planeación de pruebas."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
        {
            "tecnologia": "Computación en la nube",
            "analisis": (
                "Aporta recursos de almacenamiento, procesamiento y despliegue escalable para "
                "soluciones que requieren acceso remoto y disponibilidad."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
        {
            "tecnologia": "Realidad aumentada",
            "analisis": (
                "Puede enriquecer experiencias de aprendizaje, asistencia, divulgación o "
                "visualización mediante la superposición de información digital."
            ),
            "cita_apa": "Referencia académica pendiente de verificación.",
        },
    ]

    articulos_validacion = [
        {
            "numero": indice,
            "tecnologia": tecnologia["tecnologia"],
            "articulo": (
                f"Fuente técnica de validación para {tecnologia['tecnologia']} "
                "(registro de demostración)"
            ),
            "enlace": "https://scholar.google.com/",
            "referencia_apa": (
                "Referencia de demostración. Buscar y registrar una fuente real en formato APA 7."
            ),
        }
        for indice, tecnologia in enumerate(
            (tecnologias_relevantes + tecnologias_emergentes)[:5],
            start=1,
        )
    ]

    return {
        "introduccion": (
            "El presente Estado del Arte establece una base conceptual, tecnológica y "
            "documental para orientar la planeación del proyecto. La revisión identifica "
            "antecedentes, referentes comparables, tecnologías previstas, tendencias "
            "emergentes y fuentes de validación que deben contrastarse durante el diseño "
            "y desarrollo del prototipo."
        ),
        "objetivos": [
            "Identificar antecedentes técnicos, académicos y tecnológicos relacionados con la iniciativa.",
            "Analizar referentes nacionales e internacionales que permitan comparar enfoques, soluciones y resultados.",
            "Examinar la pertinencia de las tecnologías previstas y reconocer tecnologías emergentes aplicables.",
            "Consolidar criterios documentales para orientar las decisiones técnicas y la validación del proyecto.",
        ],
        "antecedentes_contexto": antecedentes,
        "contexto_nacional_internacional": (
            "En el ámbito nacional se observa una adopción progresiva de tecnologías "
            "digitales en emprendimientos, instituciones educativas, organizaciones "
            "productivas y entidades públicas. A escala internacional, la convergencia de "
            "servicios en la nube, inteligencia artificial, conectividad, automatización y "
            "analítica impulsa soluciones modulares e interoperables. La comparación entre "
            "ambos contextos permite identificar oportunidades de adaptación local, "
            "requisitos de infraestructura y criterios para validar la propuesta."
        ),
        "proyectos_similares": proyectos_similares,
        "tecnologias_relevantes": tecnologias_relevantes,
        "tecnologias_emergentes": tecnologias_emergentes,
        "articulos_validacion": articulos_validacion,
        "conclusiones": (
            "La revisión evidencia que la iniciativa puede fortalecerse mediante una "
            "selección tecnológica sustentada en requerimientos y fuentes verificables. "
            "Los referentes permiten reconocer alternativas de diseño, riesgos y criterios "
            "de prueba. Antes de aprobar el documento definitivo deben sustituirse las "
            "referencias de demostración por publicaciones reales y comprobar cada enlace."
        ),
        "bibliografia": [
            "Referencias de demostración para modo prueba. Deben reemplazarse por fuentes verificables en formato APA 7.",
            "Google Scholar. (s. f.). Buscador de literatura académica.",
            "Servicio Nacional de Aprendizaje. (s. f.). Repositorio institucional SENA.",
        ],
    }


# =====================================================
# GENERACIÓN CON IA Y BÚSQUEDA WEB
# =====================================================

def investigar_fuentes_con_openai(
    nombre_proyecto: str,
    codigo_proyecto: str,
    descripcion_proyecto: str,
    tecnologias_previstas: list[str],
    modelo: str,
) -> str:
    """Obtiene notas de investigación. Si web search no está disponible, retorna aviso."""
    if OpenAI is None:
        return (
            "No fue posible ejecutar búsqueda web porque no está instalada la librería "
            "openai. Generar el documento y verificar manualmente todas las referencias."
        )

    api_key = obtener_api_key()
    if not api_key:
        return (
            "No fue posible ejecutar búsqueda web porque no se encontró OPENAI_API_KEY. "
            "Verificar manualmente todas las referencias."
        )

    client = OpenAI(api_key=api_key)
    tecnologias_texto = ", ".join(tecnologias_previstas) or "No especificadas"

    prompt_busqueda = f"""
Investiga fuentes públicas, académicas, institucionales y técnicas para construir un
Estado del Arte profesional en español.

Proyecto: {nombre_proyecto}
Código: {codigo_proyecto}
Descripción: {descripcion_proyecto}
Tecnologías previstas: {tecnologias_texto}

Recopila notas verificables sobre:
1. Contexto nacional e internacional del sector y del problema.
2. Mínimo cinco proyectos o iniciativas similares, con nombre, entidad y enlace.
3. Tecnologías relevantes para la solución.
4. Cinco o seis tecnologías emergentes aplicables.
5. Mínimo cinco artículos, normas, documentos técnicos o fuentes académicas de validación.
6. Datos suficientes para referencias APA 7.

No inventes autores, títulos, fechas, DOI ni enlaces. Señala expresamente cualquier dato
que no haya podido verificarse.
"""

    try:
        respuesta = client.responses.create(
            model=modelo,
            tools=[{"type": "web_search_preview"}],
            input=prompt_busqueda,
            temperature=0.15,
        )
        notas = getattr(respuesta, "output_text", "") or ""
        return notas.strip() or "La búsqueda no produjo notas utilizables."
    except Exception as error:
        return (
            "La búsqueda web de la API no estuvo disponible. "
            f"Detalle técnico: {error}. Todas las referencias deben verificarse manualmente."
        )


def _normalizar_lista_diccionarios(
    valor: object,
    campos: list[str],
) -> list[dict]:
    if not isinstance(valor, list):
        return []

    resultado: list[dict] = []

    for item in valor:
        if not isinstance(item, dict):
            continue

        normalizado = {
            campo: item.get(campo, "")
            for campo in campos
        }
        resultado.append(normalizado)

    return resultado


def normalizar_contenido_estado_arte(
    datos: object,
    respaldo: dict,
    tecnologias_previstas: list[str],
) -> dict:
    if not isinstance(datos, dict):
        return respaldo

    contenido = dict(datos)

    campos_texto = [
        "introduccion",
        "antecedentes_contexto",
        "contexto_nacional_internacional",
        "conclusiones",
    ]

    for campo in campos_texto:
        if not isinstance(contenido.get(campo), str) or not contenido[campo].strip():
            contenido[campo] = respaldo[campo]

    objetivos = contenido.get("objetivos")
    if not isinstance(objetivos, list):
        objetivos = []

    objetivos = [
        limpiar_texto(objetivo)
        for objetivo in objetivos
        if limpiar_texto(objetivo)
    ]

    objetivos_respaldo = respaldo["objetivos"]
    contenido["objetivos"] = (objetivos + objetivos_respaldo)[:4]

    contenido["proyectos_similares"] = _normalizar_lista_diccionarios(
        contenido.get("proyectos_similares"),
        ["numero", "nombre", "enlace", "descripcion_breve", "referencia_apa"],
    )

    contenido["tecnologias_relevantes"] = _normalizar_lista_diccionarios(
        contenido.get("tecnologias_relevantes"),
        ["tecnologia", "analisis", "cita_apa"],
    )

    contenido["tecnologias_emergentes"] = _normalizar_lista_diccionarios(
        contenido.get("tecnologias_emergentes"),
        ["tecnologia", "analisis", "cita_apa"],
    )

    contenido["articulos_validacion"] = _normalizar_lista_diccionarios(
        contenido.get("articulos_validacion"),
        ["numero", "tecnologia", "articulo", "enlace", "referencia_apa"],
    )

    bibliografia = contenido.get("bibliografia")
    if not isinstance(bibliografia, list):
        bibliografia = []

    contenido["bibliografia"] = [
        limpiar_texto(referencia)
        for referencia in bibliografia
        if limpiar_texto(referencia)
    ]

    if len(contenido["proyectos_similares"]) < 5:
        faltantes = 5 - len(contenido["proyectos_similares"])
        contenido["proyectos_similares"].extend(
            respaldo["proyectos_similares"][:faltantes]
        )

    if not contenido["tecnologias_relevantes"]:
        contenido["tecnologias_relevantes"] = [
            {
                "tecnologia": tecnologia,
                "analisis": (
                    f"{tecnologia} debe evaluarse según los requerimientos técnicos, "
                    "la compatibilidad, los riesgos de implementación y las pruebas previstas."
                ),
                "cita_apa": "Referencia técnica pendiente de verificación.",
            }
            for tecnologia in tecnologias_previstas
        ] or respaldo["tecnologias_relevantes"]

    if len(contenido["tecnologias_emergentes"]) < 5:
        existentes = {
            str(item.get("tecnologia", "")).casefold()
            for item in contenido["tecnologias_emergentes"]
        }

        for item in respaldo["tecnologias_emergentes"]:
            if str(item.get("tecnologia", "")).casefold() not in existentes:
                contenido["tecnologias_emergentes"].append(item)
                existentes.add(str(item.get("tecnologia", "")).casefold())

            if len(contenido["tecnologias_emergentes"]) >= 5:
                break

    if len(contenido["articulos_validacion"]) < 5:
        faltantes = 5 - len(contenido["articulos_validacion"])
        contenido["articulos_validacion"].extend(
            respaldo["articulos_validacion"][:faltantes]
        )

    if not contenido["bibliografia"]:
        contenido["bibliografia"] = respaldo["bibliografia"]

    return contenido


def generar_estado_arte_con_chatgpt(
    nombre_proyecto: str,
    codigo_proyecto: str,
    descripcion_proyecto: str,
    tecnologias_previstas: list[str],
    modelo: str = "gpt-4.1-mini",
) -> dict:
    respaldo = generar_estado_arte_modo_prueba(
        nombre_proyecto,
        codigo_proyecto,
        descripcion_proyecto,
        tecnologias_previstas,
    )

    notas_investigacion = investigar_fuentes_con_openai(
        nombre_proyecto,
        codigo_proyecto,
        descripcion_proyecto,
        tecnologias_previstas,
        modelo,
    )

    tecnologias_texto = ", ".join(tecnologias_previstas) or "No especificadas"

    instrucciones = """
Actúa como investigador académico senior y formulador de proyectos de base tecnológica
para la Red Tecnoparque SENA.

Genera un Estado del Arte profesional, académico e investigativo en español.
Usa únicamente la información del proyecto y las notas de investigación suministradas.
No inventes autores, entidades, títulos, años, DOI, normas ni enlaces.
Cuando un dato no sea verificable, indícalo como pendiente de verificación.
Las referencias deben presentarse en APA 7.
No uses markdown.
Responde exclusivamente en JSON válido.

Requisitos:
- Introducción clara y técnica.
- Exactamente cuatro objetivos.
- Antecedentes y contexto con mínimo 500 palabras.
- Contexto nacional e internacional.
- Mínimo cinco proyectos similares.
- Incluir todas las tecnologías previstas por el usuario.
- Incluir cinco o seis tecnologías emergentes.
- Mínimo cinco artículos o fuentes técnicas de validación.
- Conclusiones y bibliografía consolidada.
"""

    entrada = f"""
DATOS DEL PROYECTO

Nombre: {nombre_proyecto}
Código: {codigo_proyecto}
Descripción detallada:
{descripcion_proyecto}

Tecnologías previstas:
{tecnologias_texto}

NOTAS DE INVESTIGACIÓN
{notas_investigacion}

ESTRUCTURA JSON OBLIGATORIA

{{
  "introduccion": "texto",
  "objetivos": ["objetivo 1", "objetivo 2", "objetivo 3", "objetivo 4"],
  "antecedentes_contexto": "texto de mínimo 500 palabras",
  "contexto_nacional_internacional": "texto",
  "proyectos_similares": [
    {{
      "numero": 1,
      "nombre": "nombre verificable",
      "enlace": "https://...",
      "descripcion_breve": "texto",
      "referencia_apa": "referencia APA 7"
    }}
  ],
  "tecnologias_relevantes": [
    {{
      "tecnologia": "nombre",
      "analisis": "texto",
      "cita_apa": "referencia APA 7"
    }}
  ],
  "tecnologias_emergentes": [
    {{
      "tecnologia": "nombre",
      "analisis": "texto",
      "cita_apa": "referencia APA 7"
    }}
  ],
  "articulos_validacion": [
    {{
      "numero": 1,
      "tecnologia": "tecnología relacionada",
      "articulo": "título o nombre de la fuente",
      "enlace": "https://...",
      "referencia_apa": "referencia APA 7"
    }}
  ],
  "conclusiones": "texto",
  "bibliografia": ["referencia APA 7"]
}}
"""

    datos = generar_json_openai(
        instrucciones=instrucciones,
        entrada=entrada,
        modelo=modelo,
        temperature=0.15,
    )

    return normalizar_contenido_estado_arte(
        datos,
        respaldo,
        tecnologias_previstas,
    )


# =====================================================
# GENERACIÓN DEL PDF
# =====================================================

def generar_pdf_estado_arte(datos: dict) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Estado_del_Arte_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(OUTPUT_DIR / nombre_archivo)

    page_width, page_height = letter
    contenido = datos.get("contenido_estado_arte", {}) or {}

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        leftMargin=2.0 * cm,
        rightMargin=2.0 * cm,
        topMargin=2.7 * cm,
        bottomMargin=1.8 * cm,
        title=f"Estado del Arte - {datos.get('nombre_proyecto', '')}",
        author="Red Tecnoparque SENA",
    )

    def encabezado_pie(canvas_pdf, doc_pdf) -> None:
        canvas_pdf.saveState()

        ruta_logo = obtener_ruta_logo_tecnoparque()
        if ruta_logo:
            try:
                imagen = ImageReader(ruta_logo)
                canvas_pdf.drawImage(
                    imagen,
                    1.5 * cm,
                    page_height - 2.05 * cm,
                    width=4.8 * cm,
                    height=1.15 * cm,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                pass

        canvas_pdf.setFillColor(colors.HexColor("#2E7D32"))
        canvas_pdf.setFont("Helvetica-Bold", 10)
        canvas_pdf.drawRightString(
            page_width - 1.5 * cm,
            page_height - 1.35 * cm,
            "ESTADO DEL ARTE",
        )

        canvas_pdf.setStrokeColor(colors.HexColor("#93C47D"))
        canvas_pdf.setLineWidth(0.6)
        canvas_pdf.line(
            1.5 * cm,
            page_height - 2.18 * cm,
            page_width - 1.5 * cm,
            page_height - 2.18 * cm,
        )

        canvas_pdf.setFillColor(colors.grey)
        canvas_pdf.setFont("Helvetica", 8)
        canvas_pdf.drawString(
            1.5 * cm,
            0.8 * cm,
            FORMATO_ESTADO_ARTE,
        )
        canvas_pdf.drawRightString(
            page_width - 1.5 * cm,
            0.8 * cm,
            f"Página {doc_pdf.page}",
        )

        canvas_pdf.restoreState()

    estilo_portada_titulo = ParagraphStyle(
        name="PortadaTituloEstadoArte",
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=21,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#2E7D32"),
        spaceAfter=18,
    )

    estilo_portada_proyecto = ParagraphStyle(
        name="PortadaProyectoEstadoArte",
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=17,
        alignment=TA_CENTER,
        spaceAfter=10,
    )

    estilo_portada_datos = ParagraphStyle(
        name="PortadaDatosEstadoArte",
        fontName="Helvetica",
        fontSize=10.5,
        leading=15,
        alignment=TA_CENTER,
        spaceAfter=5,
    )

    estilo_titulo = ParagraphStyle(
        name="TituloEstadoArte",
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#2E7D32"),
        spaceBefore=10,
        spaceAfter=8,
        keepWithNext=True,
    )

    estilo_subtitulo = ParagraphStyle(
        name="SubtituloEstadoArte",
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1B5E20"),
        spaceBefore=7,
        spaceAfter=4,
        keepWithNext=True,
    )

    estilo_normal = ParagraphStyle(
        name="NormalEstadoArte",
        fontName="Helvetica",
        fontSize=9.8,
        leading=14.2,
        alignment=TA_JUSTIFY,
        spaceAfter=7,
    )

    estilo_lista = ParagraphStyle(
        name="ListaEstadoArte",
        parent=estilo_normal,
        leftIndent=0.55 * cm,
        firstLineIndent=-0.35 * cm,
        spaceAfter=5,
    )

    estilo_referencia = ParagraphStyle(
        name="ReferenciaEstadoArte",
        fontName="Helvetica",
        fontSize=8.8,
        leading=12.2,
        alignment=TA_LEFT,
        leftIndent=0.6 * cm,
        firstLineIndent=-0.6 * cm,
        spaceAfter=5,
    )

    estilo_tabla = ParagraphStyle(
        name="TablaEstadoArte",
        fontName="Helvetica",
        fontSize=8.2,
        leading=10.5,
        alignment=TA_LEFT,
    )

    estilo_tabla_negrita = ParagraphStyle(
        name="TablaNegritaEstadoArte",
        fontName="Helvetica-Bold",
        fontSize=8.2,
        leading=10.5,
        alignment=TA_LEFT,
    )

    historia = []

    # Portada
    historia.append(Spacer(1, 2.2 * cm))
    historia.append(Paragraph("ESTADO DEL ARTE", estilo_portada_titulo))
    historia.append(
        Paragraph(
            _texto_seguro(datos.get("nombre_proyecto", "")),
            estilo_portada_proyecto,
        )
    )
    historia.append(
        Paragraph(
            f"<b>Código del proyecto:</b> {_texto_seguro(datos.get('codigo_proyecto', ''))}",
            estilo_portada_datos,
        )
    )

    fecha_documento = datos.get("fecha_documento")
    fecha_texto = (
        fecha_documento.strftime("%d/%m/%Y")
        if isinstance(fecha_documento, date)
        else str(fecha_documento or "")
    )

    historia.append(
        Paragraph(
            f"<b>Fecha:</b> {_texto_seguro(fecha_texto)}",
            estilo_portada_datos,
        )
    )
    historia.append(
        Paragraph(
            "<b>Tecnologías previstas:</b> "
            + _texto_seguro(", ".join(datos.get("tecnologias_previstas", []))),
            estilo_portada_datos,
        )
    )
    historia.append(Spacer(1, 1.2 * cm))
    historia.append(
        Paragraph(
            "RED TECNOPARQUE SENA",
            estilo_portada_proyecto,
        )
    )
    historia.append(PageBreak())

    # Datos generales
    datos_generales = [
        [
            Paragraph("<b>Código del proyecto</b>", estilo_tabla_negrita),
            _parrafo(datos.get("codigo_proyecto", ""), estilo_tabla),
        ],
        [
            Paragraph("<b>Nombre del proyecto</b>", estilo_tabla_negrita),
            _parrafo(datos.get("nombre_proyecto", ""), estilo_tabla),
        ],
        [
            Paragraph("<b>Fecha del documento</b>", estilo_tabla_negrita),
            _parrafo(fecha_texto, estilo_tabla),
        ],
        [
            Paragraph("<b>Tecnologías previstas</b>", estilo_tabla_negrita),
            _parrafo(", ".join(datos.get("tecnologias_previstas", [])), estilo_tabla),
        ],
        [
            Paragraph("<b>Modo de generación</b>", estilo_tabla_negrita),
            _parrafo(datos.get("modo_generacion", ""), estilo_tabla),
        ],
    ]

    tabla_datos = Table(
        datos_generales,
        colWidths=[4.2 * cm, 12.1 * cm],
    )
    tabla_datos.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#666666")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E8F1E8")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    historia.append(tabla_datos)
    historia.append(Spacer(1, 0.25 * cm))

    def agregar_seccion(titulo: str, texto: object) -> None:
        historia.append(Paragraph(escape(titulo), estilo_titulo))
        historia.append(_parrafo(texto, estilo_normal))

    agregar_seccion("1. Introducción", contenido.get("introduccion", ""))

    historia.append(Paragraph("2. Objetivos", estilo_titulo))
    for indice, objetivo in enumerate(contenido.get("objetivos", []), start=1):
        historia.append(
            Paragraph(
                f"{indice}. {_texto_seguro(objetivo)}",
                estilo_lista,
            )
        )

    agregar_seccion(
        "3. Antecedentes y contexto del proyecto",
        contenido.get("antecedentes_contexto", ""),
    )

    agregar_seccion(
        "4. Contexto nacional e internacional",
        contenido.get("contexto_nacional_internacional", ""),
    )

    historia.append(Paragraph("5. Proyectos e iniciativas similares", estilo_titulo))
    for indice, item in enumerate(contenido.get("proyectos_similares", []), start=1):
        numero = item.get("numero") or indice
        historia.append(
            Paragraph(
                f"{_texto_seguro(numero)}. {_texto_seguro(item.get('nombre', ''))}",
                estilo_subtitulo,
            )
        )
        historia.append(
            _parrafo(item.get("descripcion_breve", ""), estilo_normal)
        )

        if item.get("enlace"):
            historia.append(
                _parrafo(item.get("enlace", ""), estilo_referencia, permitir_enlace=True)
            )

        if item.get("referencia_apa"):
            historia.append(
                Paragraph(
                    f"<b>Referencia APA:</b> {_texto_seguro(item.get('referencia_apa', ''))}",
                    estilo_referencia,
                )
            )

    historia.append(Paragraph("6. Tecnologías relevantes previstas", estilo_titulo))
    for item in contenido.get("tecnologias_relevantes", []):
        historia.append(
            Paragraph(
                _texto_seguro(item.get("tecnologia", "")),
                estilo_subtitulo,
            )
        )
        historia.append(_parrafo(item.get("analisis", ""), estilo_normal))
        if item.get("cita_apa"):
            historia.append(
                Paragraph(
                    f"<b>Referencia:</b> {_texto_seguro(item.get('cita_apa', ''))}",
                    estilo_referencia,
                )
            )

    historia.append(Paragraph("7. Tecnologías emergentes", estilo_titulo))
    for item in contenido.get("tecnologias_emergentes", []):
        historia.append(
            Paragraph(
                _texto_seguro(item.get("tecnologia", "")),
                estilo_subtitulo,
            )
        )
        historia.append(_parrafo(item.get("analisis", ""), estilo_normal))
        if item.get("cita_apa"):
            historia.append(
                Paragraph(
                    f"<b>Referencia:</b> {_texto_seguro(item.get('cita_apa', ''))}",
                    estilo_referencia,
                )
            )

    historia.append(Paragraph("8. Artículos y fuentes de validación", estilo_titulo))

    tabla_articulos_data = [
        [
            Paragraph("<b>N.º</b>", estilo_tabla_negrita),
            Paragraph("<b>Tecnología</b>", estilo_tabla_negrita),
            Paragraph("<b>Artículo o fuente</b>", estilo_tabla_negrita),
            Paragraph("<b>Enlace / referencia APA 7</b>", estilo_tabla_negrita),
        ]
    ]

    for indice, item in enumerate(contenido.get("articulos_validacion", []), start=1):
        numero = item.get("numero") or indice
        enlace_y_referencia = []

        if item.get("enlace"):
            enlace_y_referencia.append(
                _parrafo(item.get("enlace", ""), estilo_tabla, permitir_enlace=True)
            )

        if item.get("referencia_apa"):
            enlace_y_referencia.append(
                _parrafo(item.get("referencia_apa", ""), estilo_tabla)
            )

        if not enlace_y_referencia:
            enlace_y_referencia.append(_parrafo("", estilo_tabla))

        tabla_articulos_data.append(
            [
                _parrafo(numero, estilo_tabla),
                _parrafo(item.get("tecnologia", ""), estilo_tabla),
                _parrafo(item.get("articulo", ""), estilo_tabla),
                enlace_y_referencia,
            ]
        )

    tabla_articulos = Table(
        tabla_articulos_data,
        colWidths=[0.8 * cm, 3.0 * cm, 4.9 * cm, 7.6 * cm],
        repeatRows=1,
    )
    tabla_articulos.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#666666")),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAD3")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    historia.append(tabla_articulos)

    agregar_seccion("9. Conclusiones", contenido.get("conclusiones", ""))

    historia.append(Paragraph("10. Bibliografía", estilo_titulo))
    for referencia in contenido.get("bibliografia", []):
        historia.append(
            Paragraph(
                _texto_seguro(referencia),
                estilo_referencia,
            )
        )

    doc.build(
        historia,
        onFirstPage=encabezado_pie,
        onLaterPages=encabezado_pie,
    )

    datos_json = _serializar_datos_estado_arte(datos)
    datos_json["ruta_pdf"] = ruta_pdf
    guardar_datos_json(
        datos_json,
        nombre_archivo="datos_estado_arte.json",
    )

    return ruta_pdf


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_estado_arte(
    modo_prueba: bool = True,
    modelo_openai: str = "gpt-4.1-mini",
) -> None:
    st.markdown("---")
    st.subheader("Formulario para Estado del Arte")
    st.caption(VERSION_ESTADO_ARTE)

    st.info(
        "Este módulo genera un documento académico e investigativo con proyectos "
        "similares, tecnologías relevantes, artículos de validación y referencias APA 7."
    )

    if modo_prueba:
        st.warning(
            "El modo prueba utiliza referencias de demostración. Antes de usar el "
            "documento como soporte académico deben reemplazarse por fuentes verificables."
        )

    if "datos_estado_arte_generado" not in st.session_state:
        st.session_state.datos_estado_arte_generado = None

    if "ruta_pdf_estado_arte_generado" not in st.session_state:
        st.session_state.ruta_pdf_estado_arte_generado = None

    with st.form("form_estado_arte"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2026-143440-00001",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Título oficial de la iniciativa",
                height=95,
            )

        with col_b:
            fecha_documento = st.date_input(
                "Fecha del documento",
                value=date.today(),
            )

            tecnologias_previstas_texto = st.text_area(
                "Tecnologías previstas",
                placeholder=(
                    "Ejemplo: inteligencia artificial, aplicaciones web, "
                    "códigos QR, sensores, modelado 3D"
                ),
                height=95,
            )

        descripcion_proyecto = st.text_area(
            "Háblame sobre el proyecto",
            placeholder=(
                "Describe de qué trata, cuál es el origen de la iniciativa, "
                "qué necesidad atiende, quiénes la utilizarán y qué la hace innovadora."
            ),
            height=230,
        )

        texto_boton = (
            "Generar Estado del Arte en modo prueba"
            if modo_prueba
            else "Generar Estado del Arte con búsqueda académica"
        )
        generar_estado_arte = st.form_submit_button(texto_boton)

    if generar_estado_arte:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Descripción detallada": descripcion_proyecto,
            "Tecnologías previstas": tecnologias_previstas_texto,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        tecnologias_previstas = limpiar_lista_tecnologias(
            tecnologias_previstas_texto
        )

        mensaje_spinner = (
            "Generando Estado del Arte en modo prueba."
            if modo_prueba
            else "Investigando fuentes y generando Estado del Arte con referencias APA 7."
        )

        with st.spinner(mensaje_spinner):
            try:
                if modo_prueba:
                    contenido_estado_arte = generar_estado_arte_modo_prueba(
                        limpiar_texto(nombre_proyecto),
                        limpiar_texto(codigo_proyecto),
                        limpiar_texto(descripcion_proyecto),
                        tecnologias_previstas,
                    )
                else:
                    contenido_estado_arte = generar_estado_arte_con_chatgpt(
                        limpiar_texto(nombre_proyecto),
                        limpiar_texto(codigo_proyecto),
                        limpiar_texto(descripcion_proyecto),
                        tecnologias_previstas,
                        modelo_openai,
                    )
            except Exception as error:
                st.error(f"No se pudo generar el Estado del Arte: {error}")
                st.stop()

        datos_estado_arte = {
            "tipo_documento": "Estado del Arte",
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "fecha_documento": fecha_documento,
            "descripcion_proyecto": limpiar_texto(descripcion_proyecto),
            "tecnologias_previstas": tecnologias_previstas,
            "contenido_estado_arte": contenido_estado_arte,
            "modo_generacion": (
                "Prueba local"
                if modo_prueba
                else "ChatGPT API con búsqueda web"
            ),
            "version": VERSION_ESTADO_ARTE,
        }

        st.session_state.datos_estado_arte_generado = datos_estado_arte
        st.session_state.ruta_pdf_estado_arte_generado = None

        st.success(
            "Estado del Arte generado correctamente. Ahora puedes revisarlo y generar el PDF."
        )

    if st.session_state.get("datos_estado_arte_generado"):
        datos_estado_arte = st.session_state.datos_estado_arte_generado
        contenido = datos_estado_arte.get("contenido_estado_arte", {})

        st.markdown("## Resumen para validación")
        st.write("**Modo de generación:**", datos_estado_arte["modo_generacion"])
        st.write("**Código del proyecto:**", datos_estado_arte["codigo_proyecto"])
        st.write("**Nombre del proyecto:**", datos_estado_arte["nombre_proyecto"])
        st.write(
            "**Fecha:**",
            datos_estado_arte["fecha_documento"].strftime("%d/%m/%Y"),
        )
        st.write(
            "**Tecnologías previstas:**",
            ", ".join(datos_estado_arte["tecnologias_previstas"]),
        )

        st.markdown("### Introducción")
        st.write(contenido.get("introduccion", ""))

        st.markdown("### Objetivos")
        for indice, objetivo in enumerate(
            contenido.get("objetivos", []),
            start=1,
        ):
            st.write(f"{indice}. {objetivo}")

        with st.expander("Ver antecedentes y contexto"):
            st.write(contenido.get("antecedentes_contexto", ""))
            st.write(contenido.get("contexto_nacional_internacional", ""))

        st.markdown("### Proyectos similares")
        for item in contenido.get("proyectos_similares", []):
            st.write(
                f"**{item.get('numero', '')}. {item.get('nombre', '')}**"
            )
            st.write(item.get("descripcion_breve", ""))
            if item.get("enlace"):
                st.write(item.get("enlace", ""))

        st.markdown("### Tecnologías emergentes")
        for item in contenido.get("tecnologias_emergentes", []):
            st.write(
                f"**{item.get('tecnologia', '')}:** "
                f"{item.get('analisis', '')}"
            )

        col_json, col_pdf = st.columns(2)

        with col_json:
            datos_json_descarga = _serializar_datos_estado_arte(
                datos_estado_arte
            )

            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(
                    datos_json_descarga,
                    ensure_ascii=False,
                    indent=4,
                ),
                file_name="datos_estado_arte.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button(
                "📄 Generar PDF del Estado del Arte",
                key="generar_pdf_estado_arte",
            ):
                try:
                    ruta_pdf = generar_pdf_estado_arte(datos_estado_arte)
                    st.session_state.ruta_pdf_estado_arte_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as error:
                    st.error(f"No se pudo generar el PDF: {error}")

        ruta_pdf = st.session_state.get("ruta_pdf_estado_arte_generado")

        if ruta_pdf and Path(ruta_pdf).exists():
            with open(ruta_pdf, "rb") as archivo_pdf:
                st.download_button(
                    label="⬇️ Descargar PDF del Estado del Arte",
                    data=archivo_pdf,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )
