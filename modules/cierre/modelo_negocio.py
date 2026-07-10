from pathlib import Path
from datetime import datetime
import re

import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.units import cm

from config.constants import OUTPUT_DIR, NIVELES_TRL
from services.openai_service import generar_json_openai
from services.json_service import guardar_datos_json
from utils.nombres_archivo import safe_filename
from utils.validaciones import validar_campos_obligatorios
from utils.formatos import conteo_palabras
from utils.textos import limpiar_texto, limitar_palabras


VERSION_MODELO_NEGOCIOS = (
    "VERSION_VALIDADA_MODELO_NEGOCIOS_TRL_6_7_8_TEXTOS_150_220_PALABRAS_"
    "SIN_NOMBRE_PROYECTO_EN_CAMPOS"
)


ITEMS_MODELO_NEGOCIO = [
    "Propuesta de valor",
    "Segmento de clientes y estrategia de adopción",
    "Canales de distribución",
    "Relaciones con clientes",
    "Flujo de ingresos",
    "Recursos claves",
    "Actividades claves",
    "Alianzas claves",
    "Estructura de costos",
]


PREGUNTAS_MODELO_NEGOCIO = {
    "Propuesta de valor": (
        "¿Qué valor proporciona el producto, prototipo o servicio a los clientes o usuarios?, "
        "¿qué necesidad atiende y qué beneficios ofrece en términos de calidad, novedad, "
        "diseño, funcionalidad, eficiencia o diferenciación?"
    ),
    "Segmento de clientes y estrategia de adopción": (
        "¿A qué clientes, usuarios o beneficiarios se dirige la solución y cuál sería la "
        "estrategia para facilitar su adopción en el contexto definido?"
    ),
    "Canales de distribución": (
        "¿A través de qué canales se podría presentar, entregar, implementar o transferir "
        "la solución a los usuarios o clientes objetivo?"
    ),
    "Relaciones con clientes": (
        "¿Qué tipo de relación se debe establecer con los usuarios o clientes para orientar, "
        "acompañar, soportar y fortalecer el uso de la solución?"
    ),
    "Flujo de ingresos": (
        "¿Qué posibles fuentes de ingresos, sostenibilidad económica o mecanismos de valor "
        "podrían asociarse a la solución?"
    ),
    "Recursos claves": (
        "¿Qué recursos técnicos, humanos, físicos, digitales, financieros o de conocimiento "
        "son necesarios para implementar y sostener la solución?"
    ),
    "Actividades claves": (
        "¿Qué actividades principales deben realizarse para desarrollar, validar, implementar, "
        "operar o mejorar la solución?"
    ),
    "Alianzas claves": (
        "¿Qué aliados estratégicos podrían fortalecer la implementación, adopción, validación, "
        "financiación o escalamiento de la solución?"
    ),
    "Estructura de costos": (
        "¿Cuáles son los costos principales asociados al desarrollo, implementación, operación, "
        "mantenimiento, soporte y posible escalamiento de la solución?"
    ),
}


def remover_nombre_proyecto(texto: str, nombre_proyecto: str, codigo_proyecto: str) -> str:
    texto_limpio = str(texto or "")

    nombres_a_remover = [
        nombre_proyecto,
        codigo_proyecto,
    ]

    for nombre in nombres_a_remover:
        nombre = str(nombre or "").strip()
        if nombre:
            texto_limpio = re.sub(
                re.escape(nombre),
                "la solución",
                texto_limpio,
                flags=re.IGNORECASE,
            )

    texto_limpio = re.sub(r"\s+", " ", texto_limpio).strip()
    return texto_limpio


def normalizar_respuesta_canvas(
    texto: str,
    nombre_proyecto: str,
    codigo_proyecto: str,
    min_palabras: int = 150,
    max_palabras: int = 220,
) -> str:
    texto = limpiar_texto(texto)
    texto = remover_nombre_proyecto(texto, nombre_proyecto, codigo_proyecto)
    texto = texto.replace("…", "").replace("...", "")
    texto = re.sub(r"\s+", " ", texto).strip()

    complemento = (
        " Además, debe contemplar criterios de apropiación, acompañamiento, validación, "
        "soporte y sostenibilidad para facilitar su implementación progresiva en el contexto "
        "definido, manteniendo coherencia con las capacidades disponibles y las necesidades "
        "de los usuarios potenciales."
    )

    while conteo_palabras(texto) < min_palabras:
        texto += complemento

    texto = limitar_palabras(texto, max_palabras=max_palabras)
    texto = remover_nombre_proyecto(texto, nombre_proyecto, codigo_proyecto)
    texto = texto.replace("…", "").replace("...", "")
    texto = re.sub(r"\s+", " ", texto).strip()

    if texto and texto[-1] not in ".!?":
        texto += "."

    return texto


def generar_modelo_negocio_modo_prueba(
    nivel_trl: str,
    nombre_proyecto: str,
    codigo_proyecto: str,
    descripcion_producto: str,
    region_contexto: str,
    aspectos: str,
    clientes_objetivo: str,
) -> dict:
    contenido = {}

    base_contextual = (
        "La solución se orienta a atender una necesidad concreta del contexto descrito, "
        "integrando capacidades técnicas, criterios de apropiación y posibilidades de uso "
        "por parte de los usuarios objetivo. Su enfoque permite transformar una necesidad "
        "identificada en una alternativa funcional que puede ser validada, ajustada y "
        "fortalecida de acuerdo con las condiciones reales de implementación. "
    )

    textos = {
        "Propuesta de valor": (
            f"{base_contextual} Su valor se centra en ofrecer una respuesta práctica, "
            "diferenciada y pertinente, capaz de mejorar procesos, reducir barreras de uso "
            "y facilitar decisiones técnicas o productivas. La propuesta combina funcionalidad, "
            "novedad, facilidad de adopción y potencial de mejora continua, considerando "
            "la descripción del producto, los aspectos estratégicos suministrados y el nivel "
            f"de madurez {nivel_trl}."
        ),
        "Segmento de clientes y estrategia de adopción": (
            f"{base_contextual} Los usuarios o beneficiarios objetivo se relacionan con el "
            "sector, territorio o comunidad descrita, priorizando quienes presentan la necesidad "
            "que la solución busca atender. La estrategia de adopción debe iniciar con procesos "
            "de socialización, demostraciones funcionales, validaciones controladas y ajustes "
            "según retroalimentación. También resulta clave facilitar el entendimiento del valor "
            "de uso, la operación básica y los beneficios esperados."
        ),
        "Canales de distribución": (
            f"{base_contextual} Los canales de distribución pueden incluir demostraciones "
            "presenciales, jornadas técnicas, espacios institucionales, plataformas digitales, "
            "acompañamiento directo, pilotos controlados y actividades de transferencia. "
            "La selección de canales debe facilitar que los usuarios conozcan la solución, "
            "comprendan su funcionamiento, accedan a sus beneficios y puedan participar en "
            "procesos de validación, adopción o implementación."
        ),
        "Relaciones con clientes": (
            f"{base_contextual} La relación con los usuarios debe basarse en acompañamiento, "
            "orientación técnica, comunicación clara, soporte oportuno y seguimiento posterior "
            "a la implementación. Esta relación permite resolver dudas, recibir observaciones, "
            "identificar oportunidades de mejora y fortalecer la confianza en el uso de la "
            "solución. También favorece la apropiación tecnológica y la continuidad del proceso."
        ),
        "Flujo de ingresos": (
            f"{base_contextual} Las fuentes de ingreso o sostenibilidad pueden plantearse "
            "a partir de servicios de implementación, licenciamiento, acompañamiento técnico, "
            "mantenimiento, capacitación, transferencia tecnológica, personalización o esquemas "
            "de uso según el tipo de usuario. En etapas tempranas, el flujo debe entenderse como "
            "una hipótesis de sostenibilidad que requiere validación técnica, comercial y operativa."
        ),
        "Recursos claves": (
            f"{base_contextual} Los recursos claves incluyen talento humano con capacidades "
            "técnicas, conocimiento especializado, infraestructura de validación, herramientas "
            "digitales o físicas, documentación, prototipos, datos de prueba y recursos para "
            "soporte. También pueden requerirse espacios de demostración, materiales, equipos "
            "y mecanismos de seguimiento que permitan mantener la funcionalidad y continuidad "
            "de la solución."
        ),
        "Actividades claves": (
            f"{base_contextual} Las actividades claves comprenden el diseño, ajuste, validación, "
            "documentación, pruebas funcionales, socialización, acompañamiento a usuarios, "
            "levantamiento de retroalimentación y mejora continua. También se deben considerar "
            "acciones de seguimiento, control de calidad, revisión de costos, gestión de aliados "
            "y preparación de evidencia técnica que respalde la evolución del producto."
        ),
        "Alianzas claves": (
            f"{base_contextual} Las alianzas claves pueden involucrar instituciones, empresas, "
            "centros de formación, expertos técnicos, usuarios piloto, entidades territoriales, "
            "actores sectoriales y organizaciones interesadas en validar o implementar la solución. "
            "Estas alianzas permiten acceder a conocimiento, escenarios de prueba, recursos, "
            "retroalimentación y oportunidades de apropiación o escalamiento."
        ),
        "Estructura de costos": (
            f"{base_contextual} La estructura de costos debe considerar diseño, desarrollo, "
            "materiales, infraestructura, pruebas, documentación, soporte, mantenimiento, "
            "capacitación, implementación y eventuales ajustes técnicos. También pueden existir "
            "costos asociados a personal, servicios digitales, desplazamientos, licencias, "
            "reposición de componentes y acciones de validación o escalamiento."
        ),
    }

    for item in ITEMS_MODELO_NEGOCIO:
        contenido[item] = normalizar_respuesta_canvas(
            textos[item],
            nombre_proyecto=nombre_proyecto,
            codigo_proyecto=codigo_proyecto,
        )

    return contenido


def generar_modelo_negocio_con_ia(
    nivel_trl: str,
    nombre_proyecto: str,
    codigo_proyecto: str,
    descripcion_producto: str,
    region_contexto: str,
    aspectos: str,
    clientes_objetivo: str,
    modelo_openai: str,
) -> dict:
    instrucciones = """
Eres un consultor senior en modelos de negocio, innovación, transferencia tecnológica y proyectos de base tecnológica de SENA Tecnoparque.

Debes generar un Informe de Identificación del Modelo de Negocios aplicado al producto, prototipo o servicio descrito por el usuario.

Reglas obligatorias:
- Responde únicamente JSON válido.
- No uses markdown.
- No copies literalmente los textos escritos por el usuario.
- Interpreta los insumos, corrige ortografía, gramática y redacción, y amplía el concepto.
- No repitas el nombre del proyecto dentro de cada campo.
- No repitas el código del proyecto dentro de cada campo.
- El nombre y el código del proyecto solo sirven como contexto general.
- Cada respuesta debe resolver directamente la pregunta del bloque.
- No expliques cómo se debería diligenciar.
- No uses frases como “este bloque debe”, “se recomienda diligenciar”, “el usuario debe”, “en este apartado”.
- Cada campo debe tener entre 150 y 220 palabras.
- El contenido debe estar relacionado con el producto, sus características, la región o contexto de implementación y los aspectos suministrados.
- No uses puntos suspensivos.
- No inventes clientes reales, ventas, ingresos certificados, alianzas confirmadas, normas, patentes ni validaciones no aportadas.
- Redacta en tono técnico, claro, institucional y coherente.
"""

    entrada = f"""
Nivel de madurez seleccionado:
{nivel_trl}

Nombre del proyecto, solo como contexto general. No lo repitas dentro de cada respuesta:
{nombre_proyecto}

Código del proyecto, solo como contexto general. No lo repitas dentro de cada respuesta:
{codigo_proyecto}

Descripción del prototipo, producto o servicio:
{descripcion_producto}

Región o contexto de implementación:
{region_contexto}

Aspectos a tener en cuenta:
{aspectos}

Clientes, usuarios o beneficiarios objetivo:
{clientes_objetivo}

Preguntas que debe responder cada campo:
{PREGUNTAS_MODELO_NEGOCIO}

Genera exactamente estas claves JSON:
{ITEMS_MODELO_NEGOCIO}

Formato:
{{
  "Propuesta de valor": "...",
  "Segmento de clientes y estrategia de adopción": "...",
  "Canales de distribución": "...",
  "Relaciones con clientes": "...",
  "Flujo de ingresos": "...",
  "Recursos claves": "...",
  "Actividades claves": "...",
  "Alianzas claves": "...",
  "Estructura de costos": "..."
}}
"""

    respaldo = generar_modelo_negocio_modo_prueba(
        nivel_trl=nivel_trl,
        nombre_proyecto=nombre_proyecto,
        codigo_proyecto=codigo_proyecto,
        descripcion_producto=descripcion_producto,
        region_contexto=region_contexto,
        aspectos=aspectos,
        clientes_objetivo=clientes_objetivo,
    )

    try:
        datos = generar_json_openai(
            instrucciones=instrucciones,
            entrada=entrada,
            modelo=modelo_openai,
            temperature=0.25,
        )

        contenido = {}

        for item in ITEMS_MODELO_NEGOCIO:
            texto = datos.get(item, "")

            if not isinstance(texto, str) or not texto.strip():
                texto = respaldo[item]

            contenido[item] = normalizar_respuesta_canvas(
                texto,
                nombre_proyecto=nombre_proyecto,
                codigo_proyecto=codigo_proyecto,
            )

        return contenido

    except Exception as error:
        st.warning(
            "No fue posible generar el contenido con OpenAI. "
            "Se usará el modo prueba local para conservar el flujo del documento."
        )
        st.caption(f"Detalle técnico: {error}")

        return respaldo


def estilo_parrafo(nombre: str, font_size: float, leading: float, alignment=TA_JUSTIFY):
    return ParagraphStyle(
        name=nombre,
        fontName="Helvetica",
        fontSize=font_size,
        leading=leading,
        alignment=alignment,
        spaceAfter=4,
    )


def generar_pdf_modelo_negocio(datos: dict) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = f"Modelo_Negocios_{safe_filename(datos['codigo_proyecto'])}.pdf"
    ruta_pdf = OUTPUT_DIR / nombre_archivo

    doc = SimpleDocTemplate(
        str(ruta_pdf),
        pagesize=landscape(letter),
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.1 * cm,
        bottomMargin=1.0 * cm,
    )

    estilo_titulo = ParagraphStyle(
        name="Titulo",
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=17,
        alignment=TA_CENTER,
        spaceAfter=8,
    )

    estilo_subtitulo = ParagraphStyle(
        name="Subtitulo",
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        alignment=TA_LEFT,
        spaceBefore=6,
        spaceAfter=4,
    )

    estilo_normal = ParagraphStyle(
        name="Normal",
        fontName="Helvetica",
        fontSize=8.2,
        leading=10,
        alignment=TA_JUSTIFY,
        spaceAfter=5,
    )

    estilo_pregunta = ParagraphStyle(
        name="Pregunta",
        fontName="Helvetica-Oblique",
        fontSize=7.4,
        leading=9,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#333333"),
        spaceAfter=4,
    )

    estilo_tabla = ParagraphStyle(
        name="TablaCanvas",
        fontName="Helvetica",
        fontSize=4.7,
        leading=5.2,
        alignment=TA_LEFT,
    )

    historia = []

    historia.append(Paragraph("INFORME DE IDENTIFICACIÓN", estilo_titulo))
    historia.append(Paragraph("DEL MODELO DE NEGOCIOS", estilo_titulo))
    historia.append(Paragraph(datos["nivel_trl"], estilo_titulo))

    info = [
        ["Código del proyecto", datos["codigo_proyecto"]],
        ["Nombre del proyecto", datos["nombre_proyecto"]],
        ["Nivel de madurez", datos["nivel_trl"]],
        ["Región / contexto de implementación", datos["region_contexto"]],
        ["Fecha de generación", datetime.now().strftime("%d/%m/%Y")],
    ]

    tabla_info = Table(info, colWidths=[6.0 * cm, 18.5 * cm])
    tabla_info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF5EA")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    historia.append(tabla_info)
    historia.append(Spacer(1, 0.25 * cm))

    contenido = datos["contenido"]

    for idx, item in enumerate(ITEMS_MODELO_NEGOCIO, start=1):
        historia.append(Paragraph(f"{idx}. {item}", estilo_subtitulo))
        historia.append(Paragraph(PREGUNTAS_MODELO_NEGOCIO[item], estilo_pregunta))
        historia.append(Paragraph(contenido[item], estilo_normal))

    historia.append(PageBreak())
    historia.append(Paragraph("MODELO DE NEGOCIOS.", estilo_titulo))
    historia.append(Spacer(1, 0.2 * cm))

    tabla_canvas = [
        [
            Paragraph("<b>Alianzas claves</b><br/>" + contenido["Alianzas claves"], estilo_tabla),
            Paragraph("<b>Actividades claves</b><br/>" + contenido["Actividades claves"], estilo_tabla),
            Paragraph("<b>Propuesta de valor</b><br/>" + contenido["Propuesta de valor"], estilo_tabla),
            Paragraph("<b>Relaciones con clientes</b><br/>" + contenido["Relaciones con clientes"], estilo_tabla),
            Paragraph("<b>Segmento de clientes</b><br/>" + contenido["Segmento de clientes y estrategia de adopción"], estilo_tabla),
        ],
        [
            Paragraph("", estilo_tabla),
            Paragraph("<b>Recursos claves</b><br/>" + contenido["Recursos claves"], estilo_tabla),
            Paragraph("", estilo_tabla),
            Paragraph("<b>Canales</b><br/>" + contenido["Canales de distribución"], estilo_tabla),
            Paragraph("", estilo_tabla),
        ],
        [
            Paragraph("<b>Estructura de costos</b><br/>" + contenido["Estructura de costos"], estilo_tabla),
            Paragraph("", estilo_tabla),
            Paragraph("", estilo_tabla),
            Paragraph("<b>Flujo de ingresos</b><br/>" + contenido["Flujo de ingresos"], estilo_tabla),
            Paragraph("", estilo_tabla),
        ],
    ]

    tabla = Table(
        tabla_canvas,
        colWidths=[5.1 * cm, 5.1 * cm, 5.1 * cm, 5.1 * cm, 5.1 * cm],
        rowHeights=[5.6 * cm, 5.0 * cm, 4.7 * cm],
    )

    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BACKGROUND", (0, 0), (0, 1), colors.HexColor("#D9EAD3")),
                ("BACKGROUND", (1, 0), (1, 1), colors.HexColor("#FFF2CC")),
                ("BACKGROUND", (2, 0), (2, 1), colors.HexColor("#D9EAD3")),
                ("BACKGROUND", (3, 0), (3, 1), colors.HexColor("#D9EAD3")),
                ("BACKGROUND", (4, 0), (4, 1), colors.HexColor("#D9EAD3")),
                ("BACKGROUND", (0, 2), (1, 2), colors.HexColor("#FCE4D6")),
                ("BACKGROUND", (3, 2), (4, 2), colors.HexColor("#FCE4D6")),
                ("SPAN", (0, 2), (2, 2)),
                ("SPAN", (3, 2), (4, 2)),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    historia.append(tabla)

    doc.build(historia)

    guardar_datos_json(
        datos={**datos, "ruta_pdf": str(ruta_pdf)},
        nombre_archivo="datos_modelo_negocio.json",
    )

    return str(ruta_pdf)


def render_modelo_negocio(modo_prueba: bool, modelo_openai: str) -> None:
    st.markdown("---")
    st.subheader("Modelo de negocio Lean Canvas")

    st.caption(VERSION_MODELO_NEGOCIOS)

    st.info(
        "Este módulo genera el Informe de Identificación del Modelo de Negocios "
        "con textos entre 150 y 220 palabras por campo, según el nivel TRL seleccionado."
    )

    with st.form("form_modelo_negocio"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2026-143440-00001",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Nombre oficial del proyecto",
                height=80,
            )

            nivel_trl = st.selectbox(
                "Nivel de madurez tecnológica",
                options=NIVELES_TRL,
                index=0,
            )

            region_contexto = st.text_area(
                "Región o contexto de implementación del proyecto",
                placeholder=(
                    "Ejemplo: Regional Huila, sector cafetero, centros de formación SENA, "
                    "emprendimientos locales, instituciones educativas, etc."
                ),
                height=100,
            )

        with col_b:
            descripcion_producto = st.text_area(
                "Descripción del prototipo, producto o servicio",
                placeholder=(
                    "Describe qué es, qué hace, cómo funciona, qué necesidad atiende "
                    "y a quién podría servir."
                ),
                height=130,
            )

            aspectos = st.text_area(
                "Aspectos a tener en cuenta para la generación del modelo de negocio",
                placeholder=(
                    "Ejemplo: diferenciadores, restricciones, canales, aliados, forma de adopción, "
                    "modelo de ingresos, costos, soporte o validaciones."
                ),
                height=130,
            )

            clientes_objetivo = st.text_area(
                "Clientes, usuarios o beneficiarios objetivo",
                placeholder=(
                    "Campo opcional. Ejemplo: productores, centros de formación, emprendedores, "
                    "empresas, instituciones, comunidades o entidades públicas."
                ),
                height=90,
            )

        generar = st.form_submit_button("Generar Informe de Modelo de Negocios")

    if generar:
        campos = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Región o contexto": region_contexto,
            "Descripción del producto": descripcion_producto,
            "Aspectos a tener en cuenta": aspectos,
        }

        if not validar_campos_obligatorios(campos):
            st.stop()

        progreso = st.progress(0)
        estado = st.empty()

        estado.info("Generando contenido del modelo de negocios...")
        progreso.progress(25)

        if modo_prueba:
            contenido = generar_modelo_negocio_modo_prueba(
                nivel_trl=nivel_trl,
                nombre_proyecto=nombre_proyecto,
                codigo_proyecto=codigo_proyecto,
                descripcion_producto=descripcion_producto,
                region_contexto=region_contexto,
                aspectos=aspectos,
                clientes_objetivo=clientes_objetivo,
            )
        else:
            contenido = generar_modelo_negocio_con_ia(
                nivel_trl=nivel_trl,
                nombre_proyecto=nombre_proyecto,
                codigo_proyecto=codigo_proyecto,
                descripcion_producto=descripcion_producto,
                region_contexto=region_contexto,
                aspectos=aspectos,
                clientes_objetivo=clientes_objetivo,
                modelo_openai=modelo_openai,
            )

        progreso.progress(70)
        estado.info("Construyendo PDF del modelo de negocios...")

        datos = {
            "tipo_documento": "Modelo de Negocios",
            "codigo_proyecto": codigo_proyecto,
            "nombre_proyecto": nombre_proyecto,
            "nivel_trl": nivel_trl,
            "region_contexto": region_contexto,
            "descripcion_producto": descripcion_producto,
            "aspectos": aspectos,
            "clientes_objetivo": clientes_objetivo,
            "contenido": contenido,
            "modo_generacion": "Prueba local" if modo_prueba else "ChatGPT API",
            "version": VERSION_MODELO_NEGOCIOS,
        }

        ruta_pdf = generar_pdf_modelo_negocio(datos)

        st.session_state.datos_modelo_negocio = datos
        st.session_state.ruta_pdf_modelo_negocio = ruta_pdf

        progreso.progress(100)
        estado.success("Informe de Modelo de Negocios generado correctamente.")

    if st.session_state.get("datos_modelo_negocio"):
        datos = st.session_state.datos_modelo_negocio

        st.markdown("## Resumen para validación")
        st.write("**Código del proyecto:**", datos["codigo_proyecto"])
        st.write("**Nombre del proyecto:**", datos["nombre_proyecto"])
        st.write("**Nivel de madurez:**", datos["nivel_trl"])
        st.write("**Modo de generación:**", datos["modo_generacion"])

        for item in ITEMS_MODELO_NEGOCIO:
            with st.expander(f"{item} - {conteo_palabras(datos['contenido'][item])} palabras"):
                st.write(datos["contenido"][item])

        ruta_pdf = st.session_state.get("ruta_pdf_modelo_negocio")

        if ruta_pdf and Path(ruta_pdf).exists():
            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar PDF Modelo de Negocios",
                    data=f,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )