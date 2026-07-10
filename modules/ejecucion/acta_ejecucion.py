from __future__ import annotations

from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
import json

import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Table,
    TableStyle,
)

from config.constants import OUTPUT_DIR, RUTA_LOGO_SENA
from services.json_service import guardar_datos_json
from services.openai_service import generar_json_openai
from utils.nombres_archivo import safe_filename
from utils.textos import limpiar_texto
from utils.validaciones import validar_campos_obligatorios


VERSION_ACTA_EJECUCION = (
    "VERSION_MODULAR_ACTA_EJECUCION_FORMATO_INSTITUCIONAL_VALIDADO"
)
VALOR_HORA_EXPERTO = 25266

DIAS_SEMANA = [
    "Lunes",
    "Martes",
    "Miércoles",
    "Jueves",
    "Viernes",
    "Sábado",
    "Domingo",
]

MAPA_DIAS = {
    "lunes": 0,
    "martes": 1,
    "miércoles": 2,
    "miercoles": 2,
    "jueves": 3,
    "viernes": 4,
    "sábado": 5,
    "sabado": 5,
    "domingo": 6,
}


# =====================================================
# UTILIDADES
# =====================================================

def formato_moneda_colombiana(valor: float | int) -> str:
    try:
        valor_int = int(round(float(valor)))
    except (TypeError, ValueError):
        valor_int = 0

    return "$" + f"{valor_int:,}".replace(",", ".")


def obtener_ruta_logo_sena() -> str | None:
    candidatos = [
        Path(RUTA_LOGO_SENA),
        Path("resources/logo_sena.png"),
        Path("recursos/logo_sena.png"),
    ]

    for ruta in candidatos:
        if ruta.exists():
            return str(ruta)

    return None


def normalizar_dia_semana(nombre_dia: str) -> int:
    clave = str(nombre_dia or "").strip().casefold()

    if clave not in MAPA_DIAS:
        raise ValueError(f"Día de la semana no válido: {nombre_dia}")

    return MAPA_DIAS[clave]


def obtener_fechas_programadas(
    fecha_inicio: date,
    fecha_fin: date,
    dias_semana: list[str],
) -> list[date]:
    if fecha_fin < fecha_inicio:
        return []

    dias_validos = {
        normalizar_dia_semana(nombre_dia)
        for nombre_dia in dias_semana
    }

    fechas: list[date] = []
    fecha_actual = fecha_inicio

    while fecha_actual <= fecha_fin:
        if fecha_actual.weekday() in dias_validos:
            fechas.append(fecha_actual)

        fecha_actual += timedelta(days=1)

    return fechas


def distribuir_fechas_para_asesorias(
    fecha_inicio: date,
    fecha_fin: date,
    dias_semana: list[str],
    cantidad_asesorias: int,
) -> list[date]:
    fechas_disponibles = obtener_fechas_programadas(
        fecha_inicio,
        fecha_fin,
        dias_semana,
    )

    if not fechas_disponibles or cantidad_asesorias <= 0:
        return []

    fechas_resultado: list[date] = []

    for indice in range(cantidad_asesorias):
        posicion = round(
            indice
            * (len(fechas_disponibles) - 1)
            / max(cantidad_asesorias - 1, 1)
        )
        fechas_resultado.append(fechas_disponibles[posicion])

    return fechas_resultado


def _texto_pdf(valor: object) -> str:
    return escape(str(valor or "").strip()).replace("\n", "<br/>")


def _parrafo(
    valor: object,
    estilo: ParagraphStyle,
    negrita: bool = False,
) -> Paragraph:
    texto = _texto_pdf(valor)

    if negrita:
        texto = f"<b>{texto}</b>"

    return Paragraph(texto, estilo)


def _serializar_datos(datos: dict) -> dict:
    resultado = dict(datos)

    for campo in ["fecha_inicio", "fecha_fin"]:
        valor = resultado.get(campo)

        if isinstance(valor, date):
            resultado[campo] = valor.strftime("%d/%m/%Y")

    return resultado


# =====================================================
# ASESORÍAS
# =====================================================

def generar_asesorias_ejecucion_modo_prueba(
    descripcion_proyecto: str,
    cantidad_asesorias: int,
    fechas_asesorias: list[date],
    horas_por_asesoria: float,
) -> list[dict]:
    del descripcion_proyecto

    actividades_base = [
        (
            "Revisión del avance técnico del proyecto y validación de "
            "requerimientos definidos en la fase de planeación."
        ),
        (
            "Asesoría para la estructuración de componentes técnicos y "
            "definición de criterios de diseño de la solución."
        ),
        (
            "Acompañamiento en la selección de tecnologías, materiales, "
            "herramientas o recursos requeridos para el desarrollo."
        ),
        (
            "Revisión del diseño preliminar, ajustes funcionales y "
            "recomendaciones para la construcción o implementación."
        ),
        (
            "Asesoría en pruebas iniciales, verificación de resultados y "
            "análisis de funcionamiento de la solución propuesta."
        ),
        (
            "Acompañamiento en ajustes técnicos derivados de la validación "
            "del prototipo o componente desarrollado."
        ),
        (
            "Revisión de evidencias, documentación técnica y consolidación "
            "de avances del proyecto."
        ),
        (
            "Asesoría para cierre parcial de actividades, identificación "
            "de mejoras y definición de siguientes pasos."
        ),
    ]

    asesorias: list[dict] = []

    for indice in range(cantidad_asesorias):
        fecha = (
            fechas_asesorias[indice]
            if indice < len(fechas_asesorias)
            else date.today()
        )

        asesorias.append(
            {
                "fecha": fecha.strftime("%d/%m/%Y"),
                "horas": float(horas_por_asesoria),
                "descripcion": actividades_base[
                    indice % len(actividades_base)
                ],
            }
        )

    return asesorias


def generar_asesorias_ejecucion_con_ia(
    descripcion_proyecto: str,
    cantidad_asesorias: int,
    fechas_asesorias: list[date],
    horas_por_asesoria: float,
    modelo_openai: str = "gpt-4.1-mini",
) -> list[dict]:
    fechas_texto = [
        fecha.strftime("%d/%m/%Y")
        for fecha in fechas_asesorias
    ]

    if not fechas_texto:
        return []

    instrucciones = """
Eres un experto en seguimiento técnico de proyectos de base tecnológica
de la Red Tecnoparque SENA.

Genera actividades de asesoría y uso de infraestructura para un acta de ejecución.
Cada actividad debe ser técnica, clara, verificable y coherente con la descripción
del proyecto.

No inventes nombres de personas, códigos, valores económicos ni entidades.
Usa exclusivamente las fechas suministradas.
Conserva exactamente la cantidad de registros solicitada.
Responde únicamente en JSON válido, sin markdown ni explicaciones.
"""

    entrada = f"""
Descripción general del proyecto:
{descripcion_proyecto}

Cantidad exacta de asesorías:
{cantidad_asesorias}

Horas por asesoría:
{horas_por_asesoria}

Fechas autorizadas:
{", ".join(fechas_texto)}

Formato JSON obligatorio:
{{
  "asesorias": [
    {{
      "fecha": "dd/mm/aaaa",
      "horas": {horas_por_asesoria},
      "descripcion": "Descripción técnica de la asesoría realizada"
    }}
  ]
}}
"""

    try:
        respuesta = generar_json_openai(
            instrucciones=instrucciones,
            entrada=entrada,
            modelo=modelo_openai,
            temperature=0.35,
        )
    except Exception:
        return generar_asesorias_ejecucion_modo_prueba(
            descripcion_proyecto,
            cantidad_asesorias,
            fechas_asesorias,
            horas_por_asesoria,
        )

    if not isinstance(respuesta, dict):
        respuesta = {}

    registros = respuesta.get("asesorias", [])

    if not isinstance(registros, list):
        registros = []

    fechas_validas = set(fechas_texto)
    asesorias_limpias: list[dict] = []

    for indice, item in enumerate(registros):
        if not isinstance(item, dict):
            continue

        fecha_item = str(item.get("fecha", "")).strip()

        if fecha_item not in fechas_validas:
            fecha_item = fechas_texto[
                min(indice, len(fechas_texto) - 1)
            ]

        descripcion = limpiar_texto(
            str(item.get("descripcion", ""))
        )

        if not descripcion:
            descripcion = (
                "Asesoría técnica para seguimiento del proyecto."
            )

        asesorias_limpias.append(
            {
                "fecha": fecha_item,
                "horas": float(horas_por_asesoria),
                "descripcion": descripcion,
            }
        )

    if len(asesorias_limpias) < cantidad_asesorias:
        faltantes = cantidad_asesorias - len(asesorias_limpias)
        inicio_faltantes = len(asesorias_limpias)

        adicionales = generar_asesorias_ejecucion_modo_prueba(
            descripcion_proyecto,
            faltantes,
            fechas_asesorias[
                inicio_faltantes:inicio_faltantes + faltantes
            ],
            horas_por_asesoria,
        )
        asesorias_limpias.extend(adicionales)

    return asesorias_limpias[:cantidad_asesorias]


# =====================================================
# PDF - FORMATO VALIDADO
# =====================================================

def generar_pdf_acta_ejecucion(datos: dict) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Acta_Ejecucion_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(OUTPUT_DIR / nombre_archivo)

    page_size = landscape(letter)

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=page_size,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.0 * cm,
        title=(
            "Seguimiento de Asesorías y Usos de Infraestructura"
        ),
        author="Red Tecnoparque SENA",
    )

    estilo_titulo = ParagraphStyle(
        name="TituloActaEjecucion",
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=13,
        alignment=TA_CENTER,
    )

    estilo_header = ParagraphStyle(
        name="HeaderActaEjecucion",
        fontName="Helvetica-Bold",
        fontSize=7.4,
        leading=8.5,
        alignment=TA_CENTER,
    )

    estilo_celda = ParagraphStyle(
        name="CeldaActaEjecucion",
        fontName="Helvetica",
        fontSize=7.2,
        leading=8.5,
        alignment=TA_CENTER,
    )

    estilo_celda_left = ParagraphStyle(
        name="CeldaLeftActaEjecucion",
        fontName="Helvetica",
        fontSize=7.2,
        leading=8.5,
        alignment=TA_LEFT,
    )

    estilo_negrita_left = ParagraphStyle(
        name="NegritaLeftActaEjecucion",
        fontName="Helvetica-Bold",
        fontSize=7.4,
        leading=8.5,
        alignment=TA_LEFT,
    )

    historia = []

    # Encabezado institucional validado
    ruta_logo = obtener_ruta_logo_sena()

    if ruta_logo:
        try:
            logo = Image(
                ruta_logo,
                width=2.2 * cm,
                height=1.8 * cm,
            )
        except Exception:
            logo = Paragraph("SENA", estilo_titulo)
    else:
        logo = Paragraph("SENA", estilo_titulo)

    encabezado_data = [
        [
            logo,
            Paragraph(
                "Seguimiento de Asesorías y Usos de Infraestructura",
                estilo_titulo,
            ),
        ],
        [
            "",
            Paragraph(
                "ACTA No. 02 del proyecto No "
                f"{_texto_pdf(datos.get('codigo_proyecto', ''))}",
                estilo_titulo,
            ),
        ],
    ]

    tabla_encabezado = Table(
        encabezado_data,
        colWidths=[7.0 * cm, 19.0 * cm],
        rowHeights=[1.0 * cm, 0.8 * cm],
    )
    tabla_encabezado.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("SPAN", (0, 0), (0, 1)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    historia.append(tabla_encabezado)

    gris = colors.HexColor("#BFBFBF")

    # Información general
    seccion_info = Table(
        [[Paragraph("Información general", estilo_header)]],
        colWidths=[26.0 * cm],
        rowHeights=[0.55 * cm],
    )
    seccion_info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), gris),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    historia.append(seccion_info)

    info_data = [
        [
            Paragraph("Código del proyecto", estilo_header),
            Paragraph("Nombre del proyecto", estilo_header),
            Paragraph("Experto a cargo del proyecto", estilo_header),
            Paragraph("Sublínea tecnológica", estilo_header),
        ],
        [
            _parrafo(datos.get("codigo_proyecto", ""), estilo_celda),
            _parrafo(datos.get("nombre_proyecto", ""), estilo_celda),
            _parrafo(datos.get("nombre_experto", ""), estilo_celda),
            _parrafo(
                datos.get("sublinea_tecnologica", ""),
                estilo_celda,
            ),
        ],
    ]

    tabla_info = Table(
        info_data,
        colWidths=[
            5.8 * cm,
            10.0 * cm,
            5.2 * cm,
            5.0 * cm,
        ],
        rowHeights=[0.55 * cm, 2.05 * cm],
    )
    tabla_info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    historia.append(tabla_info)

    # Talentos
    seccion_talentos = Table(
        [[Paragraph("Talentos del Proyecto", estilo_header)]],
        colWidths=[26.0 * cm],
        rowHeights=[0.55 * cm],
    )
    seccion_talentos.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), gris),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    historia.append(seccion_talentos)

    talentos_data = [
        [
            Paragraph("Número de documento", estilo_header),
            Paragraph("Nombres y apellidos", estilo_header),
            Paragraph("Número de contacto", estilo_header),
        ],
        [
            _parrafo(datos.get("documento_talento", ""), estilo_celda),
            _parrafo(datos.get("nombre_talento", ""), estilo_celda),
            _parrafo(datos.get("telefono_talento", ""), estilo_celda),
        ],
    ]

    tabla_talentos = Table(
        talentos_data,
        colWidths=[8.5 * cm, 11.5 * cm, 6.0 * cm],
        rowHeights=[0.55 * cm, 0.75 * cm],
    )
    tabla_talentos.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    historia.append(tabla_talentos)

    # Asesorías y usos
    seccion_asesorias = Table(
        [[Paragraph("Asesorías y usos", estilo_header)]],
        colWidths=[26.0 * cm],
        rowHeights=[0.55 * cm],
    )
    seccion_asesorias.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), gris),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    historia.append(seccion_asesorias)

    asesorias_data = [
        [
            Paragraph(
                "Fecha de la Asesoría y uso de infraestructura",
                estilo_header,
            ),
            Paragraph("Horas de Asesoría", estilo_header),
            Paragraph("Descripción", estilo_header),
        ]
    ]

    for item in datos.get("asesorias", []):
        asesorias_data.append(
            [
                _parrafo(item.get("fecha", ""), estilo_celda),
                _parrafo(item.get("horas", ""), estilo_celda),
                _parrafo(
                    item.get("descripcion", ""),
                    estilo_celda_left,
                ),
            ]
        )

    asesorias_data.append(
        [
            "",
            "",
            Paragraph(
                "<b>Valor total de la asesoría "
                f"(Valor hora: {formato_moneda_colombiana(VALOR_HORA_EXPERTO)}) "
                f"{formato_moneda_colombiana(datos.get('total_honorarios', 0))}</b>",
                estilo_negrita_left,
            ),
        ]
    )

    filas_asesorias = len(asesorias_data)
    row_heights_asesorias = (
        [0.55 * cm]
        + [None for _ in datos.get("asesorias", [])]
        + [0.55 * cm]
    )

    tabla_asesorias = Table(
        asesorias_data,
        colWidths=[7.0 * cm, 4.0 * cm, 15.0 * cm],
        rowHeights=row_heights_asesorias,
        repeatRows=1,
    )
    tabla_asesorias.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                (
                    "SPAN",
                    (0, filas_asesorias - 1),
                    (1, filas_asesorias - 1),
                ),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    historia.append(tabla_asesorias)

    # Materiales, equipos e insumos
    seccion_materiales = Table(
        [[
            Paragraph(
                "Materiales, equipos e insumos utilizados en el proyecto",
                estilo_header,
            )
        ]],
        colWidths=[26.0 * cm],
        rowHeights=[0.55 * cm],
    )
    seccion_materiales.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), gris),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    historia.append(seccion_materiales)

    materiales_data = [
        [
            Paragraph(
                "Nombre del equipo/Material Usado",
                estilo_header,
            ),
            Paragraph("horas de uso/cantidad", estilo_header),
            Paragraph("valor total", estilo_header),
        ]
    ]

    for item in datos.get("equipos_materiales", []):
        materiales_data.append(
            [
                _parrafo(item.get("nombre", ""), estilo_celda_left),
                _parrafo(
                    item.get("cantidad_horas", ""),
                    estilo_celda,
                ),
                _parrafo(
                    formato_moneda_colombiana(
                        item.get("valor_total", 0)
                    ),
                    estilo_celda,
                ),
            ]
        )

    materiales_data.append(
        [
            "",
            Paragraph(
                "<b>Costo total uso de equipos e infraestructura</b>",
                estilo_negrita_left,
            ),
            Paragraph(
                "<b>"
                f"{formato_moneda_colombiana(datos.get('total_equipos_materiales', 0))}"
                "</b>",
                estilo_celda,
            ),
        ]
    )

    filas_materiales = len(materiales_data)
    row_heights_materiales = (
        [0.55 * cm]
        + [None for _ in datos.get("equipos_materiales", [])]
        + [0.55 * cm]
    )

    tabla_materiales = Table(
        materiales_data,
        colWidths=[8.0 * cm, 12.0 * cm, 6.0 * cm],
        rowHeights=row_heights_materiales,
        repeatRows=1,
    )
    tabla_materiales.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                (
                    "SPAN",
                    (0, filas_materiales - 1),
                    (0, filas_materiales - 1),
                ),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    historia.append(tabla_materiales)

    # Total general
    total_general = Table(
        [
            [
                Paragraph(
                    "Costo total honorarios experto más valor de uso "
                    "de equipos y materiales",
                    estilo_negrita_left,
                ),
                Paragraph(
                    "<b>"
                    f"{formato_moneda_colombiana(datos.get('total_general', 0))}"
                    "</b>",
                    estilo_celda,
                ),
            ]
        ],
        colWidths=[20.0 * cm, 6.0 * cm],
        rowHeights=[0.65 * cm],
    )
    total_general.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    historia.append(total_general)

    # Firmas
    seccion_firmas = Table(
        [[Paragraph("Firma Expertos y Talentos", estilo_header)]],
        colWidths=[26.0 * cm],
        rowHeights=[0.55 * cm],
    )
    seccion_firmas.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), gris),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    firmas_data = [
        [
            _parrafo(
                f"{datos.get('nombre_experto', '')} - Experto",
                estilo_celda,
            ),
            _parrafo(
                f"{datos.get('nombre_talento', '')} - Talento Interlocutor",
                estilo_celda,
            ),
        ],
        [
            Paragraph("", estilo_celda),
            Paragraph("", estilo_celda),
        ],
    ]

    tabla_firmas = Table(
        firmas_data,
        colWidths=[13.0 * cm, 13.0 * cm],
        rowHeights=[0.70 * cm, 0.95 * cm],
    )
    tabla_firmas.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    historia.append(KeepTogether([seccion_firmas, tabla_firmas]))

    doc.build(historia)

    datos_json = _serializar_datos(datos)
    datos_json["ruta_pdf"] = ruta_pdf

    guardar_datos_json(
        datos_json,
        nombre_archivo="datos_acta_ejecucion.json",
    )

    return ruta_pdf


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_acta_ejecucion(
    modo_prueba: bool = True,
    modelo_openai: str = "gpt-4.1-mini",
) -> None:
    st.markdown("---")
    st.subheader("Acta de ejecución")
    st.caption(VERSION_ACTA_EJECUCION)

    st.info(
        "Este módulo genera el documento institucional de seguimiento "
        "de asesorías y usos de infraestructura. Incluye cálculo "
        "automático de honorarios, equipos, materiales y total general."
    )

    if "datos_acta_ejecucion_generada" not in st.session_state:
        st.session_state.datos_acta_ejecucion_generada = None

    if "ruta_pdf_acta_ejecucion_generado" not in st.session_state:
        st.session_state.ruta_pdf_acta_ejecucion_generado = None

    metodo_asesorias = st.radio(
        "Método de generación de asesorías y usos",
        options=["Generación con IA", "Generación manual"],
        horizontal=True,
        key="ejecucion_metodo_asesorias",
    )

    st.markdown("### Equipos y materiales")
    cantidad_equipos = st.number_input(
        "¿Cuántos registros de uso de equipos y materiales desea ingresar?",
        min_value=0,
        max_value=30,
        value=1,
        step=1,
        key="ejecucion_cantidad_equipos",
    )

    st.markdown("### Asesorías y usos")

    if metodo_asesorias == "Generación manual":
        cantidad_asesorias_manual = st.number_input(
            "Cantidad de registros de asesorías y usos",
            min_value=1,
            max_value=30,
            value=4,
            step=1,
            key="ejecucion_cantidad_asesorias_manual",
        )
    else:
        cantidad_asesorias_manual = 0

    with st.form("form_acta_ejecucion"):
        st.markdown("## Información general")

        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2025-1431026-17218",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Nombre oficial del proyecto",
                height=90,
            )

            sublinea_tecnologica = st.text_input(
                "Sublínea tecnológica",
                placeholder="Ejemplo: IND - Productos y procesos",
            )

            nombre_experto = st.text_input(
                "Nombre del experto",
                placeholder="Nombre completo del experto",
            )

        with col_b:
            nombre_talento = st.text_input(
                "Nombre del talento",
                placeholder="Nombre completo del talento interlocutor",
            )

            documento_talento = st.text_input(
                "Documento de identidad del talento",
                placeholder="Número de documento",
            )

            telefono_talento = st.text_input(
                "Teléfono del talento",
                placeholder="Número de contacto",
            )

        st.markdown("## Equipos, materiales e insumos utilizados")

        equipos_materiales: list[dict] = []

        for indice in range(int(cantidad_equipos)):
            st.markdown(
                f"**Registro equipo/material {indice + 1}**"
            )

            col_1, col_2, col_3 = st.columns([2, 1, 1])

            with col_1:
                nombre_equipo = st.text_input(
                    f"Nombre del equipo / Material usado {indice + 1}",
                    key=f"ejecucion_equipo_nombre_{indice}",
                )

            with col_2:
                cantidad_horas = st.text_input(
                    f"Horas de uso / Cantidad {indice + 1}",
                    key=f"ejecucion_equipo_cantidad_{indice}",
                )

            with col_3:
                valor_total = st.number_input(
                    f"Valor total {indice + 1}",
                    min_value=0,
                    value=0,
                    step=1000,
                    key=f"ejecucion_equipo_valor_{indice}",
                )

            equipos_materiales.append(
                {
                    "nombre": nombre_equipo,
                    "cantidad_horas": cantidad_horas,
                    "valor_total": valor_total,
                }
            )

        st.markdown("## Asesorías y usos")

        asesorias: list[dict] = []
        descripcion_proyecto = ""
        fecha_inicio: date | None = None
        fecha_fin: date | None = None
        dias_ejecucion: list[str] = []
        cantidad_asesorias_ia = 0
        horas_por_asesoria_ia = 0.0

        if metodo_asesorias == "Generación manual":
            for indice in range(int(cantidad_asesorias_manual)):
                st.markdown(
                    f"**Asesoría / uso {indice + 1}**"
                )

                col_1, col_2, col_3 = st.columns([1.2, 1, 3])

                with col_1:
                    fecha_asesoria = st.date_input(
                        f"Fecha asesoría {indice + 1}",
                        value=date.today(),
                        key=f"ejecucion_asesoria_fecha_{indice}",
                    )

                with col_2:
                    horas_asesoria = st.number_input(
                        f"Horas asesoría {indice + 1}",
                        min_value=0.5,
                        max_value=12.0,
                        value=2.0,
                        step=0.5,
                        key=f"ejecucion_asesoria_horas_{indice}",
                    )

                with col_3:
                    descripcion_asesoria = st.text_area(
                        f"Descripción {indice + 1}",
                        height=70,
                        key=f"ejecucion_asesoria_descripcion_{indice}",
                    )

                asesorias.append(
                    {
                        "fecha": fecha_asesoria.strftime("%d/%m/%Y"),
                        "horas": float(horas_asesoria),
                        "descripcion": descripcion_asesoria,
                    }
                )

        else:
            col_1, col_2 = st.columns(2)

            with col_1:
                cantidad_asesorias_ia = st.number_input(
                    "Cantidad de asesorías",
                    min_value=1,
                    max_value=30,
                    value=4,
                    step=1,
                    key="ejecucion_cantidad_asesorias_ia",
                )

                horas_por_asesoria_ia = st.number_input(
                    "Horas por asesoría",
                    min_value=0.5,
                    max_value=12.0,
                    value=2.0,
                    step=0.5,
                    key="ejecucion_horas_por_asesoria",
                )

                fecha_inicio = st.date_input(
                    "Fecha de inicio",
                    value=date.today(),
                    key="ejecucion_fecha_inicio",
                )

            with col_2:
                fecha_fin = st.date_input(
                    "Fecha de fin",
                    value=date.today() + timedelta(days=30),
                    key="ejecucion_fecha_fin",
                )

                dias_ejecucion = st.multiselect(
                    "Días de ejecución",
                    options=DIAS_SEMANA,
                    default=["Martes", "Jueves"],
                    key="ejecucion_dias",
                )

            descripcion_proyecto = st.text_area(
                "Descripción general del proyecto",
                placeholder=(
                    "Describe el proyecto para que la IA genere actividades "
                    "de asesoría coherentes."
                ),
                height=140,
            )

        generar_acta = st.form_submit_button(
            "Generar Acta de Ejecución"
        )

    if generar_acta:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Sublínea tecnológica": sublinea_tecnologica,
            "Nombre del experto": nombre_experto,
            "Nombre del talento": nombre_talento,
            "Documento del talento": documento_talento,
            "Teléfono del talento": telefono_talento,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        equipos_validos: list[dict] = []

        for item in equipos_materiales:
            tiene_contenido = (
                limpiar_texto(str(item.get("nombre", "")))
                or limpiar_texto(
                    str(item.get("cantidad_horas", ""))
                )
                or float(item.get("valor_total", 0)) > 0
            )

            if tiene_contenido:
                equipos_validos.append(
                    {
                        "nombre": limpiar_texto(
                            str(item.get("nombre", ""))
                        ),
                        "cantidad_horas": limpiar_texto(
                            str(item.get("cantidad_horas", ""))
                        ),
                        "valor_total": float(
                            item.get("valor_total", 0)
                        ),
                    }
                )

        if metodo_asesorias == "Generación manual":
            asesorias_validas = []

            for item in asesorias:
                descripcion = limpiar_texto(
                    str(item.get("descripcion", ""))
                )

                if descripcion:
                    asesorias_validas.append(
                        {
                            "fecha": str(item.get("fecha", "")),
                            "horas": float(item.get("horas", 0)),
                            "descripcion": descripcion,
                        }
                    )

            if not asesorias_validas:
                st.error(
                    "Debe ingresar al menos una asesoría manual "
                    "con descripción."
                )
                st.stop()

            modo_generacion = "Manual"

        else:
            if not limpiar_texto(descripcion_proyecto):
                st.error(
                    "Debe ingresar la descripción general del proyecto."
                )
                st.stop()

            if not dias_ejecucion:
                st.error(
                    "Debe seleccionar al menos un día de ejecución."
                )
                st.stop()

            if (
                fecha_inicio is None
                or fecha_fin is None
                or fecha_fin < fecha_inicio
            ):
                st.error(
                    "La fecha de fin no puede ser anterior "
                    "a la fecha de inicio."
                )
                st.stop()

            fechas_asesorias = distribuir_fechas_para_asesorias(
                fecha_inicio,
                fecha_fin,
                dias_ejecucion,
                int(cantidad_asesorias_ia),
            )

            if not fechas_asesorias:
                st.error(
                    "No se encontraron fechas válidas para distribuir "
                    "las asesorías."
                )
                st.stop()

            with st.spinner(
                "Generando asesorías de ejecución."
            ):
                if modo_prueba:
                    asesorias_validas = (
                        generar_asesorias_ejecucion_modo_prueba(
                            descripcion_proyecto,
                            int(cantidad_asesorias_ia),
                            fechas_asesorias,
                            float(horas_por_asesoria_ia),
                        )
                    )
                    modo_generacion = "Prueba local"
                else:
                    asesorias_validas = (
                        generar_asesorias_ejecucion_con_ia(
                            descripcion_proyecto,
                            int(cantidad_asesorias_ia),
                            fechas_asesorias,
                            float(horas_por_asesoria_ia),
                            modelo_openai,
                        )
                    )
                    modo_generacion = "ChatGPT API"

        total_equipos_materiales = sum(
            float(item.get("valor_total", 0))
            for item in equipos_validos
        )

        total_horas = sum(
            float(item.get("horas", 0))
            for item in asesorias_validas
        )

        total_honorarios = total_horas * VALOR_HORA_EXPERTO
        total_general = (
            total_equipos_materiales + total_honorarios
        )

        datos_acta_ejecucion = {
            "tipo_documento": "Acta de Ejecución",
            "titulo_acta": (
                f"Acta 2 - {limpiar_texto(codigo_proyecto)}"
            ),
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "sublinea_tecnologica": limpiar_texto(
                sublinea_tecnologica
            ),
            "nombre_experto": limpiar_texto(nombre_experto),
            "nombre_talento": limpiar_texto(nombre_talento),
            "documento_talento": limpiar_texto(
                documento_talento
            ),
            "telefono_talento": limpiar_texto(
                telefono_talento
            ),
            "equipos_materiales": equipos_validos,
            "asesorias": asesorias_validas,
            "metodo_asesorias": metodo_asesorias,
            "total_equipos_materiales": (
                total_equipos_materiales
            ),
            "total_honorarios": total_honorarios,
            "total_general": total_general,
            "valor_hora_experto": VALOR_HORA_EXPERTO,
            "modo_generacion": modo_generacion,
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin,
            "dias_ejecucion": dias_ejecucion,
            "version": VERSION_ACTA_EJECUCION,
        }

        st.session_state.datos_acta_ejecucion_generada = (
            datos_acta_ejecucion
        )
        st.session_state.ruta_pdf_acta_ejecucion_generado = None

        st.success(
            "Acta de ejecución generada correctamente. "
            "Ahora puedes revisar y generar el PDF."
        )

    datos = st.session_state.get(
        "datos_acta_ejecucion_generada"
    )

    if datos:
        st.markdown("## Resumen para validación")
        st.write("**Título:**", datos["titulo_acta"])
        st.write("**Proyecto:**", datos["nombre_proyecto"])
        st.write("**Código:**", datos["codigo_proyecto"])
        st.write("**Experto:**", datos["nombre_experto"])
        st.write("**Talento:**", datos["nombre_talento"])
        st.write(
            "**Método de asesorías:**",
            datos["metodo_asesorias"],
        )
        st.write(
            "**Modo de generación:**",
            datos["modo_generacion"],
        )
        st.write(
            "**Total equipos y materiales:**",
            formato_moneda_colombiana(
                datos["total_equipos_materiales"]
            ),
        )
        st.write(
            "**Total honorarios:**",
            formato_moneda_colombiana(
                datos["total_honorarios"]
            ),
        )
        st.write(
            "**Total general:**",
            formato_moneda_colombiana(
                datos["total_general"]
            ),
        )

        st.markdown("### Asesorías y usos")

        for indice, item in enumerate(
            datos["asesorias"],
            start=1,
        ):
            st.write(
                f"**{indice}.** {item['fecha']} - "
                f"{item['horas']} horas - "
                f"{item['descripcion']}"
            )

        st.markdown("### Equipos y materiales")

        if datos["equipos_materiales"]:
            for indice, item in enumerate(
                datos["equipos_materiales"],
                start=1,
            ):
                st.write(
                    f"**{indice}.** {item['nombre']} - "
                    f"{item['cantidad_horas']} - "
                    f"{formato_moneda_colombiana(item['valor_total'])}"
                )
        else:
            st.info(
                "No se registraron equipos o materiales."
            )

        col_json, col_pdf = st.columns(2)

        with col_json:
            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(
                    _serializar_datos(datos),
                    ensure_ascii=False,
                    indent=4,
                ),
                file_name="datos_acta_ejecucion.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button(
                "📄 Generar PDF del Acta de Ejecución",
                key="generar_pdf_acta_ejecucion",
            ):
                try:
                    ruta_pdf = generar_pdf_acta_ejecucion(
                        datos
                    )
                    st.session_state.ruta_pdf_acta_ejecucion_generado = (
                        ruta_pdf
                    )
                    st.success(
                        f"PDF generado correctamente: {ruta_pdf}"
                    )
                except Exception as error:
                    st.error(
                        f"No se pudo generar el PDF: {error}"
                    )

        ruta_pdf = st.session_state.get(
            "ruta_pdf_acta_ejecucion_generado"
        )

        if ruta_pdf and Path(ruta_pdf).exists():
            with open(ruta_pdf, "rb") as archivo_pdf:
                st.download_button(
                    label=(
                        "⬇️ Descargar PDF del Acta de Ejecución"
                    ),
                    data=archivo_pdf,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )
