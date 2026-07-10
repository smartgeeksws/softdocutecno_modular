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


VERSION_CRONOGRAMA = "VERSION_MODULAR_CRONOGRAMA_ACTIVIDADES_FORMATO_VALIDADO"
FORMATO_CRONOGRAMA = "TP-PEPBT V.1"


# =====================================================
# FECHAS Y DISTRIBUCIÓN DE ACTIVIDADES
# =====================================================

def normalizar_dia_semana(nombre_dia: str) -> int:
    mapa = {
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

    clave = str(nombre_dia).lower().strip()
    if clave not in mapa:
        raise ValueError(f"Día de la semana no válido: {nombre_dia}")

    return mapa[clave]


def obtener_fechas_programadas(
    fecha_inicio: date,
    fecha_fin: date,
    dias_semana: list[str],
) -> list[date]:
    dias_num = {normalizar_dia_semana(dia) for dia in dias_semana}
    fechas: list[date] = []
    actual = fecha_inicio

    while actual <= fecha_fin:
        if actual.weekday() in dias_num:
            fechas.append(actual)
        actual += timedelta(days=1)

    return fechas


def dividir_fechas_por_actividad(
    fechas: list[date],
    cantidad_actividades: int,
) -> list[list[date]]:
    if cantidad_actividades <= 0:
        return []

    if not fechas:
        return [[] for _ in range(cantidad_actividades)]

    bloques: list[list[date]] = []
    total_fechas = len(fechas)

    for indice in range(cantidad_actividades):
        inicio = round(indice * total_fechas / cantidad_actividades)
        fin = round((indice + 1) * total_fechas / cantidad_actividades)
        bloque = fechas[inicio:fin]

        if not bloque:
            bloque = [fechas[min(indice, total_fechas - 1)]]

        bloques.append(bloque)

    return bloques


def nombre_mes_es(fecha: date) -> str:
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
    return meses[fecha.month]


def agrupar_fechas_por_bloques_de_meses(
    fechas: list[date],
    max_meses_por_hoja: int = 3,
    max_fechas_por_hoja: int = 36,
) -> list[list[date]]:
    """Agrupa fechas para conservar legibilidad en cada hoja horizontal."""
    if not fechas:
        return []

    bloques: list[list[date]] = []
    bloque_actual: list[date] = []
    meses_actuales: list[tuple[int, int]] = []

    for fecha in sorted(fechas):
        clave_mes = (fecha.year, fecha.month)
        agrega_mes = clave_mes not in meses_actuales
        excede_meses = agrega_mes and len(meses_actuales) >= max_meses_por_hoja
        excede_fechas = len(bloque_actual) >= max_fechas_por_hoja

        if bloque_actual and (excede_meses or excede_fechas):
            bloques.append(bloque_actual)
            bloque_actual = []
            meses_actuales = []

        if clave_mes not in meses_actuales:
            meses_actuales.append(clave_mes)

        bloque_actual.append(fecha)

    if bloque_actual:
        bloques.append(bloque_actual)

    return bloques


# =====================================================
# GENERACIÓN DE ACTIVIDADES
# =====================================================

def generar_actividades_cronograma_modo_prueba(
    descripcion_proyecto: str,
    cantidad_actividades: int,
) -> list[str]:
    del descripcion_proyecto

    actividades_base = [
        "Revisión técnica y conceptual del proyecto",
        "Identificación de requerimientos técnicos, funcionales y operativos",
        "Definición de alternativas de diseño y criterios de selección",
        "Diseño preliminar de la solución propuesta",
        "Modelado, simulación o representación técnica de la solución",
        "Validación técnica de componentes, materiales o tecnologías",
        "Desarrollo, construcción o integración del prototipo",
        "Programación, configuración o ajuste de los componentes tecnológicos",
        "Realización de pruebas funcionales y registro de resultados",
        "Implementación de ajustes y mejoras sobre el prototipo",
        "Validación final de la solución con respecto a los requerimientos",
        "Elaboración de documentación técnica y consolidación de evidencias",
        "Socialización de resultados y entrega técnica del proyecto",
    ]

    actividades: list[str] = []
    for indice in range(cantidad_actividades):
        if indice < len(actividades_base):
            actividades.append(actividades_base[indice])
        else:
            actividades.append(
                f"Actividad técnica complementaria de validación y documentación {indice + 1}"
            )

    return actividades


def generar_actividades_cronograma_con_ia(
    descripcion_proyecto: str,
    cantidad_actividades: int,
    modelo: str = "gpt-4.1-mini",
) -> list[str]:
    instrucciones = """
Eres un experto en planeación de proyectos de base tecnológica de Tecnoparque SENA.

Genera actividades técnicas, concretas, secuenciales y verificables para un cronograma.
Cada actividad debe iniciar con un sustantivo de acción o un verbo en infinitivo.
Incluye, según corresponda, análisis de requerimientos, diseño, desarrollo, integración,
prototipado, programación, pruebas, validación, ajustes y documentación.
No incluyas fechas, nombres de personas, presupuestos ni información no suministrada.
No repitas actividades y evita descripciones demasiado generales.
Responde únicamente en JSON válido.
"""

    entrada = f"""
Genera exactamente {cantidad_actividades} actividades técnicas para el cronograma del siguiente proyecto.

Descripción del proyecto:
{descripcion_proyecto}

Formato obligatorio:
{{
  "actividades": [
    "Actividad 1",
    "Actividad 2"
  ]
}}
"""

    datos = generar_json_openai(
        instrucciones=instrucciones,
        entrada=entrada,
        modelo=modelo,
        temperature=0.35,
    )

    actividades_recibidas = datos.get("actividades", []) if isinstance(datos, dict) else []
    actividades_limpias = [
        limpiar_texto(str(actividad))
        for actividad in actividades_recibidas
        if str(actividad).strip()
    ]

    if len(actividades_limpias) < cantidad_actividades:
        respaldo = generar_actividades_cronograma_modo_prueba(
            descripcion_proyecto,
            cantidad_actividades,
        )
        actividades_limpias.extend(respaldo[len(actividades_limpias):])

    return actividades_limpias[:cantidad_actividades]


# =====================================================
# PDF HORIZONTAL TIPO GANTT
# =====================================================

def obtener_ruta_logo_tecnoparque() -> str | None:
    posibles_rutas = [
        Path("resources/logo_tecnoparque.png"),
        Path("resources/logo_tecnoparque.jpg"),
        Path("resources/logo_tecnoparque.jpeg"),
        Path("recursos/logo_tecnoparque.png"),
        Path("recursos/logo_tecnoparque.jpg"),
        Path("recursos/logo_tecnoparque.jpeg"),
        Path(RUTA_LOGO_SENA),
    ]

    for ruta in posibles_rutas:
        if ruta.exists():
            return str(ruta)

    return None


def _serializar_datos_cronograma(datos: dict) -> dict:
    resultado = dict(datos)

    if isinstance(resultado.get("fecha_inicio"), date):
        resultado["fecha_inicio"] = resultado["fecha_inicio"].strftime("%d/%m/%Y")

    if isinstance(resultado.get("fecha_fin"), date):
        resultado["fecha_fin"] = resultado["fecha_fin"].strftime("%d/%m/%Y")

    resultado["fechas_programadas"] = [
        fecha.strftime("%d/%m/%Y") if isinstance(fecha, date) else str(fecha)
        for fecha in resultado.get("fechas_programadas", [])
    ]

    resultado["bloques_fechas"] = [
        [
            fecha.strftime("%d/%m/%Y") if isinstance(fecha, date) else str(fecha)
            for fecha in bloque
        ]
        for bloque in resultado.get("bloques_fechas", [])
    ]

    return resultado


def generar_pdf_cronograma(datos: dict) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Cronograma_Actividades_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(OUTPUT_DIR / nombre_archivo)

    page_width, page_height = landscape(letter)

    def encabezado_pie(canvas_pdf, _doc) -> None:
        canvas_pdf.saveState()

        ruta_logo = obtener_ruta_logo_tecnoparque()
        if ruta_logo:
            try:
                logo = ImageReader(ruta_logo)
                canvas_pdf.drawImage(
                    logo,
                    0.75 * cm,
                    page_height - 1.55 * cm,
                    width=2.1 * cm,
                    height=1.05 * cm,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                pass

        canvas_pdf.setFillColor(colors.HexColor("#2E7D32"))
        canvas_pdf.setFont("Helvetica-Bold", 12)
        canvas_pdf.drawCentredString(
            page_width / 2,
            page_height - 1.05 * cm,
            "CRONOGRAMA DE ACTIVIDADES",
        )

        canvas_pdf.setFillColor(colors.grey)
        canvas_pdf.setFont("Helvetica", 8)
        canvas_pdf.drawCentredString(page_width / 2, 0.45 * cm, FORMATO_CRONOGRAMA)
        canvas_pdf.setFillColor(colors.black)
        canvas_pdf.restoreState()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=landscape(letter),
        leftMargin=0.65 * cm,
        rightMargin=0.65 * cm,
        topMargin=1.85 * cm,
        bottomMargin=0.85 * cm,
    )

    estilo_normal = ParagraphStyle(
        name="NormalCronograma",
        fontName="Helvetica",
        fontSize=7.0,
        leading=8.2,
        alignment=TA_LEFT,
    )

    estilo_negrita = ParagraphStyle(
        name="NegritaCronograma",
        fontName="Helvetica-Bold",
        fontSize=7.0,
        leading=8.2,
        alignment=TA_LEFT,
    )

    estilo_centro = ParagraphStyle(
        name="CentroCronograma",
        fontName="Helvetica-Bold",
        fontSize=5.2,
        leading=5.8,
        alignment=TA_CENTER,
    )

    estilo_actividad = ParagraphStyle(
        name="ActividadCronograma",
        fontName="Helvetica",
        fontSize=6.3,
        leading=7.3,
        alignment=TA_LEFT,
    )

    historia = []
    fechas_todas: list[date] = datos.get("fechas_programadas", [])
    actividades: list[str] = datos.get("actividades", [])
    bloques_fechas_actividad: list[list[date]] = datos.get("bloques_fechas", [])

    bloques_hojas = agrupar_fechas_por_bloques_de_meses(
        fechas_todas,
        max_meses_por_hoja=3,
        max_fechas_por_hoja=36,
    ) or [[]]

    ancho_disponible = page_width - doc.leftMargin - doc.rightMargin

    for numero_hoja, fechas_hoja in enumerate(bloques_hojas, start=1):
        if numero_hoja > 1:
            historia.append(PageBreak())

        if fechas_hoja:
            periodo_hoja = (
                f"{fechas_hoja[0].strftime('%d/%m/%Y')} al "
                f"{fechas_hoja[-1].strftime('%d/%m/%Y')}"
            )
        else:
            periodo_hoja = (
                f"{datos['fecha_inicio'].strftime('%d/%m/%Y')} al "
                f"{datos['fecha_fin'].strftime('%d/%m/%Y')}"
            )

        datos_generales = [
            [
                Paragraph("<b>NOMBRE DEL PROYECTO</b>", estilo_negrita),
                Paragraph(escape(str(datos.get("nombre_proyecto", ""))), estilo_normal),
                Paragraph("<b>NOMBRE DEL TALENTO</b>", estilo_negrita),
                Paragraph(escape(str(datos.get("nombre_talento", ""))), estilo_normal),
            ],
            [
                Paragraph("<b>CÓDIGO DEL PROYECTO</b>", estilo_negrita),
                Paragraph(escape(str(datos.get("codigo_proyecto", ""))), estilo_normal),
                Paragraph("<b>EXPERTO</b>", estilo_negrita),
                Paragraph(escape(str(datos.get("nombre_experto", ""))), estilo_normal),
            ],
            [
                Paragraph("<b>LÍNEA</b>", estilo_negrita),
                Paragraph(escape(str(datos.get("linea", ""))), estilo_normal),
                Paragraph("<b>TIEMPO DE EJECUCIÓN</b>", estilo_negrita),
                Paragraph(
                    f"{datos['fecha_inicio'].strftime('%d/%m/%Y')} al "
                    f"{datos['fecha_fin'].strftime('%d/%m/%Y')}",
                    estilo_normal,
                ),
            ],
            [
                Paragraph("<b>DÍAS PROGRAMADOS</b>", estilo_negrita),
                Paragraph(
                    escape(", ".join(datos.get("dias_semana", []))),
                    estilo_normal,
                ),
                Paragraph("<b>PERIODO MOSTRADO</b>", estilo_negrita),
                Paragraph(periodo_hoja, estilo_normal),
            ],
        ]

        ancho_etiqueta = 3.6 * cm
        ancho_valor_1 = ancho_disponible * 0.38
        ancho_valor_2 = ancho_disponible - (2 * ancho_etiqueta) - ancho_valor_1

        tabla_datos = Table(
            datos_generales,
            colWidths=[ancho_etiqueta, ancho_valor_1, ancho_etiqueta, ancho_valor_2],
        )
        tabla_datos.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E8F1E8")),
                    ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#E8F1E8")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        historia.append(tabla_datos)
        historia.append(Spacer(1, 0.25 * cm))

        fechas_hoja = list(fechas_hoja)
        encabezado_meses = [
            Paragraph("<b>N.º</b>", estilo_centro),
            Paragraph("<b>ACTIVIDAD</b>", estilo_centro),
        ]
        encabezado_dias = ["", ""]

        rangos_meses: list[tuple[int, int, str]] = []
        inicio_mes = 2
        mes_actual: tuple[int, int] | None = None

        for posicion, fecha_programada in enumerate(fechas_hoja, start=2):
            clave_mes = (fecha_programada.year, fecha_programada.month)

            if mes_actual is None:
                mes_actual = clave_mes
                inicio_mes = posicion
            elif clave_mes != mes_actual:
                fecha_referencia = fechas_hoja[posicion - 3]
                rangos_meses.append(
                    (inicio_mes, posicion - 1, f"{nombre_mes_es(fecha_referencia)} {fecha_referencia.year}")
                )
                mes_actual = clave_mes
                inicio_mes = posicion

            encabezado_meses.append("")
            encabezado_dias.append(
                Paragraph(
                    f"{fecha_programada.strftime('%a')[:2]}<br/>{fecha_programada.day}",
                    estilo_centro,
                )
            )

        if fechas_hoja:
            fecha_referencia = fechas_hoja[-1]
            rangos_meses.append(
                (
                    inicio_mes,
                    len(fechas_hoja) + 1,
                    f"{nombre_mes_es(fecha_referencia)} {fecha_referencia.year}",
                )
            )

        for inicio, _fin, titulo_mes in rangos_meses:
            encabezado_meses[inicio] = Paragraph(f"<b>{titulo_mes}</b>", estilo_centro)

        filas_cronograma = [encabezado_meses, encabezado_dias]

        for indice, actividad in enumerate(actividades, start=1):
            fechas_actividad = set(
                bloques_fechas_actividad[indice - 1]
                if indice - 1 < len(bloques_fechas_actividad)
                else []
            )

            fila = [
                Paragraph(str(indice), estilo_centro),
                Paragraph(escape(str(actividad)), estilo_actividad),
            ]

            for fecha_programada in fechas_hoja:
                marca = "X" if fecha_programada in fechas_actividad else ""
                fila.append(Paragraph(marca, estilo_centro))

            filas_cronograma.append(fila)

        ancho_numero = 0.75 * cm
        ancho_actividad = 7.0 * cm
        cantidad_fechas = max(len(fechas_hoja), 1)
        ancho_fecha = max(
            0.42 * cm,
            (ancho_disponible - ancho_numero - ancho_actividad) / cantidad_fechas,
        )

        anchos_columnas = [ancho_numero, ancho_actividad] + [
            ancho_fecha for _ in fechas_hoja
        ]

        tabla_cronograma = Table(
            filas_cronograma,
            colWidths=anchos_columnas,
            repeatRows=2,
        )

        estilos_tabla = [
            ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
            ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#D9EAD3")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (2, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]

        for inicio, fin, _titulo_mes in rangos_meses:
            estilos_tabla.append(("SPAN", (inicio, 0), (fin, 0)))
            estilos_tabla.append(("ALIGN", (inicio, 0), (fin, 0), "CENTER"))

        for indice_actividad, bloque in enumerate(bloques_fechas_actividad, start=2):
            fechas_bloque = set(bloque)
            for indice_fecha, fecha_programada in enumerate(fechas_hoja, start=2):
                if fecha_programada in fechas_bloque:
                    estilos_tabla.append(
                        (
                            "BACKGROUND",
                            (indice_fecha, indice_actividad),
                            (indice_fecha, indice_actividad),
                            colors.HexColor("#93C47D"),
                        )
                    )

        tabla_cronograma.setStyle(TableStyle(estilos_tabla))
        historia.append(tabla_cronograma)

    doc.build(
        historia,
        onFirstPage=encabezado_pie,
        onLaterPages=encabezado_pie,
    )

    datos_json = _serializar_datos_cronograma(datos)
    datos_json["ruta_pdf"] = ruta_pdf
    guardar_datos_json(datos_json, nombre_archivo="datos_cronograma_actividades.json")

    return ruta_pdf


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_cronograma(
    modo_prueba: bool = True,
    modelo_openai: str = "gpt-4.1-mini",
) -> None:
    st.markdown("---")
    st.subheader("Formulario para Cronograma de Actividades")
    st.caption(VERSION_CRONOGRAMA)

    st.info(
        "Este módulo genera un cronograma horizontal tipo diagrama de Gantt. "
        "Las actividades se construyen según la descripción técnica del proyecto."
    )

    if "datos_cronograma_generado" not in st.session_state:
        st.session_state.datos_cronograma_generado = None

    if "ruta_pdf_cronograma_generado" not in st.session_state:
        st.session_state.ruta_pdf_cronograma_generado = None

    with st.form("form_cronograma"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2026-143440-00001",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Nombre oficial del proyecto",
                height=90,
            )

            nombre_talento = st.text_input(
                "Nombre del talento",
                placeholder="Nombre completo del talento",
            )

            nombre_experto = st.text_input(
                "Nombre del experto",
                placeholder="Nombre completo del experto",
            )

        with col_b:
            linea = st.text_input(
                "Línea",
                placeholder="Ejemplo: Ingeniería y diseño",
            )

            cantidad_actividades = st.number_input(
                "Cantidad de actividades",
                min_value=3,
                max_value=20,
                value=7,
                step=1,
            )

            fecha_inicio = st.date_input(
                "Fecha de inicio",
                value=date.today(),
            )

            fecha_fin = st.date_input(
                "Fecha de finalización",
                value=date.today() + timedelta(days=60),
            )

            dias_semana = st.multiselect(
                "Día(s) de la semana para programar actividades",
                options=[
                    "Lunes",
                    "Martes",
                    "Miércoles",
                    "Jueves",
                    "Viernes",
                    "Sábado",
                    "Domingo",
                ],
                default=["Sábado", "Domingo"],
            )

        descripcion_proyecto = st.text_area(
            "Describe el proyecto",
            placeholder=(
                "Describe la necesidad, la solución propuesta, los componentes técnicos, "
                "el prototipo y el resultado esperado."
            ),
            height=160,
        )

        generar_cronograma = st.form_submit_button(
            "Generar cronograma de actividades"
        )

    if generar_cronograma:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Nombre del talento": nombre_talento,
            "Nombre del experto": nombre_experto,
            "Línea": linea,
            "Descripción del proyecto": descripcion_proyecto,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        if not dias_semana:
            st.error("Selecciona al menos un día de la semana.")
            st.stop()

        if fecha_fin < fecha_inicio:
            st.error("La fecha de finalización no puede ser anterior a la fecha de inicio.")
            st.stop()

        fechas_programadas = obtener_fechas_programadas(
            fecha_inicio,
            fecha_fin,
            dias_semana,
        )

        if not fechas_programadas:
            st.error(
                "No se encontraron fechas programadas con los días seleccionados "
                "dentro del rango indicado."
            )
            st.stop()

        with st.spinner("Generando actividades técnicas para el cronograma..."):
            try:
                if modo_prueba:
                    actividades = generar_actividades_cronograma_modo_prueba(
                        descripcion_proyecto,
                        int(cantidad_actividades),
                    )
                else:
                    actividades = generar_actividades_cronograma_con_ia(
                        descripcion_proyecto,
                        int(cantidad_actividades),
                        modelo_openai,
                    )
            except Exception as error:
                st.error(f"No se pudieron generar las actividades: {error}")
                st.stop()

        bloques_fechas = dividir_fechas_por_actividad(
            fechas_programadas,
            int(cantidad_actividades),
        )

        datos_cronograma = {
            "tipo_documento": "Cronograma de actividades",
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "nombre_talento": limpiar_texto(nombre_talento),
            "nombre_experto": limpiar_texto(nombre_experto),
            "linea": limpiar_texto(linea),
            "cantidad_actividades": int(cantidad_actividades),
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin,
            "dias_semana": dias_semana,
            "descripcion_proyecto": limpiar_texto(descripcion_proyecto),
            "actividades": actividades,
            "fechas_programadas": fechas_programadas,
            "bloques_fechas": bloques_fechas,
            "modo_generacion": "Prueba local" if modo_prueba else "ChatGPT API",
            "version": VERSION_CRONOGRAMA,
        }

        st.session_state.datos_cronograma_generado = datos_cronograma
        st.session_state.ruta_pdf_cronograma_generado = None

        st.success(
            "Cronograma generado correctamente. Ahora puedes revisarlo y generar el PDF."
        )

    datos_cronograma = st.session_state.get("datos_cronograma_generado")

    if datos_cronograma:
        st.markdown("## Resumen para validación")
        st.write("**Modo de generación:**", datos_cronograma["modo_generacion"])
        st.write("**Código del proyecto:**", datos_cronograma["codigo_proyecto"])
        st.write("**Nombre del proyecto:**", datos_cronograma["nombre_proyecto"])
        st.write("**Talento:**", datos_cronograma["nombre_talento"])
        st.write("**Experto:**", datos_cronograma["nombre_experto"])
        st.write("**Línea:**", datos_cronograma["linea"])
        st.write(
            "**Periodo:**",
            f"{datos_cronograma['fecha_inicio'].strftime('%d/%m/%Y')} al "
            f"{datos_cronograma['fecha_fin'].strftime('%d/%m/%Y')}",
        )
        st.write(
            "**Días programados:**",
            ", ".join(datos_cronograma["dias_semana"]),
        )

        st.markdown("### Actividades generadas")
        for indice, actividad in enumerate(datos_cronograma["actividades"], start=1):
            fechas_actividad = datos_cronograma["bloques_fechas"][indice - 1]

            if fechas_actividad:
                periodo = (
                    f"{fechas_actividad[0].strftime('%d/%m/%Y')} al "
                    f"{fechas_actividad[-1].strftime('%d/%m/%Y')}"
                )
            else:
                periodo = "Sin fecha asignada"

            st.write(f"**{indice}.** {actividad} - {periodo}")

        col_json, col_pdf = st.columns(2)

        with col_json:
            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(
                    _serializar_datos_cronograma(datos_cronograma),
                    ensure_ascii=False,
                    indent=4,
                ),
                file_name="datos_cronograma_actividades.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button(
                "📄 Generar PDF del cronograma",
                key="generar_pdf_cronograma",
            ):
                try:
                    ruta_pdf = generar_pdf_cronograma(datos_cronograma)
                    st.session_state.ruta_pdf_cronograma_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as error:
                    st.error(f"No se pudo generar el PDF: {error}")

        ruta_pdf = st.session_state.get("ruta_pdf_cronograma_generado")

        if ruta_pdf and Path(ruta_pdf).exists():
            with open(ruta_pdf, "rb") as archivo_pdf:
                st.download_button(
                    label="⬇️ Descargar PDF del cronograma",
                    data=archivo_pdf,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )
