from pathlib import Path
from datetime import datetime, date, time, timedelta
import json

import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

from config.constants import (
    OUTPUT_DIR,
    RUTA_LOGO_SENA,
    FORMATO_CODIGO,
    LUGAR_ENLACE_DEFAULT,
    DIRECCION_REGIONAL_CENTRO_DEFAULT,
    DEPENDENCIA_TALENTO_DEFAULT,
    DEPENDENCIA_EXPERTO_DEFAULT,
    ANEXOS_DEFAULT,
)
from services.openai_service import generar_json_openai
from services.json_service import guardar_datos_json
from utils.nombres_archivo import safe_filename
from utils.validaciones import validar_campos_obligatorios
from utils.textos import limpiar_texto


VERSION_ACTA_INICIO = "VERSION_MODULAR_ACTA_INICIO_FORMATO_INSTITUCIONAL_VALIDADO"


def calcular_hora_fin(fecha_acta: date, hora_inicio: time) -> time:
    inicio_dt = datetime.combine(fecha_acta, hora_inicio)
    return (inicio_dt + timedelta(minutes=30)).time()


def construir_objetivo_reunion(nombre_proyecto: str) -> str:
    return f"Fijar el alcance, objetivo general y objetivos específicos del proyecto: {nombre_proyecto}"


def generar_textos_modo_prueba(nombre_proyecto: str, descripcion_proyecto: str) -> dict:
    return {
        "objetivo_general": (
            f"Desarrollar una propuesta técnica para el proyecto {nombre_proyecto}, orientada a la definición, "
            "estructuración y validación inicial de una solución de base tecnológica que responda a la necesidad identificada."
        ),
        "objetivos_especificos": [
            "Identificar los requerimientos técnicos, funcionales y operativos asociados a la necesidad planteada en la descripción del proyecto.",
            "Definir los componentes principales de la solución tecnológica, considerando los recursos disponibles y las condiciones de uso previstas.",
            "Estructurar una ruta inicial de desarrollo para orientar las actividades de diseño, prototipado, validación básica y documentación técnica.",
            "Validar preliminarmente el alcance del proyecto y los entregables esperados, de acuerdo con la información suministrada.",
        ],
        "alcance": (
            f"El alcance del proyecto {nombre_proyecto} contempla la asesoría técnica inicial para la definición de la necesidad, "
            "la identificación de requerimientos, la estructuración de objetivos y la delimitación de actividades orientadas al desarrollo "
            "de una solución de base tecnológica. Incluye lineamientos para diseño, prototipado, validación básica y documentación."
        ),
    }


def generar_textos_con_ia(
    nombre_proyecto: str,
    descripcion_proyecto: str,
    modelo_openai: str = "gpt-4.1-mini",
) -> dict:
    instrucciones = """
Eres un experto en formulación de proyectos de base tecnológica, innovación,
desarrollo tecnológico y prototipado en el marco de Tecnoparque SENA.

Redacta en lenguaje formal, técnico e institucional.
No inventes fechas, códigos, nombres de personas, entidades o información no suministrada.
Los objetivos deben iniciar con verbos en infinitivo.
El objetivo general debe ser una sola oración clara.
Los objetivos específicos deben ser exactamente cuatro.
El alcance debe explicar qué incluye el proyecto, qué se desarrollará y cuáles son los entregables esperados.
No uses lenguaje comercial ni promesas no verificables.
Responde únicamente en JSON válido, sin texto antes ni después.
"""

    entrada = f"""
Genera los siguientes campos para un acta de inicio de proyecto:

1. objetivo_general
2. objetivos_especificos: exactamente 4 objetivos específicos
3. alcance

Nombre del proyecto:
{nombre_proyecto}

Descripción general del proyecto:
{descripcion_proyecto}

Formato obligatorio:
{{
  "objetivo_general": "...",
  "objetivos_especificos": ["...", "...", "...", "..."],
  "alcance": "..."
}}
"""

    datos = generar_json_openai(
        instrucciones=instrucciones,
        entrada=entrada,
        modelo=modelo_openai,
        temperature=0.3,
    )

    if not isinstance(datos, dict):
        return generar_textos_modo_prueba(nombre_proyecto, descripcion_proyecto)

    if not isinstance(datos.get("objetivos_especificos", []), list):
        datos["objetivos_especificos"] = []

    respaldo = generar_textos_modo_prueba(nombre_proyecto, descripcion_proyecto)

    if not str(datos.get("objetivo_general", "")).strip():
        datos["objetivo_general"] = respaldo["objetivo_general"]

    if len(datos.get("objetivos_especificos", [])) < 4:
        datos["objetivos_especificos"] = (
            datos.get("objetivos_especificos", []) + respaldo["objetivos_especificos"]
        )[:4]

    if not str(datos.get("alcance", "")).strip():
        datos["alcance"] = respaldo["alcance"]

    return datos


# =====================================================
# FUNCIONES GRÁFICAS PDF - FORMATO INSTITUCIONAL
# =====================================================

def wrap_text(texto: str, font_name: str, font_size: float, max_width: float) -> list[str]:
    texto = str(texto or "").replace("\n", " ").strip()
    if not texto:
        return [""]

    palabras = texto.split()
    lineas = []
    linea = ""

    for palabra in palabras:
        intento = palabra if not linea else f"{linea} {palabra}"
        if stringWidth(intento, font_name, font_size) <= max_width:
            linea = intento
        else:
            if linea:
                lineas.append(linea)
            while stringWidth(palabra, font_name, font_size) > max_width and len(palabra) > 5:
                palabra = palabra[:-1]
            linea = palabra

    if linea:
        lineas.append(linea)

    return lineas


def calcular_font_para_celda(
    texto,
    w,
    h,
    font="Helvetica",
    size=7.5,
    min_size=4.8,
    leading_factor=1.18,
    label=None,
):
    padding_x = 5
    padding_y = 5
    available_w = max(w - padding_x * 2, 10)
    available_h = max(h - padding_y * 2, 8)

    current = size
    while current >= min_size:
        if label:
            label_width = stringWidth(label, "Helvetica-Bold", current)
            if label_width >= available_w * 0.85:
                label_lines = len(wrap_text(label, "Helvetica-Bold", current, available_w))
                text_width = available_w
            else:
                label_lines = 0
                text_width = max(available_w - label_width - 3, 10)
        else:
            label_lines = 0
            text_width = available_w

        lineas = wrap_text(texto, font, current, text_width)
        leading = current * leading_factor
        needed_h = max(1, len(lineas) + label_lines) * leading
        if needed_h <= available_h:
            return current, leading
        current -= 0.25

    return min_size, min_size * leading_factor


def draw_wrapped_text(c, texto, x, y, w, h, font="Helvetica", size=7.5, label=None, center=False):
    padding_x = 5
    padding_y = 5
    texto = str(texto or "").strip()
    label_text = str(label or "").strip()

    font_size, leading = calcular_font_para_celda(
        texto,
        w,
        h,
        font=font,
        size=size,
        label=label_text or None,
    )
    cursor_y = y + h - padding_y - font_size
    max_w = w - padding_x * 2

    if center and not label_text:
        lineas = wrap_text(texto, font, font_size, max_w)
        total_h = len(lineas) * leading
        cursor_y = y + (h + total_h) / 2 - font_size
        c.setFont(font, font_size)
        for linea in lineas:
            c.drawCentredString(x + w / 2, cursor_y, linea)
            cursor_y -= leading
        return

    if label_text:
        label_width = stringWidth(label_text, "Helvetica-Bold", font_size)
        if label_width < max_w * 0.80:
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(x + padding_x, cursor_y, label_text)
            c.setFont(font, font_size)
            text_x = x + padding_x + label_width + 3
            text_w = max_w - label_width - 3
            lineas = wrap_text(texto, font, font_size, text_w)
            if lineas:
                c.drawString(text_x, cursor_y, lineas[0])
                cursor_y -= leading
                lineas = lineas[1:]
            for linea in lineas:
                if cursor_y < y + 2:
                    break
                c.drawString(x + padding_x, cursor_y, linea)
                cursor_y -= leading
        else:
            c.setFont("Helvetica-Bold", font_size)
            for linea in wrap_text(label_text, "Helvetica-Bold", font_size, max_w):
                if cursor_y < y + 2:
                    break
                c.drawString(x + padding_x, cursor_y, linea)
                cursor_y -= leading
            c.setFont(font, font_size)
            for linea in wrap_text(texto, font, font_size, max_w):
                if cursor_y < y + 2:
                    break
                c.drawString(x + padding_x, cursor_y, linea)
                cursor_y -= leading
        return

    c.setFont(font, font_size)
    for linea in wrap_text(texto, font, font_size, max_w):
        if cursor_y < y + 2:
            break
        c.drawString(x + padding_x, cursor_y, linea)
        cursor_y -= leading


def draw_cell(c, x, y, w, h, texto="", label=None, font="Helvetica", size=7.5, center=False, fill=None):
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.55)
    if fill:
        c.setFillColor(fill)
        c.rect(x, y, w, h, stroke=1, fill=1)
        c.setFillColor(colors.black)
    else:
        c.rect(x, y, w, h, stroke=1, fill=0)

    draw_wrapped_text(c, texto, x, y, w, h, font=font, size=size, label=label, center=center)


def draw_logo(c, page_width, top_y):
    logo_path = Path(RUTA_LOGO_SENA)

    if logo_path.exists():
        try:
            img = ImageReader(str(logo_path))
            logo_w = 52
            logo_h = 44
            logo_x = page_width / 2 - logo_w / 2
            logo_y = top_y - logo_h
            c.drawImage(img, logo_x, logo_y, width=logo_w, height=logo_h, mask="auto")
            return
        except Exception:
            pass

    c.setFillColor(colors.HexColor("#69B342"))
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(page_width / 2, top_y - 18, "SENA")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(page_width / 2, top_y - 36, "▲")
    c.setFillColor(colors.black)


def footer_codigo(c, page_width):
    c.setFillColor(colors.grey)
    c.setFont("Helvetica", 11)
    c.drawCentredString(page_width / 2, 20, FORMATO_CODIGO)
    c.setFillColor(colors.black)


def generar_pdf_acta_inicio(datos: dict) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    nombre_archivo = f"Acta_Inicio_{safe_filename(datos.get('codigo_acta', 'sin_codigo'))}.pdf"
    ruta_pdf = str(OUTPUT_DIR / nombre_archivo)

    c = canvas.Canvas(ruta_pdf, pagesize=letter)
    page_width, _page_height = letter

    x0 = 24
    table_w = page_width - 48
    logo_top_y = 785
    y_top_content = 705
    y_safe_bottom = 62

    FONT_TITLE = 8.5
    FONT_SECTION = 8.0
    FONT_BODY = 7.2
    FONT_SMALL = 6.8
    FONT_TINY = 6.3

    def iniciar_pagina() -> float:
        draw_logo(c, page_width, logo_top_y)
        return y_top_content

    def cerrar_pagina() -> None:
        footer_codigo(c, page_width)
        c.showPage()

    def asegurar_espacio(y_actual: float, alto_requerido: float) -> float:
        if y_actual - alto_requerido < y_safe_bottom:
            cerrar_pagina()
            return iniciar_pagina()
        return y_actual

    def alto_texto(
        texto: str,
        ancho: float,
        base: float = 24,
        font_size: float = FONT_BODY,
        label: str | None = None,
        max_h: float = 70,
    ) -> float:
        padding = 10
        usable_w = max(ancho - padding, 20)
        texto_total = f"{label or ''} {texto or ''}".strip()
        lineas = wrap_text(texto_total, "Helvetica", font_size, usable_w)
        h = max(base, len(lineas) * (font_size + 2) + 12)
        return min(h, max_h)

    y = iniciar_pagina()

    h = 22
    draw_cell(
        c,
        x0,
        y - h,
        table_w,
        h,
        f"ACTA NO. 1 DEL PROYECTO: {datos.get('codigo_acta', '')}",
        font="Helvetica-Bold",
        size=FONT_TITLE,
        center=True,
    )
    y -= h

    nombre_comite = f"Acta de inicio del proyecto {datos.get('codigo_acta', '')} - {datos.get('nombre_proyecto', '')}"
    h = alto_texto(nombre_comite, table_w, base=42, font_size=FONT_BODY, label="NOMBRE DEL COMITÉ O DE LA REUNIÓN:", max_h=58)
    draw_cell(c, x0, y - h, table_w, h, nombre_comite, label="NOMBRE DEL COMITÉ O DE LA REUNIÓN:", size=FONT_BODY)
    y -= h

    h = 38
    w1 = table_w * 0.58
    w2 = table_w * 0.21
    w3 = table_w * 0.21
    draw_cell(c, x0, y - h, w1, h, f"{datos.get('ciudad', '')} (Huila) - {datos.get('fecha_iso', datos.get('fecha_acta', ''))}", label="CIUDAD Y FECHA:", size=FONT_BODY)
    draw_cell(c, x0 + w1, y - h, w2, h, datos.get("hora_inicio", ""), label="HORA INICIO:", size=FONT_BODY)
    draw_cell(c, x0 + w1 + w2, y - h, w3, h, datos.get("hora_fin", ""), label="HORA FIN:", size=FONT_BODY)
    y -= h

    h = 46
    w_lugar = table_w * 0.58
    w_dir = table_w * 0.42
    draw_cell(c, x0, y - h, w_lugar, h, LUGAR_ENLACE_DEFAULT, label="LUGAR Y/O ENLACE:", size=FONT_BODY)
    draw_cell(c, x0 + w_lugar, y - h, w_dir, h, DIRECCION_REGIONAL_CENTRO_DEFAULT, label="DIRECCIÓN / REGIONAL / CENTRO:", size=FONT_BODY)
    y -= h

    agenda = (
        "1. Caracterización del proyecto de acuerdo con los objetivos y alcance propuestos.\n"
        "2. Documentación que soportan el inicio del proyecto."
    )
    h = 42
    draw_cell(c, x0, y - h, table_w, h, agenda, label="AGENDA O PUNTOS PARA DESARROLLAR:", size=FONT_BODY)
    y -= h

    h = alto_texto(datos.get("objetivo_reunion", ""), table_w, base=44, font_size=FONT_BODY, label="OBJETIVO(S) DE LA REUNIÓN:", max_h=60)
    draw_cell(c, x0, y - h, table_w, h, datos.get("objetivo_reunion", ""), label="OBJETIVO(S) DE LA REUNIÓN:", size=FONT_BODY)
    y -= h

    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "DESARROLLO DE LA REUNIÓN", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    codigo_nombre = f"{datos.get('codigo_acta', '')} - {datos.get('nombre_proyecto', '')}"
    h = alto_texto(codigo_nombre, table_w, base=26, font_size=FONT_BODY, label="Código y nombre del Proyecto:", max_h=48)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, codigo_nombre, label="Código y nombre del Proyecto:", size=FONT_BODY)
    y -= h

    h = alto_texto(datos.get("linea_sublinea", ""), table_w, base=24, font_size=FONT_BODY, label="Linea y sublinea:", max_h=38)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, datos.get("linea_sublinea", ""), label="Linea y sublinea:", size=FONT_BODY)
    y -= h

    bloque_talentos = 22 + 22 + 24
    y = asegurar_espacio(y, bloque_talentos)
    draw_cell(c, x0, y - 22, table_w, 22, "TALENTOS QUE PARTICIPAN EN EL PROYECTO", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= 22
    inter_w = table_w * 0.18
    draw_cell(c, x0, y - 22, inter_w, 22, "Interlocutor", font="Helvetica-Bold", size=FONT_SMALL)
    draw_cell(c, x0 + inter_w, y - 22, table_w - inter_w, 22, "Talento", font="Helvetica-Bold", size=FONT_SMALL)
    y -= 22
    draw_cell(c, x0, y - 24, inter_w, 24, "SI", size=FONT_SMALL)
    draw_cell(c, x0 + inter_w, y - 24, table_w - inter_w, 24, datos.get("nombre_talento", ""), size=FONT_SMALL)
    y -= 24

    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "OBJETIVOS DEL PROYECTO Y ALCANCE", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    h = alto_texto(datos.get("objetivo_general", ""), table_w, base=46, font_size=FONT_BODY, label="OBJETIVO GENERAL:", max_h=70)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, datos.get("objetivo_general", ""), label="OBJETIVO GENERAL:", size=FONT_BODY)
    y -= h

    h = 21
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "OBJETIVOS ESPECÍFICOS", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    objetivos = list(datos.get("objetivos_especificos", []) or [])[:4]
    while len(objetivos) < 4:
        objetivos.append("")

    num_w = 52
    obj_w = table_w - num_w
    for idx, obj in enumerate(objetivos, start=1):
        h = alto_texto(obj, obj_w, base=24, font_size=FONT_SMALL, max_h=42)
        y = asegurar_espacio(y, h)
        draw_cell(c, x0, y - h, num_w, h, str(idx), font="Helvetica-Bold", size=FONT_SMALL, center=True)
        draw_cell(c, x0 + num_w, y - h, obj_w, h, obj, size=FONT_SMALL)
        y -= h

    h = alto_texto(datos.get("alcance", ""), table_w, base=36, font_size=FONT_TINY, label="ALCANCE DEL PROYECTO:", max_h=82)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, datos.get("alcance", ""), label="ALCANCE DEL PROYECTO:", size=FONT_TINY)
    y -= h

    bloque_conclusiones = 22 + 28
    y = asegurar_espacio(y, bloque_conclusiones)
    draw_cell(c, x0, y - 22, table_w, 22, "CONCLUSIONES", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= 22
    conclusion = "Se fijan alcance y objetivos entre el talento y experto participantes de la reunión."
    draw_cell(c, x0, y - 28, table_w, 28, conclusion, size=FONT_BODY)
    y -= 28

    asistentes_alto = 22 + 34 + 34 + 34
    y = asegurar_espacio(y, asistentes_alto)
    draw_cell(c, x0, y - 22, table_w, 22, "DE: ASISTENTES Y APROBACIÓN DE DECISIONES:", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= 22

    col_w = [table_w * 0.18, table_w * 0.17, table_w * 0.20, table_w * 0.22, table_w * 0.23]
    headers = ["NOMBRE", "DEPENDENCIA / EMPRESA", "APRUEBA (SI/NO)", "OBSERVACIÓN", "FIRMA O PARTICIPACIÓN VIRTUAL"]
    x = x0
    for w, h_text in zip(col_w, headers):
        draw_cell(c, x, y - 34, w, 34, h_text, font="Helvetica-Bold", size=FONT_SMALL, center=True)
        x += w
    y -= 34

    filas = [
        [datos.get("nombre_talento", ""), DEPENDENCIA_TALENTO_DEFAULT, "SI", "", ""],
        [datos.get("nombre_experto", ""), DEPENDENCIA_EXPERTO_DEFAULT, "SI", "", ""],
    ]

    for fila in filas:
        x = x0
        for w, value in zip(col_w, fila):
            center = value in [DEPENDENCIA_TALENTO_DEFAULT, DEPENDENCIA_EXPERTO_DEFAULT, "SI", ""]
            draw_cell(c, x, y - 34, w, 34, value, font="Helvetica-Bold" if x == x0 else "Helvetica", size=FONT_SMALL, center=center)
            x += w
        y -= 34

    bloque_final_alto = 55 + 32
    y = asegurar_espacio(y, bloque_final_alto)

    proteccion = (
        "De acuerdo con La Ley 1581 de 2012, Protección de Datos Personales, el Servicio Nacional de Aprendizaje SENA, "
        "se compromete a garantizar la seguridad y protección de los datos personales que se encuentran almacenados en este documento, "
        "y les dará el tratamiento correspondiente en cumplimiento de lo establecido legalmente."
    )
    draw_cell(c, x0, y - 55, table_w, 55, proteccion, size=FONT_BODY)
    y -= 55

    texto_anexos = "ANEXOS\n" + ANEXOS_DEFAULT
    draw_cell(c, x0, y - 32, table_w, 32, texto_anexos, font="Helvetica-Bold", size=FONT_SECTION, center=True)

    footer_codigo(c, page_width)
    c.save()

    guardar_datos_json({**datos, "ruta_pdf": ruta_pdf}, nombre_archivo="datos_acta_inicio.json")
    return ruta_pdf


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_acta_inicio(modo_prueba: bool = True, modelo_openai: str = "gpt-4.1-mini") -> None:
    st.markdown("---")
    st.subheader("Formulario para Acta de Inicio")
    st.caption(VERSION_ACTA_INICIO)

    st.info(
        "Este módulo genera el Acta de Inicio con el formato institucional validado en la versión inicial del software."
    )

    with st.form("form_acta_inicio"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_acta = st.text_input(
                "Código o número del acta / proyecto",
                placeholder="Ejemplo: P2026-143440-00001",
            )
            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Ejemplo: Diseño de un sistema electrónico para medición.",
                height=90,
            )
            fecha_acta = st.date_input("Fecha del acta", value=date.today())
            hora_inicio = st.time_input("Hora de inicio", value=time(8, 0))
            ciudad = st.text_input("Ciudad", value="Campoalegre")

        with col_b:
            nombre_talento = st.text_input("Nombre del talento", placeholder="Nombre completo del talento")
            nombre_experto = st.text_input("Nombre del experto", placeholder="Nombre completo del experto")
            linea_sublinea = st.text_input(
                "Línea y sublínea de Tecnoparque",
                placeholder="Ejemplo: Productos y procesos / Diseño de producto",
            )

        descripcion_proyecto = st.text_area(
            "Descripción general del proyecto",
            placeholder="Describe la necesidad, el problema, la solución propuesta, los componentes tecnológicos y el resultado esperado.",
            height=170,
        )

        texto_boton = "Generar resumen del acta en modo prueba" if modo_prueba else "Generar resumen del acta con ChatGPT"
        generar = st.form_submit_button(texto_boton)

    if generar:
        campos_obligatorios = {
            "Código o número del acta / proyecto": codigo_acta,
            "Nombre del proyecto": nombre_proyecto,
            "Nombre del talento": nombre_talento,
            "Nombre del experto": nombre_experto,
            "Línea y sublínea": linea_sublinea,
            "Descripción general del proyecto": descripcion_proyecto,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        hora_fin = calcular_hora_fin(fecha_acta, hora_inicio)
        objetivo_reunion = construir_objetivo_reunion(nombre_proyecto)

        datos_acta = {
            "tipo_documento": "Acta de inicio",
            "codigo_acta": limpiar_texto(codigo_acta),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "ciudad": limpiar_texto(ciudad),
            "fecha_acta": fecha_acta.strftime("%d/%m/%Y"),
            "fecha_iso": fecha_acta.strftime("%Y-%m-%d"),
            "hora_inicio": hora_inicio.strftime("%H:%M"),
            "hora_fin": hora_fin.strftime("%H:%M"),
            "nombre_talento": limpiar_texto(nombre_talento),
            "nombre_experto": limpiar_texto(nombre_experto),
            "linea_sublinea": limpiar_texto(linea_sublinea),
            "descripcion_proyecto": limpiar_texto(descripcion_proyecto),
            "objetivo_reunion": objetivo_reunion,
            "version": VERSION_ACTA_INICIO,
        }

        mensaje_spinner = "Generando textos en modo prueba." if modo_prueba else "Generando objetivo general, objetivos específicos y alcance con ChatGPT."
        with st.spinner(mensaje_spinner):
            try:
                if modo_prueba:
                    textos_ia = generar_textos_modo_prueba(nombre_proyecto, descripcion_proyecto)
                else:
                    textos_ia = generar_textos_con_ia(nombre_proyecto, descripcion_proyecto, modelo_openai)

                datos_acta["objetivo_general"] = limpiar_texto(textos_ia.get("objetivo_general", ""))
                datos_acta["objetivos_especificos"] = [
                    limpiar_texto(obj) for obj in textos_ia.get("objetivos_especificos", [])[:4]
                ]
                datos_acta["alcance"] = limpiar_texto(textos_ia.get("alcance", ""))
                datos_acta["modo_generacion"] = "Prueba local" if modo_prueba else "ChatGPT API"

            except Exception as e:
                st.error(f"No se pudo generar el contenido: {e}")
                st.stop()

        st.session_state.datos_acta_inicio_generada = datos_acta
        st.session_state.ruta_pdf_acta_inicio_generado = None
        st.success("Resumen generado correctamente. La hora de finalización fue calculada automáticamente.")

    if st.session_state.get("datos_acta_inicio_generada"):
        datos_acta = st.session_state.datos_acta_inicio_generada

        st.markdown("## Resumen para validación")
        st.write("**Modo de generación:**", datos_acta["modo_generacion"])
        st.write("**Código / Acta:**", datos_acta["codigo_acta"])
        st.write("**Nombre del proyecto:**", datos_acta["nombre_proyecto"])
        st.write("**Ciudad y fecha:**", f'{datos_acta["ciudad"]}, {datos_acta["fecha_acta"]}')
        st.write("**Hora de inicio:**", datos_acta["hora_inicio"])
        st.write("**Hora de finalización:**", datos_acta["hora_fin"])
        st.write("**Talento:**", datos_acta["nombre_talento"])
        st.write("**Experto:**", datos_acta["nombre_experto"])
        st.write("**Línea y sublínea:**", datos_acta["linea_sublinea"])
        st.write("**Objetivo de la reunión:**", datos_acta["objetivo_reunion"])
        st.write("**Descripción general:**", datos_acta["descripcion_proyecto"])

        st.markdown("## Textos generados")
        st.markdown("### Objetivo general")
        st.write(datos_acta["objetivo_general"])

        st.markdown("### Objetivos específicos")
        if datos_acta["objetivos_especificos"]:
            for i, objetivo in enumerate(datos_acta["objetivos_especificos"], start=1):
                st.write(f"{i}. {objetivo}")
        else:
            st.warning("No se generaron objetivos específicos en formato de lista.")

        st.markdown("### Alcance del proyecto")
        st.write(datos_acta["alcance"])

        col_json, col_pdf = st.columns(2)

        with col_json:
            st.download_button(
                label="Descargar datos generados en JSON",
                data=json.dumps(datos_acta, ensure_ascii=False, indent=4),
                file_name="datos_acta_inicio.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button("📄 Generar PDF del acta"):
                try:
                    ruta_pdf = generar_pdf_acta_inicio(datos_acta)
                    st.session_state.ruta_pdf_acta_inicio_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as e:
                    st.error(f"No se pudo generar el PDF: {e}")

        if (
            st.session_state.get("ruta_pdf_acta_inicio_generado")
            and Path(st.session_state.ruta_pdf_acta_inicio_generado).exists()
        ):
            ruta_pdf = st.session_state.ruta_pdf_acta_inicio_generado
            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar PDF del acta",
                    data=f,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )