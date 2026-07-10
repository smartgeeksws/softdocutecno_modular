from pathlib import Path
from datetime import date, time, datetime, timedelta
import json

import streamlit as st

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.pdfmetrics import stringWidth
except ImportError:
    canvas = None

from config.constants import (
    OUTPUT_DIR,
    RUTA_LOGO_SENA,
    FORMATO_ACTA_CIERRE,
)
from services.json_service import guardar_datos_json
from services.openai_service import generar_json_openai
from utils.nombres_archivo import safe_filename
from utils.formatos import formato_moneda_colombiana
from utils.textos import limpiar_texto
from utils.validaciones import validar_campos_obligatorios


VERSION_ACTA_CIERRE = "VERSION_MODULAR_ACTA_CIERRE_FORMATO_VALIDADO_PROYECTO_INICIAL"


# =====================================================
# FUNCIONES BASE DEL FORMATO INSTITUCIONAL VALIDADO
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
            c.drawImage(
                img,
                logo_x,
                logo_y,
                width=logo_w,
                height=logo_h,
                mask="auto",
            )
            return
        except Exception:
            pass

    c.setFillColor(colors.HexColor("#69B342"))
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(page_width / 2, top_y - 18, "SENA")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(page_width / 2, top_y - 36, "▲")
    c.setFillColor(colors.black)


def calcular_hora_fin(fecha_acta: date, hora_inicio: time) -> time:
    inicio_dt = datetime.combine(fecha_acta, hora_inicio)
    fin_dt = inicio_dt + timedelta(minutes=30)
    return fin_dt.time()


def generar_objetivo_cierre(codigo_proyecto: str, nombre_proyecto: str) -> str:
    return (
        f"Dar por finalizada la ejecución del proyecto {codigo_proyecto} - {nombre_proyecto}, "
        "revisando los objetivos alcanzados, entregables, desempeño, y obtener la aprobación formal "
        "del cierre del proyecto."
    )


# =====================================================
# EVIDENCIAS DEL ACTA DE CIERRE
# =====================================================

def generar_evidencias_cierre_modo_prueba(nombre_proyecto: str, evidencias_producto: str) -> dict:
    texto_lower = evidencias_producto.lower()

    if "susceptible de inscribir un nuevo proyecto" in texto_lower or "inscribir un nuevo proyecto" in texto_lower:
        conclusion_adicional = (
            "Adicionalmente, se identifica que el proyecto es susceptible de inscribir un nuevo proyecto "
            "o una nueva idea de base tecnológica, con el fin de continuar su fortalecimiento, validación, "
            "escalamiento o desarrollo de nuevas funcionalidades."
        )
    else:
        conclusion_adicional = (
            "Se evaluará la posibilidad de inscribir un nuevo prototipo, idea o proyecto de base tecnológica, "
            "de acuerdo con los resultados obtenidos, las oportunidades de mejora identificadas y el potencial "
            "de continuidad técnica del desarrollo alcanzado."
        )

    entregables = [
        item.strip(" -•0123456789.)")
        for item in evidencias_producto.replace(";", "\n").split("\n")
        if item.strip()
    ]

    if entregables:
        entregables_texto = " ".join(
            [
                f"Se evidencia el desarrollo y entrega de {entregable}, como componente técnico asociado al prototipo, producto o resultado obtenido durante la ejecución del proyecto."
                for entregable in entregables
            ]
        )
    else:
        entregables_texto = (
            "Se registran evidencias asociadas al desarrollo del prototipo, incluyendo productos, componentes, "
            "documentos técnicos, validaciones, diseños, pruebas o implementaciones generadas durante la ejecución del proyecto."
        )

    return {
        "evidencias_normatividad": (
            f"De acuerdo con la naturaleza técnica del proyecto {nombre_proyecto}, se identifican referentes normativos "
            "aplicables para orientar la validación, documentación y cierre técnico del prototipo desarrollado. "
            "Según el tipo de solución, pueden considerarse Normas Técnicas Colombianas NTC relacionadas con gestión "
            "de calidad, documentación técnica, seguridad de producto, requisitos de operación, trazabilidad, "
            "medición, validación funcional y buenas prácticas de desarrollo tecnológico. De manera complementaria, "
            "pueden tomarse como referencia normas internacionales ISO aplicables a sistemas de gestión de calidad, "
            "diseño de productos, procesos de ensayo, documentación de resultados, seguridad de operación, "
            "interoperabilidad tecnológica, pruebas funcionales y validación de componentes. Estas referencias permiten "
            "establecer criterios mínimos para verificar que el prototipo, producto o solución desarrollada cuente con "
            "evidencias técnicas suficientes, trazabilidad documental y condiciones adecuadas para futuras fases de "
            "fortalecimiento, transferencia, mejora o escalamiento."
        ),
        "evidencias_modelo_negocio": (
            "Se adjunta el Modelo Canvas aplicado al Proyecto de Base Tecnológica, como herramienta de análisis "
            "para la identificación de la propuesta de valor, segmentos de cliente, canales, recursos clave, "
            "actividades clave, aliados estratégicos, estructura de costos y fuentes de ingreso."
        ),
        "evidencias_pruebas_documentadas": (
            "Se adjunta el Informe Técnico Final, en el cual se documenta la metodología desarrollada, los procesos "
            "de validación, las pruebas realizadas, los resultados obtenidos y la implementación técnica del proyecto."
        ),
        "evidencias_prototipo": entregables_texto,
        "conclusion_adicional": conclusion_adicional,
    }


def generar_evidencias_cierre_con_chatgpt(
    nombre_proyecto: str,
    evidencias_producto: str,
    modelo: str = "gpt-4.1-mini",
) -> dict:
    instrucciones = """
Eres un experto en cierre técnico de proyectos de base tecnológica de la Red Tecnoparque SENA.

Debes generar evidencias para un Acta de Cierre institucional.
Responde únicamente en JSON válido.
No uses markdown.
No inventes nombres de personas, fechas, códigos ni entidades.
Las normas deben ser genéricas y aplicables según el tipo de prototipo, citando NTC o normas internacionales cuando corresponda.
"""

    entrada = f"""
Proyecto:
{nombre_proyecto}

Texto genérico ingresado por el usuario en Evidencias del Producto:
{evidencias_producto}

Genera un JSON con esta estructura exacta:

{{
  "evidencias_normatividad": "Redacción técnica sobre NTC o normas internacionales aplicables al tipo de prototipo.",
  "evidencias_modelo_negocio": "Frase estándar indicando que se adjunta Modelo Canvas aplicado al PBT.",
  "evidencias_pruebas_documentadas": "Frase indicando que se adjunta Informe Técnico Final con metodología, validación e implementación.",
  "evidencias_prototipo": "Lista redactada de entregables específicos extraídos del texto de evidencias del producto. Debe entenderse como evidencias de prototipo y entregables desarrollados.",
  "conclusion_adicional": "Conclusión adicional solo si el texto menciona que el proyecto es susceptible de inscribir un nuevo proyecto. Si no aplica, dejar vacío."
}}
"""

    try:
        datos = generar_json_openai(
            instrucciones=instrucciones,
            entrada=entrada,
            modelo=modelo,
            temperature=0.25,
        )

        campos = [
            "evidencias_normatividad",
            "evidencias_modelo_negocio",
            "evidencias_pruebas_documentadas",
            "evidencias_prototipo",
            "conclusion_adicional",
        ]

        for campo in campos:
            if campo not in datos or not isinstance(datos[campo], str):
                datos[campo] = ""

        datos_base = generar_evidencias_cierre_modo_prueba(nombre_proyecto, evidencias_producto)

        if not datos.get("evidencias_prototipo", "").strip():
            datos["evidencias_prototipo"] = datos_base["evidencias_prototipo"]

        if not datos.get("evidencias_normatividad", "").strip() or len(datos["evidencias_normatividad"]) < 250:
            datos["evidencias_normatividad"] = datos_base["evidencias_normatividad"]

        texto_lower = evidencias_producto.lower()

        if "susceptible de inscribir un nuevo proyecto" in texto_lower or "inscribir un nuevo proyecto" in texto_lower:
            if not datos["conclusion_adicional"]:
                datos["conclusion_adicional"] = (
                    "Adicionalmente, se identifica que el proyecto es susceptible de inscribir un nuevo proyecto "
                    "o una nueva idea de base tecnológica, con el fin de continuar su fortalecimiento, validación, "
                    "escalamiento o desarrollo de nuevas funcionalidades."
                )
        else:
            datos["conclusion_adicional"] = (
                "Se evaluará la posibilidad de inscribir un nuevo prototipo, idea o proyecto de base tecnológica, "
                "de acuerdo con los resultados obtenidos, las oportunidades de mejora identificadas y el potencial "
                "de continuidad técnica del desarrollo alcanzado."
            )

        return datos

    except Exception:
        return generar_evidencias_cierre_modo_prueba(nombre_proyecto, evidencias_producto)


# =====================================================
# PDF ACTA DE CIERRE - FORMATO VALIDADO
# =====================================================

def generar_pdf_acta_cierre(datos: dict) -> str:
    if canvas is None:
        raise ImportError("No está instalada reportlab. Instálala con: pip install reportlab")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Acta_Cierre_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(OUTPUT_DIR / nombre_archivo)

    c = canvas.Canvas(ruta_pdf, pagesize=letter)
    page_width, _page_height = letter

    x0 = 24
    table_w = page_width - 48
    logo_top_y = 785
    y_top_content = 705
    y_safe_bottom = 58

    FONT_TITLE = 10.5
    FONT_SECTION = 10.5
    FONT_BODY = 7.1
    FONT_SMALL = 6.7
    FONT_TINY = 6.3

    def iniciar_pagina() -> float:
        draw_logo(c, page_width, logo_top_y)
        return y_top_content

    def cerrar_pagina() -> None:
        c.setFillColor(colors.grey)
        c.setFont("Helvetica", 10)
        c.drawCentredString(page_width / 2, 20, FORMATO_ACTA_CIERRE)
        c.setFillColor(colors.black)
        c.showPage()

    def asegurar_espacio(y_actual: float, alto_requerido: float) -> float:
        if y_actual - alto_requerido < y_safe_bottom:
            cerrar_pagina()
            return iniciar_pagina()
        return y_actual

    def alto_texto_local(
        texto: str,
        ancho: float,
        base: float = 24,
        font_size: float = FONT_BODY,
        label: str | None = None,
        max_h: float = 80,
    ) -> float:
        padding = 10
        usable_w = max(ancho - padding, 20)
        texto_total = f"{label or ''} {texto or ''}".strip()
        lineas = wrap_text(texto_total, "Helvetica", font_size, usable_w)
        h = max(base, len(lineas) * (font_size + 2.2) + 12)
        return min(h, max_h)

    y = iniciar_pagina()

    # Título
    h = 22
    draw_cell(
        c,
        x0,
        y - h,
        table_w,
        h,
        f"ACTA No. 03 del proyecto No {datos.get('codigo_proyecto', '')}",
        font="Helvetica-Bold",
        size=FONT_TITLE,
        center=True,
    )
    y -= h

    # Nombre del comité o reunión
    nombre_comite = f"Acta de cierre del proyecto {datos.get('codigo_proyecto', '')} - {datos.get('nombre_proyecto', '')}"
    h = alto_texto_local(
        nombre_comite,
        table_w,
        base=38,
        font_size=FONT_BODY,
        label="NOMBRE DEL COMITÉ O DE LA REUNIÓN:",
        max_h=58,
    )
    draw_cell(
        c,
        x0,
        y - h,
        table_w,
        h,
        nombre_comite,
        label="NOMBRE DEL COMITÉ O DE LA REUNIÓN:",
        size=FONT_BODY,
    )
    y -= h

    # Ciudad, fecha, hora
    h = 38
    w1 = table_w * 0.58
    w2 = table_w * 0.21
    w3 = table_w * 0.21
    draw_cell(
        c,
        x0,
        y - h,
        w1,
        h,
        f"Campoalegre (Huila) - {datos.get('fecha_iso', '')}",
        label="CIUDAD Y FECHA:",
        size=FONT_BODY,
    )
    draw_cell(c, x0 + w1, y - h, w2, h, datos.get("hora_inicio", ""), label="HORA INICIO:", size=FONT_BODY)
    draw_cell(c, x0 + w1 + w2, y - h, w3, h, datos.get("hora_fin", ""), label="HORA FIN:", size=FONT_BODY)
    y -= h

    # Lugar y centro
    h = 46
    w_lugar = table_w * 0.55
    w_dir = table_w * 0.45
    draw_cell(
        c,
        x0,
        y - h,
        w_lugar,
        h,
        "Tecnoparque Angostura - Campoalegre, Huila",
        label="LUGAR Y/O ENLACE:",
        size=FONT_BODY,
    )
    draw_cell(
        c,
        x0 + w_lugar,
        y - h,
        w_dir,
        h,
        "Centro de Formación Agroindustrial / SENA Regional Huila",
        label="DIRECCIÓN / REGIONAL / CENTRO:",
        size=FONT_BODY,
    )
    y -= h

    # Objetivo de la reunión
    objetivo_reunion = datos.get(
        "objetivo_cierre",
        f"Dar por finalizada la ejecución del proyecto {datos.get('codigo_proyecto', '')} - {datos.get('nombre_proyecto', '')}, revisando los objetivos alcanzados, entregables, desempeño, y obtener la aprobación formal del cierre del proyecto.",
    )
    h = alto_texto_local(
        objetivo_reunion,
        table_w,
        base=40,
        font_size=FONT_BODY,
        label="OBJETIVO(S) DE LA REUNIÓN:",
        max_h=56,
    )
    draw_cell(c, x0, y - h, table_w, h, objetivo_reunion, label="OBJETIVO(S) DE LA REUNIÓN:", size=FONT_BODY)
    y -= h

    # Desarrollo
    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "DESARROLLO DE LA REUNIÓN", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    # Código y nombre
    codigo_nombre = f"{datos.get('codigo_proyecto', '')} - {datos.get('nombre_proyecto', '')}"
    h = alto_texto_local(codigo_nombre, table_w, base=26, font_size=FONT_BODY, label="Código y nombre del Proyecto:", max_h=46)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, codigo_nombre, label="Código y nombre del Proyecto:", size=FONT_BODY)
    y -= h

    # TRL y aporte
    h = 32
    w_trl = table_w * 0.35
    w_aporte = table_w * 0.65
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, w_trl, h, datos.get("trl_obtenido", ""), label="TRL OBTENIDO:", size=FONT_BODY)
    draw_cell(
        c,
        x0 + w_trl,
        y - h,
        w_aporte,
        h,
        formato_moneda_colombiana(datos.get("aporte_tecnoparque", 0)),
        label="APORTE ESTIMADO DE TECNOPARQUE:",
        size=FONT_BODY,
    )
    y -= h

    # Objetivo de cierre
    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "OBJETIVO DEL CIERRE", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    objetivo_cierre = datos.get("objetivo_cierre", "")
    h = alto_texto_local(objetivo_cierre, table_w, base=42, font_size=FONT_BODY, max_h=72)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, objetivo_cierre, size=FONT_BODY)
    y -= h

    # Objetivos iniciales y cumplimiento
    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(
        c,
        x0,
        y - h,
        table_w,
        h,
        "OBJETIVOS INICIALES DEL PROYECTO Y CUMPLIMIENTO",
        font="Helvetica-Bold",
        size=FONT_SECTION,
        center=True,
    )
    y -= h

    objetivos_iniciales = datos.get("objetivos_iniciales", [])
    if not objetivos_iniciales:
        objetivos_iniciales = ["No se registraron objetivos iniciales."]

    col_num = 38
    col_obj = table_w - 130
    col_cumple = 92

    h = 24
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, col_num, h, "No.", font="Helvetica-Bold", size=FONT_SMALL, center=True)
    draw_cell(c, x0 + col_num, y - h, col_obj, h, "Objetivo inicial", font="Helvetica-Bold", size=FONT_SMALL, center=True)
    draw_cell(c, x0 + col_num + col_obj, y - h, col_cumple, h, "Cumplió", font="Helvetica-Bold", size=FONT_SMALL, center=True)
    y -= h

    for idx, obj in enumerate(objetivos_iniciales, start=1):
        h = alto_texto_local(obj, col_obj, base=28, font_size=FONT_SMALL, max_h=58)
        y = asegurar_espacio(y, h)

        draw_cell(c, x0, y - h, col_num, h, str(idx), font="Helvetica-Bold", size=FONT_SMALL, center=True)
        draw_cell(c, x0 + col_num, y - h, col_obj, h, obj, size=FONT_SMALL)
        draw_cell(c, x0 + col_num + col_obj, y - h, col_cumple, h, "SI", font="Helvetica-Bold", size=FONT_SMALL, center=True)

        y -= h

    # Evidencias
    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(
        c,
        x0,
        y - h,
        table_w,
        h,
        "EVIDENCIAS DEL PROYECTO",
        font="Helvetica-Bold",
        size=FONT_SECTION,
        center=True,
    )
    y -= h

    contenido = datos.get("evidencias_generadas", {})

    evidencias = [
        ("Evidencias de Normatividad:", contenido.get("evidencias_normatividad", "")),
        ("Evidencias de Modelo de Negocio:", contenido.get("evidencias_modelo_negocio", "")),
        ("Evidencias de Pruebas Documentadas:", contenido.get("evidencias_pruebas_documentadas", "")),
        ("Evidencias de Prototipo:", contenido.get("evidencias_prototipo", "")),
    ]

    for titulo, texto in evidencias:
        if not str(texto).strip():
            texto = "No se registró información específica para esta evidencia."

        h = alto_texto_local(
            texto,
            table_w,
            base=44,
            font_size=FONT_TINY,
            label=titulo,
            max_h=95,
        )

        y = asegurar_espacio(y, h)
        draw_cell(c, x0, y - h, table_w, h, texto, label=titulo, size=FONT_TINY)
        y -= h

    # Conclusiones
    h = 22
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, "CONCLUSIONES", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= h

    conclusiones = "Se cumplieron a cabalidad todos los objetivos del proyecto."
    if contenido.get("conclusion_adicional"):
        conclusiones += " " + contenido.get("conclusion_adicional")

    h = alto_texto_local(conclusiones, table_w, base=34, font_size=FONT_BODY, max_h=64)
    y = asegurar_espacio(y, h)
    draw_cell(c, x0, y - h, table_w, h, conclusiones, size=FONT_BODY)
    y -= h

    # Asistentes y aprobación
    asistentes_alto = 22 + 34 + 36 + 36
    y = asegurar_espacio(y, asistentes_alto)

    draw_cell(c, x0, y - 22, table_w, 22, "ASISTENTES Y APROBACIÓN DE DECISIONES", font="Helvetica-Bold", size=FONT_SECTION, center=True)
    y -= 22

    col_w = [table_w * 0.22, table_w * 0.22, table_w * 0.16, table_w * 0.18, table_w * 0.22]
    headers = ["NOMBRE", "DEPENDENCIA / EMPRESA", "APRUEBA (SI/NO)", "OBSERVACIÓN", "FIRMA"]
    x = x0

    for w, header in zip(col_w, headers):
        draw_cell(c, x, y - 34, w, 34, header, font="Helvetica-Bold", size=FONT_SMALL, center=True)
        x += w

    y -= 34

    filas = [
        [datos.get("nombre_talento", ""), "Emprendedor/Empresario", "SI", "", ""],
        [datos.get("nombre_experto", ""), "SENA", "SI", "", ""],
    ]

    for fila in filas:
        x = x0
        for w, value in zip(col_w, fila):
            center = value in ["Emprendedor/Empresario", "SENA", "SI", ""]
            draw_cell(c, x, y - 36, w, 36, value, size=FONT_SMALL, center=center)
            x += w

        y -= 36

    cerrar_pagina()
    c.save()

    datos_json = dict(datos)
    datos_json["ruta_pdf"] = ruta_pdf
    guardar_datos_json(datos_json, nombre_archivo="datos_acta_cierre.json")

    return ruta_pdf


# =====================================================
# INTERFAZ STREAMLIT
# =====================================================

def render_acta_cierre(modo_prueba: bool = True, modelo_openai: str = "gpt-4.1-mini") -> None:
    st.markdown("---")
    st.subheader("Formulario para Acta de Cierre")
    st.caption(VERSION_ACTA_CIERRE)

    st.info(
        "Este módulo genera el Acta de Cierre del proyecto con el formato institucional validado "
        "en la versión inicial del software."
    )

    with st.form("form_acta_cierre"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2025-143440-00001",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Nombre oficial del proyecto",
                height=90,
            )

            fecha_acta = st.date_input(
                "Fecha del acta",
                value=date.today(),
            )

            hora_inicio = st.time_input(
                "Hora de inicio",
                value=time(8, 0),
            )

        with col_b:
            nombre_talento = st.text_input(
                "Nombre del talento",
                placeholder="Nombre completo del talento",
            )

            nombre_experto = st.text_input(
                "Nombre del experto",
                placeholder="Nombre completo del experto",
            )

            trl_obtenido = st.selectbox(
                "TRL obtenido",
                options=["TRL 6", "TRL 7", "TRL 8"],
            )

            aporte_tecnoparque = st.number_input(
                "Aporte estimado de Tecnoparque",
                min_value=0,
                value=0,
                step=10000,
            )

        objetivos_iniciales = st.text_area(
            "Objetivos iniciales del proyecto",
            placeholder="Copia aquí los objetivos iniciales del proyecto. Cada objetivo puede ir en una línea diferente.",
            height=140,
        )

        evidencias_producto = st.text_area(
            "Evidencias del Producto",
            placeholder=(
                "Describe de forma general los productos, prototipos, documentos, pruebas, diseños, "
                "implementaciones o entregables generados. Si aplica, menciona si el proyecto es susceptible "
                "de inscribir un nuevo proyecto."
            ),
            height=180,
        )

        generar_acta_cierre = st.form_submit_button("Generar Acta de Cierre")

    if generar_acta_cierre:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Nombre del talento": nombre_talento,
            "Nombre del experto": nombre_experto,
            "Objetivos iniciales del proyecto": objetivos_iniciales,
            "Evidencias del Producto": evidencias_producto,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        hora_fin = calcular_hora_fin(fecha_acta, hora_inicio)
        objetivo_cierre = generar_objetivo_cierre(codigo_proyecto, nombre_proyecto)

        objetivos_lista = [
            limpiar_texto(obj.strip())
            for obj in objetivos_iniciales.replace("•", "\n").replace(";", "\n").split("\n")
            if obj.strip()
        ]

        with st.spinner("Generando evidencias del acta de cierre..."):
            try:
                if modo_prueba:
                    evidencias_generadas = generar_evidencias_cierre_modo_prueba(
                        nombre_proyecto,
                        evidencias_producto,
                    )
                else:
                    evidencias_generadas = generar_evidencias_cierre_con_chatgpt(
                        nombre_proyecto,
                        evidencias_producto,
                        modelo_openai,
                    )
            except Exception as e:
                st.warning(f"No se pudo usar IA. Se generará una versión base. Detalle: {e}")
                evidencias_generadas = generar_evidencias_cierre_modo_prueba(
                    nombre_proyecto,
                    evidencias_producto,
                )

        datos_acta_cierre = {
            "tipo_documento": "Acta de Cierre",
            "titulo_acta": f"ACTA No. 03 del proyecto No {codigo_proyecto}",
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "fecha_acta": fecha_acta.strftime("%d/%m/%Y"),
            "fecha_iso": fecha_acta.strftime("%Y-%m-%d"),
            "hora_inicio": hora_inicio.strftime("%H:%M"),
            "hora_fin": hora_fin.strftime("%H:%M"),
            "nombre_talento": limpiar_texto(nombre_talento),
            "nombre_experto": limpiar_texto(nombre_experto),
            "trl_obtenido": trl_obtenido,
            "aporte_tecnoparque": aporte_tecnoparque,
            "evidencias_producto": limpiar_texto(evidencias_producto),
            "objetivo_cierre": objetivo_cierre,
            "objetivos_iniciales": objetivos_lista,
            "evidencias_generadas": evidencias_generadas,
            "modo_generacion": "Prueba local" if modo_prueba else "ChatGPT API",
            "version": VERSION_ACTA_CIERRE,
        }

        st.session_state.datos_acta_cierre_generada = datos_acta_cierre
        st.session_state.ruta_pdf_acta_cierre_generado = None

        st.success("Acta de cierre generada correctamente. Ahora puedes revisar y generar el PDF.")

    if st.session_state.get("datos_acta_cierre_generada"):
        datos = st.session_state.datos_acta_cierre_generada
        evidencias = datos["evidencias_generadas"]

        st.markdown("## Resumen para validación")
        st.write("**Título:**", datos["titulo_acta"])
        st.write("**Proyecto:**", datos["nombre_proyecto"])
        st.write("**Código:**", datos["codigo_proyecto"])
        st.write("**Fecha:**", datos["fecha_acta"])
        st.write("**Hora inicio:**", datos["hora_inicio"])
        st.write("**Hora fin:**", datos["hora_fin"])
        st.write("**Talento:**", datos["nombre_talento"])
        st.write("**Experto:**", datos["nombre_experto"])
        st.write("**TRL obtenido:**", datos["trl_obtenido"])
        st.write("**Aporte Tecnoparque:**", formato_moneda_colombiana(datos["aporte_tecnoparque"]))

        st.markdown("### Objetivo de cierre")
        st.write(datos.get("objetivo_cierre", ""))

        st.markdown("### Objetivos iniciales del proyecto")
        for i, objetivo in enumerate(datos.get("objetivos_iniciales", []), start=1):
            st.write(f"{i}. SI Cumplió — {objetivo}")

        st.markdown("### Evidencias generadas")
        st.write("**Normatividad:**", evidencias.get("evidencias_normatividad", ""))
        st.write("**Modelo de negocio:**", evidencias.get("evidencias_modelo_negocio", ""))
        st.write("**Pruebas documentadas:**", evidencias.get("evidencias_pruebas_documentadas", ""))
        st.write("**Prototipo:**", evidencias.get("evidencias_prototipo", ""))

        if evidencias.get("conclusion_adicional"):
            st.write("**Conclusión adicional:**", evidencias.get("conclusion_adicional"))

        col_json, col_pdf = st.columns(2)

        with col_json:
            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(datos, ensure_ascii=False, indent=4),
                file_name="datos_acta_cierre.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button("📄 Generar PDF del Acta de Cierre"):
                try:
                    ruta_pdf = generar_pdf_acta_cierre(datos)
                    st.session_state.ruta_pdf_acta_cierre_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as e:
                    st.error(f"No se pudo generar el PDF: {e}")

        if (
            st.session_state.get("ruta_pdf_acta_cierre_generado")
            and Path(st.session_state.ruta_pdf_acta_cierre_generado).exists()
        ):
            ruta_pdf = st.session_state.ruta_pdf_acta_cierre_generado

            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar PDF del Acta de Cierre",
                    data=f,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )