from pathlib import Path
from datetime import date, datetime
from html import escape
import json

import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.units import cm

from config.constants import OUTPUT_DIR, RUTA_LOGO_SENA
from services.json_service import guardar_datos_json as guardar_datos_json_base
from utils.nombres_archivo import safe_filename
from utils.validaciones import validar_campos_obligatorios
from utils.textos import limpiar_texto


VERSION_USO_INFRAESTRUCTURA = (
    "VERSION_MODULAR_USO_INFRAESTRUCTURA_FORMATO_INSTITUCIONAL_VALIDADO"
)
CARPETA_SALIDA = OUTPUT_DIR


def guardar_datos_json(
    datos: dict,
    ruta: str = "datos_uso_infraestructura.json",
    nombre_archivo: str | None = None,
) -> None:
    """Adaptador para conservar compatibilidad con el código validado original."""
    guardar_datos_json_base(datos, nombre_archivo or ruta)


def fecha_larga_espanol(fecha: date) -> str:
    meses = {
        1: "enero",
        2: "febrero",
        3: "marzo",
        4: "abril",
        5: "mayo",
        6: "junio",
        7: "julio",
        8: "agosto",
        9: "septiembre",
        10: "octubre",
        11: "noviembre",
        12: "diciembre",
    }
    return f"{fecha.day} días del mes de {meses[fecha.month]} de {fecha.year}"


def obtener_ruta_logo_sena() -> str | None:
    ruta_logo = Path(RUTA_LOGO_SENA)
    return str(ruta_logo) if ruta_logo.exists() else None


def _agregar_lista_numerada(
    historia: list,
    elementos: list[str],
    estilo: ParagraphStyle,
    espacio_inferior: int = 2,
) -> None:
    for indice, punto in enumerate(elementos, start=1):
        tabla_punto = Table(
            [
                [
                    Paragraph(f"{indice}.", estilo),
                    Paragraph(escape(punto), estilo),
                ]
            ],
            colWidths=[1.0 * cm, 14.0 * cm],
        )
        tabla_punto.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), espacio_inferior),
                ]
            )
        )
        historia.append(tabla_punto)


def generar_pdf_uso_infraestructura(datos: dict) -> str:
    Path(CARPETA_SALIDA).mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Uso_Infraestructura_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(Path(CARPETA_SALIDA) / nombre_archivo)

    page_width, page_height = letter

    def encabezado_logo(c, _doc) -> None:
        c.saveState()
        ruta_logo = obtener_ruta_logo_sena()

        if ruta_logo:
            try:
                imagen = ImageReader(ruta_logo)
                ancho_logo = 58
                alto_logo = 58
                x_logo = (page_width - ancho_logo) / 2
                y_logo = page_height - 92
                c.drawImage(
                    imagen,
                    x_logo,
                    y_logo,
                    width=ancho_logo,
                    height=alto_logo,
                    mask="auto",
                    preserveAspectRatio=True,
                )
            except Exception:
                c.setFillColor(colors.HexColor("#39A935"))
                c.setFont("Helvetica-Bold", 18)
                c.drawCentredString(page_width / 2, page_height - 60, "SENA")
                c.setFillColor(colors.black)
        else:
            c.setFillColor(colors.HexColor("#39A935"))
            c.setFont("Helvetica-Bold", 18)
            c.drawCentredString(page_width / 2, page_height - 60, "SENA")
            c.setFillColor(colors.black)

        c.restoreState()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        rightMargin=3.0 * cm,
        leftMargin=3.0 * cm,
        topMargin=3.4 * cm,
        bottomMargin=2.2 * cm,
    )

    estilo_titulo = ParagraphStyle(
        name="TituloUsoInfraestructura",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        alignment=TA_CENTER,
        spaceAfter=14,
    )

    estilo_normal = ParagraphStyle(
        name="NormalJustificadoUsoInfraestructura",
        fontName="Helvetica",
        fontSize=10.5,
        leading=14.5,
        alignment=TA_JUSTIFY,
        spaceAfter=8,
    )

    estilo_normal_sin_espacio = ParagraphStyle(
        name="NormalSinEspacioUsoInfraestructura",
        fontName="Helvetica",
        fontSize=10.5,
        leading=14.5,
        alignment=TA_JUSTIFY,
        spaceAfter=3,
    )

    estilo_negrita = ParagraphStyle(
        name="NegritaUsoInfraestructura",
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14.5,
        alignment=TA_LEFT,
        spaceBefore=8,
        spaceAfter=2,
    )

    estilo_tabla = ParagraphStyle(
        name="TablaUsoInfraestructura",
        fontName="Helvetica",
        fontSize=9.5,
        leading=11.5,
        alignment=TA_LEFT,
    )

    estilo_tabla_centro = ParagraphStyle(
        name="TablaCentroUsoInfraestructura",
        fontName="Helvetica",
        fontSize=9.5,
        leading=11.5,
        alignment=TA_CENTER,
    )

    historia = []

    ciudad = escape(str(datos.get("ciudad", "")).strip())
    fecha_documento = datos.get("fecha_documento")

    if isinstance(fecha_documento, date):
        fecha_texto = fecha_larga_espanol(fecha_documento)
        fecha_corta = fecha_documento.strftime("%d/%m/%Y")
        fecha_iso = fecha_documento.strftime("%Y-%m-%d")
    else:
        fecha_texto = escape(str(datos.get("fecha_texto", "")).strip())
        fecha_corta = escape(str(datos.get("fecha_corta", "")).strip())
        fecha_iso = escape(str(datos.get("fecha_iso", "")).strip())

    codigo_proyecto = escape(str(datos.get("codigo_proyecto", "")).strip())
    nombre_proyecto = escape(str(datos.get("nombre_proyecto", "")).strip())
    nombre_talento = escape(str(datos.get("nombre_talento", "")).strip())
    telefono_talento = escape(str(datos.get("telefono_talento", "")).strip())
    nombre_experto = escape(str(datos.get("nombre_experto", "")).strip())
    linea_experto = escape(str(datos.get("linea_experto", "")).strip())

    historia.append(
        Paragraph(
            "MANUAL DE PRESTAMO Y USO DE INFRAESTRUCTURA RED TECNOPARQUE"
            "<br/>COLOMBIA",
            estilo_titulo,
        )
    )

    historia.append(
        Paragraph(
            f"En la Ciudad de {ciudad} a los {fecha_texto}",
            estilo_normal,
        )
    )

    historia.append(
        Paragraph(
            f'Luego de aceptado el Proyecto '
            f'<b>“{codigo_proyecto} {nombre_proyecto}”.</b>',
            estilo_normal,
        )
    )

    historia.append(
        Paragraph(
            "Con el fin de brindar un mejor servicio y asegurar un uso apropiado de la "
            "infraestructura de la Red Tecnoparque del SENA, usted como Talento y Gestor "
            "Tecnoparque, deberá tener en cuenta los siguientes puntos y cumplirlos "
            "respectivamente: Se leen, socializan y comprenden.",
            estilo_normal,
        )
    )

    puntos_infraestructura = [
        (
            "Todo laboratorio de la Red Tecnoparque debe tener un manual de normas y "
            "comportamientos básicos sin importar el tipo de laboratorio que sea, como, "
            "por ejemplo: no ingresar con alimentos y bebidas al interior de los "
            "laboratorios, de ser necesario utilizar elementos de sonido mantener el "
            "volumen adecuado, obrar con honestidad, respeto, responsabilidad, mantener "
            "un tono de voz adecuado, etc., y las normas que cada gestor considere "
            "necesarias para una armonía de trabajo."
        ),
        (
            "Antes de comenzar a utilizar los laboratorios, los Gestores deben realizar "
            "una capacitación o transferencia básica para el cuidado y manejo de los "
            "laboratorios. En el interior de cada laboratorio debe haber un manual o "
            "instructivo de las condiciones de operación de los equipos con que el talento "
            "necesite trabajar, y que al mismo tiempo informe sobre la indumentaria "
            "adecuada para la operación de los equipos de ser necesaria (Utilizar los "
            "elementos de seguridad dispuestos para la operación de los equipos)."
        ),
        (
            "Cada gestor debe indicar el estado de los equipos y herramientas antes de "
            "que sean utilizados por los Talentos y verificar el estado de los mismos "
            "una vez el talento termine o entregue los equipos y herramientas."
        ),
        (
            "Cuando se solicite alguna herramienta de corte o se ingrese a un laboratorio "
            "para manipular equipos que requieren un manejo especial, se debe usar todos "
            "los equipos de protección personal según el caso, y se debe presentar sin "
            "excepción el carné de la EPS actualizado a la fecha de uso, el cual debe ser "
            "entregado al gestor encargado durante su trabajo con los equipos, este será "
            "devuelto al regresar las herramientas."
        ),
        (
            "Siempre que una herramienta o equipo le sea prestado, este deberá ser "
            "registrado por uno de los asesores en el formato común de préstamos, en "
            "ningún momento el Talento debe recibir o entregar un equipo o herramienta "
            "sin que sea registrado el préstamo o devolución de la misma, ya que se "
            "asumirá que aún no lo ha devuelto."
        ),
        (
            "Las herramientas se deben regresar al finalizar el día, esto quiere decir "
            "que toda herramienta y equipo debe ser entregado antes de las 5:30 p.m, si "
            "no es entregada en este horario debe ser reportada por el Talento al día "
            "siguiente antes de las 9:00 am, y dejar anotación en la minuta de la empresa "
            "de seguridad del Nodo. Si por razones extraordinarias, los equipos o "
            "herramientas salen del Nodo o del Centro para un acompañamiento a los "
            "Proyectos, cada Gestor encargado debe gestionar los Seguros correspondientes "
            "con el Almacén del Centro padrino SENA e informar al Cuentadante del equipo."
        ),
        (
            "El Talento se responsabilizará de lo que ocurra con los equipos y "
            "herramientas, al incumplir con la hora de entrega acordada de los equipos y "
            "herramientas, esta responsabilidad va desde cubrir gastos de reparación, "
            "hasta costos de reposición por pérdidas. Toda herramienta y equipo se entrega "
            "en perfecto estado de funcionamiento, si es entregada por el gestor, en mal "
            "estado debe ser reportado por el Talento de inmediato, de lo contrario tendrá "
            "que asumir los daños no reportados."
        ),
        (
            "Si un equipo o herramienta sufre daños por mal uso durante el tiempo de "
            "préstamo, este daño debe ser asumido por el Talento y serán suspendidos los "
            "servicios de Tecnoparque del SENA al proyecto mientras el daño no sea cubierto "
            "por el Talento."
        ),
        (
            "Al terminar de usar las herramientas, equipos, laboratorios o infraestructura "
            "en general, debe quedar limpio, ordenado y en buen estado, en las condiciones "
            "como se recibieron."
        ),
        (
            "Las herramientas no se deben retirar del piso, del ambiente o laboratorio al "
            "cual corresponden, de ser retirados deben tener la previa autorización del "
            "Cuentadante del bien y Gestor encargado."
        ),
        (
            "Si existe algún convenio, alianza, carta de intención o de cooperación con "
            "otro Nodo, Centro de formación u otra institución para uso de infraestructura "
            "compartida, se debe soportar por escrito el uso de los equipos y herramientas "
            "para beneficio de los usuarios o Talentos. Los tiempos y espacios de trabajo "
            "deben quedar pactados y oficialmente definidos con horarios y tiempos de uso. "
            "Así como los responsables por daños o deterioro de la infraestructura."
        ),
    ]

    _agregar_lista_numerada(
        historia,
        puntos_infraestructura,
        estilo_normal_sin_espacio,
    )

    historia.append(Spacer(1, 4))
    historia.append(Paragraph("En cuanto al software:", estilo_normal))

    puntos_software = [
        (
            "Todo equipo que registre un software ilegal será asociado con el Talento que "
            "tenía bajo préstamo ese equipo según la fecha y la hora de préstamo, lo que "
            "hace responsable al Talento de asumir todos los perjuicios y efectos legales "
            "que esta acción conlleve."
        ),
        (
            "Aquellos equipos que contengan archivos peligrosos y material escandaloso "
            "serán asociados con el Talento según la hora y fecha de creación, lo que dará "
            "la suspensión de los servicios de Tecnoparque del SENA para el proyecto."
        ),
        (
            "La Red Tecnoparque no se responsabiliza de los archivos y documentos guardados "
            "en los equipos de cómputo ni el uso que a ellos den otros Talentos."
        ),
        (
            "El cambiar contraseñas de equipos o crear nuevas sesiones dará suspensión de "
            "los servicios de Tecnoparque para el proyecto."
        ),
        (
            "Se prohíbe el ingreso a páginas cuyo fin sea la pornografía, la estafa o el ocio."
        ),
    ]

    _agregar_lista_numerada(
        historia,
        puntos_software,
        estilo_normal_sin_espacio,
    )

    historia.append(Paragraph("EXCEPCIÓN:", estilo_negrita))

    excepcion = (
        "Aquellos Talentos que requieran trabajar en el desarrollo de sus proyectos entre "
        "semana después de las 5:30 pm y los fines de semana, deberá contar con la "
        "respectiva autorización por escrito del subdirector del Centro, del Dinamizador "
        "del Nodo y Experto a cargo con el visto bueno del coordinador administrativo y "
        "previo aviso con 4 días de anterioridad."
    )

    tabla_excepcion = Table(
        [
            [
                Paragraph("1.", estilo_normal_sin_espacio),
                Paragraph(escape(excepcion), estilo_normal_sin_espacio),
            ]
        ],
        colWidths=[1.0 * cm, 14.0 * cm],
    )
    tabla_excepcion.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    historia.append(tabla_excepcion)

    historia.append(
        Paragraph(
            "Para constancia, se firman y comprometen al cumplimiento.",
            estilo_normal,
        )
    )

    tabla_firmas_data = [
        [
            Paragraph("<b>N</b>", estilo_tabla_centro),
            Paragraph("<b>Nombre del Talento</b>", estilo_tabla_centro),
            Paragraph("<b>Teléfono.</b>", estilo_tabla_centro),
            Paragraph("<b>Firma</b>", estilo_tabla_centro),
        ],
        [
            Paragraph("1", estilo_tabla_centro),
            Paragraph(nombre_talento, estilo_tabla),
            Paragraph(telefono_talento, estilo_tabla_centro),
            Paragraph("", estilo_tabla_centro),
        ],
        [
            Paragraph("<b>N</b>", estilo_tabla_centro),
            Paragraph("<b>Nombre del Experto encargado</b>", estilo_tabla_centro),
            Paragraph("<b>Línea</b>", estilo_tabla_centro),
            Paragraph("<b>Firma</b>", estilo_tabla_centro),
        ],
        [
            Paragraph("1", estilo_tabla_centro),
            Paragraph(nombre_experto, estilo_tabla),
            Paragraph(linea_experto, estilo_tabla_centro),
            Paragraph("", estilo_tabla_centro),
        ],
    ]

    tabla_firmas = Table(
        tabla_firmas_data,
        colWidths=[1.0 * cm, 6.0 * cm, 3.8 * cm, 4.5 * cm],
        rowHeights=[0.65 * cm, 1.15 * cm, 0.65 * cm, 1.25 * cm],
    )
    tabla_firmas.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (2, 0), (2, -1), "CENTER"),
                ("ALIGN", (3, 0), (3, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )

    historia.append(Spacer(1, 8))
    historia.append(tabla_firmas)

    doc.build(
        historia,
        onFirstPage=encabezado_logo,
        onLaterPages=encabezado_logo,
    )

    datos_json = {
        "tipo_documento": "Uso de infraestructura",
        "codigo_proyecto": datos.get("codigo_proyecto", ""),
        "nombre_proyecto": datos.get("nombre_proyecto", ""),
        "nombre_talento": datos.get("nombre_talento", ""),
        "telefono_talento": datos.get("telefono_talento", ""),
        "nombre_experto": datos.get("nombre_experto", ""),
        "linea_experto": datos.get("linea_experto", ""),
        "ciudad": datos.get("ciudad", ""),
        "fecha_corta": fecha_corta,
        "fecha_iso": fecha_iso,
        "version": VERSION_USO_INFRAESTRUCTURA,
        "ruta_pdf": ruta_pdf,
    }

    guardar_datos_json(datos_json, ruta="datos_uso_infraestructura.json")
    return ruta_pdf


def render_uso_infraestructura(
    modo_prueba: bool = True,
    modelo_openai: str = "",
) -> None:
    del modo_prueba, modelo_openai

    st.markdown("---")
    st.subheader("Formulario para Manual de Préstamo y Uso de Infraestructura")
    st.caption(VERSION_USO_INFRAESTRUCTURA)

    st.info(
        "Este documento no consume API de OpenAI. Se genera con base en el "
        "formato institucional de uso de infraestructura de la Red Tecnoparque Colombia."
    )

    if "datos_infraestructura_generada" not in st.session_state:
        st.session_state.datos_infraestructura_generada = None

    if "ruta_pdf_infraestructura_generado" not in st.session_state:
        st.session_state.ruta_pdf_infraestructura_generado = None

    with st.form("form_uso_infraestructura"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: P2024-143440-16602",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder=(
                    'Ejemplo: Diseño de un "Precipitómetro" de bajo costo para '
                    "determinación de valores reales de infiltración de suelos"
                ),
                height=100,
            )

            nombre_talento = st.text_input(
                "Nombre del talento",
                placeholder="Nombre completo del talento",
            )

            telefono_talento = st.text_input(
                "Teléfono del talento",
                placeholder="Ejemplo: 324 6428300",
            )

        with col_b:
            nombre_experto = st.text_input(
                "Nombre del experto",
                placeholder="Nombre completo del experto encargado",
            )

            linea_experto = st.text_input(
                "Línea del experto",
                placeholder="Ejemplo: Ingeniería y diseño",
            )

            ciudad = st.text_input(
                "Ciudad",
                value="Campoalegre",
            )

            fecha_documento = st.date_input(
                "Fecha del documento",
                value=date.today(),
            )

        generar_infraestructura = st.form_submit_button(
            "Generar documento de uso de infraestructura"
        )

    if generar_infraestructura:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Nombre del talento": nombre_talento,
            "Teléfono del talento": telefono_talento,
            "Nombre del experto": nombre_experto,
            "Línea del experto": linea_experto,
            "Ciudad": ciudad,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        datos_infraestructura = {
            "tipo_documento": "Uso de infraestructura",
            "codigo_proyecto": limpiar_texto(codigo_proyecto),
            "nombre_proyecto": limpiar_texto(nombre_proyecto),
            "nombre_talento": limpiar_texto(nombre_talento),
            "telefono_talento": limpiar_texto(telefono_talento),
            "nombre_experto": limpiar_texto(nombre_experto),
            "linea_experto": limpiar_texto(linea_experto),
            "ciudad": limpiar_texto(ciudad),
            "fecha_documento": fecha_documento,
            "fecha_corta": fecha_documento.strftime("%d/%m/%Y"),
            "fecha_iso": fecha_documento.strftime("%Y-%m-%d"),
            "version": VERSION_USO_INFRAESTRUCTURA,
        }

        st.session_state.datos_infraestructura_generada = datos_infraestructura
        st.session_state.ruta_pdf_infraestructura_generado = None

        st.success(
            "Información registrada correctamente. Ahora puedes revisar y generar el PDF."
        )

    if st.session_state.get("datos_infraestructura_generada"):
        datos = st.session_state.datos_infraestructura_generada

        st.markdown("## Resumen para validación")
        st.write("**Código del proyecto:**", datos["codigo_proyecto"])
        st.write("**Nombre del proyecto:**", datos["nombre_proyecto"])
        st.write("**Nombre del talento:**", datos["nombre_talento"])
        st.write("**Teléfono del talento:**", datos["telefono_talento"])
        st.write("**Nombre del experto:**", datos["nombre_experto"])
        st.write("**Línea del experto:**", datos["linea_experto"])
        st.write("**Ciudad:**", datos["ciudad"])
        st.write("**Fecha:**", datos["fecha_corta"])

        col_json, col_pdf = st.columns(2)

        with col_json:
            datos_json_descarga = dict(datos)
            fecha_json = datos_json_descarga.get("fecha_documento")

            if isinstance(fecha_json, date):
                datos_json_descarga["fecha_documento"] = fecha_json.strftime(
                    "%d/%m/%Y"
                )

            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(
                    datos_json_descarga,
                    ensure_ascii=False,
                    indent=4,
                ),
                file_name="datos_uso_infraestructura.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button(
                "📄 Generar PDF de uso de infraestructura",
                key="generar_pdf_uso_infraestructura",
            ):
                try:
                    ruta_pdf = generar_pdf_uso_infraestructura(datos)
                    st.session_state.ruta_pdf_infraestructura_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as error:
                    st.error(f"No se pudo generar el PDF: {error}")

        ruta_pdf = st.session_state.get("ruta_pdf_infraestructura_generado")

        if ruta_pdf and Path(ruta_pdf).exists():
            with open(ruta_pdf, "rb") as archivo_pdf:
                st.download_button(
                    label="⬇️ Descargar PDF de uso de infraestructura",
                    data=archivo_pdf,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )
