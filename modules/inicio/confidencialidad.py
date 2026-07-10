from pathlib import Path
from datetime import date, datetime
import json
import tempfile

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
    Image,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.units import cm

from config.constants import (
    OUTPUT_DIR,
    RUTA_LOGO_SENA,
    CARPETA_FIRMAS,
    FORMATO_CONFIDENCIALIDAD,
)
from services.json_service import guardar_datos_json as guardar_datos_json_base
from utils.nombres_archivo import safe_filename
from utils.validaciones import validar_campos_obligatorios
from utils.fechas import fecha_larga_espanol_con_del


VERSION_CONFIDENCIALIDAD = "VERSION_MODULAR_CONFIDENCIALIDAD_COMPROMISO_FORMATO_VALIDADO"
CARPETA_SALIDA = OUTPUT_DIR


def guardar_datos_json(datos: dict, ruta: str = "datos_confidencialidad_compromiso.json", nombre_archivo: str | None = None) -> None:
    """Adaptador para conservar compatibilidad con el código validado del proyecto inicial."""
    guardar_datos_json_base(datos, nombre_archivo or ruta)


def ruta_firma(nombre_archivo: str) -> str:
    return str(Path(CARPETA_FIRMAS) / nombre_archivo)


def guardar_archivo_subido(uploaded_file, prefijo: str) -> str | None:
    if uploaded_file is None:
        return None

    sufijo = Path(uploaded_file.name).suffix.lower()
    if sufijo not in [".png", ".jpg", ".jpeg"]:
        raise ValueError("La firma debe estar en formato PNG, JPG o JPEG.")

    carpeta_tmp = Path(tempfile.gettempdir()) / "firmas_tecnoparque"
    carpeta_tmp.mkdir(parents=True, exist_ok=True)

    ruta = carpeta_tmp / f"{prefijo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{sufijo}"
    ruta.write_bytes(uploaded_file.getbuffer())
    return str(ruta)

# =====================================================
# PDF - CONFIDENCIALIDAD Y COMPROMISO
# =====================================================

def obtener_ruta_logo_sena() -> str | None:
    posibles_rutas = [
        Path(RUTA_LOGO_SENA),
        Path(CARPETA_FIRMAS) / "logo_sena.png",
        Path(CARPETA_FIRMAS) / "logo_sena.jpg",
        Path(CARPETA_FIRMAS) / "logo_sena.jpeg",
    ]

    for ruta in posibles_rutas:
        if ruta.exists():
            return str(ruta)

    return None


def crear_parrafo(texto: str, estilo: ParagraphStyle):
    return Paragraph(str(texto).replace("\n", "<br/>"), estilo)


def firma_img(path: str | None, width: float = 4.0 * cm, height: float = 1.05 * cm):
    if path and Path(path).exists():
        try:
            return Image(path, width=width, height=height)
        except Exception:
            return Paragraph("", ParagraphStyle(name="vacio", fontSize=8))
    return Paragraph("", ParagraphStyle(name="vacio", fontSize=8))


def generar_pdf_confidencialidad(datos: dict) -> str:
    Path(CARPETA_SALIDA).mkdir(parents=True, exist_ok=True)

    nombre_archivo = (
        f"Confidencialidad_Compromiso_"
        f"{safe_filename(datos.get('codigo_proyecto', 'proyecto'))}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    )
    ruta_pdf = str(Path(CARPETA_SALIDA) / nombre_archivo)

    fecha_documento = datos["fecha_documento"]
    ciudad = datos["ciudad"]
    nombre_proyecto = datos["nombre_proyecto"]
    codigo_proyecto = datos["codigo_proyecto"]
    nombre_talento = datos["nombre_talento"]
    cedula_talento = datos["cedula_talento"]
    ciudad_expedicion = datos["ciudad_expedicion"]
    ruta_firma_talento = datos.get("ruta_firma_talento")

    page_width, page_height = letter

    def encabezado_pie(c, doc):
        c.saveState()

        # Logo SENA centrado en encabezado
        ruta_logo = obtener_ruta_logo_sena()
        if ruta_logo:
            try:
                logo = ImageReader(ruta_logo)
                logo_w = 2.2 * cm
                logo_h = 2.2 * cm
                c.drawImage(
                    logo,
                    (page_width - logo_w) / 2,
                    page_height - 2.6 * cm,
                    width=logo_w,
                    height=logo_h,
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                c.setFillColor(colors.HexColor("#39a935"))
                c.setFont("Helvetica-Bold", 16)
                c.drawCentredString(page_width / 2, page_height - 1.7 * cm, "SENA")
                c.setFillColor(colors.black)
        else:
            c.setFillColor(colors.HexColor("#39a935"))
            c.setFont("Helvetica-Bold", 16)
            c.drawCentredString(page_width / 2, page_height - 1.7 * cm, "SENA")
            c.setFillColor(colors.black)

        # Código únicamente en pie de página centrado
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.grey)
        c.drawCentredString(page_width / 2, 1.1 * cm, FORMATO_CONFIDENCIALIDAD)
        c.setFillColor(colors.black)

        c.restoreState()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        rightMargin=2.35 * cm,
        leftMargin=2.35 * cm,
        topMargin=3.4 * cm,
        bottomMargin=2.2 * cm,
    )

    estilo_normal = ParagraphStyle(
        name="NormalConfidencialidad",
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        alignment=TA_JUSTIFY,
        spaceAfter=8,
    )

    estilo_titulo = ParagraphStyle(
        name="TituloConfidencialidad",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        alignment=TA_CENTER,
        spaceAfter=12,
    )

    estilo_subtitulo = ParagraphStyle(
        name="SubtituloConfidencialidad",
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        alignment=TA_CENTER,
        spaceBefore=8,
        spaceAfter=8,
    )

    estilo_tabla = ParagraphStyle(
        name="TablaConfidencialidad",
        fontName="Helvetica",
        fontSize=8.8,
        leading=10.5,
        alignment=TA_LEFT,
    )

    estilo_tabla_negrita = ParagraphStyle(
        name="TablaConfidencialidadNegrita",
        fontName="Helvetica-Bold",
        fontSize=8.8,
        leading=10.5,
        alignment=TA_CENTER,
    )

    historia = []

    historia.append(
        crear_parrafo(
            "PROCESO GESTIÓN DE INNOVACIÓN Y COMPETITIVIDAD<br/>"
            "FORMATO DE CONFIDENCIALIDAD Y COMPROMISOS RED TECNOPARQUE",
            estilo_titulo,
        )
    )

    texto_intro = (
        f"En la ciudad de {ciudad} a los {fecha_larga_espanol_con_del(fecha_documento)}, "
        "se celebra la presente Acta de Confidencialidad y Compromisos entre la Red Tecnoparque "
        "SENA Nodo Angostura representado por los firmantes abajo en este documento, y por otra "
        f"parte el Titular del Proyecto {nombre_talento}, identificado con cedula N° {cedula_talento} "
        f"de {ciudad_expedicion}, quien en adelante se denominará Titular y talento interlocutor del "
        f"proyecto {codigo_proyecto} - {nombre_proyecto} previas las siguientes consideraciones:"
    )
    historia.append(crear_parrafo(texto_intro, estilo_normal))

    historia.append(crear_parrafo("CONSIDERACIONES", estilo_subtitulo))
    historia.append(
        crear_parrafo(
            "Para dar inicio a la etapa de desarrollo del proyecto de base tecnológica enunciado "
            "anteriormente, es necesario establecer acuerdos entre las partes en las siguientes "
            "materias: i) el manejo de la información confidencial, ii) los derechos de propiedad "
            "intelectual, iii) los compromisos entre la Red TecnoParque SENA y la(s)/el(los) Titular(es) "
            "del proyecto, acorde con lo dispuesto por los Acuerdos 09 de 2010 y 03 de 2012 del Consejo "
            "Directivo Nacional y la Guía Metodológica Red TecnoParque Colombia, o las normas y "
            "documentos que los sustituyan.",
            estilo_normal,
        )
    )

    historia.append(crear_parrafo("En mérito de lo expuesto se", estilo_normal))
    historia.append(crear_parrafo("ACUERDA:", estilo_subtitulo))

    historia.append(crear_parrafo("CAPITULO I<br/>DE LA CONFIDENCIALIDAD", estilo_subtitulo))

    historia.append(
        crear_parrafo(
            "<b>PRIMERO. INFORMACIÓN CONFIDENCIAL.</b> De conformidad con lo dispuesto por el "
            "Acuerdo 03 de 2012, constituirá “Información Confidencial” las metodologías, tecnologías, "
            "planos, prototipos, programas de computador y propiedad intelectual e ideas de la(s)/el(los) "
            "Titular(es)) del proyecto. Esto es, para mayor detalle sin limitarse a lo enunciado: las obras "
            "protegidas por el derecho de autor, nuevas creaciones o signos distintivos objeto de propiedad "
            "industrial, técnicas, modelos, invenciones, know-how, procesos, algoritmos, programas, ejecutables, "
            "investigaciones, detalles de diseño, información financiera, lista de clientes, bases de datos, "
            "inversionistas, empleados, relaciones de negocios y contractuales, pronósticos de negocios, planes "
            "de mercadeo. Se considera igualmente información confidencial, a) cualquier información revelada, "
            "sobre terceras personas y que no sea de dominio público u obvio, antes de la firma de la presente "
            "acta, b) la que no sea de fácil acceso, c) aquella información que no esté sujeta a medidas de "
            "protección razonables, de acuerdo con las circunstancias del caso, a fin de mantener su carácter "
            "confidencial. Los resultados del proyecto pueden considerarse confidenciales, si la(s)/el(los) "
            "Titular(es) así lo define(n).",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>SEGUNDO. MANEJO DE LA INFORMACIÓN CONFIDENCIAL.</b> Las partes acuerdan que "
            "cualquier información confidencial intercambiada, facilitada o creada entre ellas durante el "
            "desarrollo del proyecto, será mantenida en estricta reserva. El/la experto/a, Dinamizador/a, "
            "Infocenter y en general la Red Tecnoparque sólo podrá revelar información confidencial a quienes "
            "la necesiten y estén autorizados previamente por la(s)/el(los) Titular(es) que firman este documento. "
            "Así mismo, la(s)/el(los) Titular(es) que se incorpora(n) a la Red TecnoParque, deberán mantener en "
            "total reserva la información confidencial obtenida de otros Titulares, Experto, etc.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "En el caso de que la Red TecnoParque SENA requiera usar información catalogada como "
            "confidencial para el desarrollo de otros proyectos, deberá ser autorizado por escrito por "
            "la(s)/el(los) Titular(es) propietaria(os) de esta información.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "El SENA podrá hacer uso de los resultados de los proyectos, cuando ello sea necesario o "
            "requerido por los procesos de formación profesional, respetando siempre los derechos de "
            "propiedad intelectual sobre los resultados y la reserva de la información confidencial.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>TERCERO. EXCEPCIONES.</b> Lo datos intercambiados entre las partes no constituyen "
            "información confidencial en los siguientes casos: a) Cuando la parte receptora tenga "
            "evidencia de que conoce previamente la información recibida; b) Cuando la información "
            "recibida sea de dominio público; c) Cuando la información es revelada por el propietario y "
            "este acepta que puede ser utilizada como información de dominio público.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>CUARTO. DURACIÓN.</b> Las condiciones para el manejo de información confidencial que "
            "asume la(s)/el(los) Titular(es) a través de este acuerdo regirán durante el tiempo que dure "
            "el desarrollo del proyecto y cinco (5) años más, en ausencia de un acuerdo diferente entre "
            "la(s)/el(los) Titular(es) del proyecto y el Nodo de la Red Tecnoparque.",
            estilo_normal,
        )
    )

    historia.append(crear_parrafo("CAPITULO II<br/>DE LOS DERECHOS DE PROPIEDAD INTELECTUAL", estilo_subtitulo))

    textos_capitulo_ii = [
        "<b>QUINTO. TITULAR DEL PROYECTO.</b> Para los efectos del presente acuerdo y la prestación de servicios de Red Tecnoparque Colombia, se entiende por Titular del proyecto la(s) persona(s) u entidad(es) que lo idea(n), formula(n) y trabaja(n) en su implementación y son propietario(s) de los derechos patrimoniales derivados.",
        "<b>SEXTO. INTERLOCUTOR DEL PROYECTO.</b> Para los efectos del presente acuerdo y la prestación de servicios de Red Tecnoparque Colombia, se entiende por Interlocutor del proyecto la persona actúa en representación, manejo de las comunicaciones y toma de decisiones con la Red Tecnoparque para el (los) Titular(es) del proyecto.",
        "<b>SÉPTIMO. EJECUTOR DEL PROYECTO.</b> Para los efectos del presente acuerdo y la prestación de servicios de Red Tecnoparque Colombia, se entiende por Ejecutor del proyecto la(s) persona(s) que formula(n), trabaja(n) y apoya(n) en la implementación del proyecto con o sin relación contractual con el tercero, persona natural o jurídica, Titular del proyecto.",
        "<b>OCTAVO. TITULARIDAD DE LOS DERECHOS DE PROPIEDAD INTELECTUAL.</b> De conformidad con el Acuerdo 09 de 2010, por el cual se establecen las políticas para el programa de TecnoAcademias y TecnoParque, Capítulo II de los TecnoParque, los derechos de propiedad intelectual del proyecto desarrollado en la Red TecnoParque serán de sus autores/inventores, es decir, de la(s)/el(los) Titular(es) del proyecto, de conformidad con las normas vigentes que regulan la materia.",
        "Es responsabilidad de la(s)/el(los) Titular(es) iniciar los procesos de protección de la propiedad industrial de los productos, procesos o diseños que resulte de su proyecto, si considera que son susceptibles de algún mecanismo de protección nacional o internacional.",
        "El manejo de derechos de autor respecto del Interlocutor/a y ejecutor/a, se realizará en los siguientes términos:",
        "En los casos en que resulten del proyecto obras susceptibles de protección de derechos de autor, acorde con las disposiciones que regulan la materia, el/la Interlocutor/a del proyecto y/o ejecutor/a, será(n) titular(es) de los derechos morales de autor, siempre que su contribución a la obra no haya sido puramente física o mecánica, es decir, en los casos en que pueda predicarse su calidad de coautor(es)/a(s). La titularidad de los derechos patrimoniales de autor, por su parte, obedecerá a lo pactado en la relación contractual entre el/la Interlocutor/a y la persona natural o jurídica contratante que representa; en ausencia de pacto, la titularidad de los derechos patrimoniales se definirá en los términos que se acuerde entre el/la Interlocutor/a y/o el/la ejecutor/a y la persona natural o jurídica Titular del proyecto, o por lo establecido en el artículo 28 de la Ley 1450 de 2011 que modifica el artículo 20 de la Ley 23 de 1982.",
    ]

    for texto in textos_capitulo_ii:
        historia.append(crear_parrafo(texto, estilo_normal))

    historia.append(crear_parrafo("CAPITULO III<br/>DE LOS COMPROMISOS Y EL DESARROLLO DE PROYECTOS Y PROTOTIPOS", estilo_subtitulo))

    historia.append(
        crear_parrafo(
            "<b>NOVENO. COMPROMISOS.</b> El desarrollo de proyectos de base tecnológica en la Red "
            "TecnoParque SENA conlleva asumir los compromisos que se enuncian en el presente "
            "documento, encaminados a optimizar el tiempo, los recursos invertidos en el desarrollo del "
            "proyecto, así como sus resultados y beneficios.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>DÉCIMO. COMPROMISOS DE LA RED TECNOPARQUE SENA.</b> Es responsable de ofrecer "
            "sin ningún costo, asesoría técnica especializada y personalizada, herramientas e infraestructura "
            "necesaria para el desarrollo de iniciativas novedosas de productos y servicios de base tecnológica, "
            "susceptible de ser materializada en prototipos funcionales, ofreciendo adicionalmente:",
            estilo_normal,
        )
    )

    compromisos_red = [
        "Orientación sobre entidades de fortalecimiento empresarial y financiación.",
        "Acceso y uso de la infraestructura tecnológica en los horarios de servicio establecidos por cada Nodo.",
        "Oportunidades para participar en diferentes eventos como ferias, transferencia de tecnología, talleres y seminarios técnicos, encuentros tecnológicos, muestras empresariales, ruedas de negocios, entre otros, teniendo en cuenta los parámetros de selección que defina la Red Tecnoparque y el SENA.",
        "Cumplimiento del cronograma de trabajo definido entre los Titulares y los Experto de la Red TecnoParque SENA, en donde los Experto de la Red, cumplen con el servicio de asesoría técnica especializada y personalizada a los proyectos.",
        "Ofrecer el servicio de acceso a laboratorios en óptimas condiciones, garantizando el buen uso de la infraestructura.",
        "Contar con profesionales idóneos para ofrecer un servicio de calidad en el acompañamiento a la ejecución y asesoría a las iniciativas innovadoras de base tecnológica que se desarrollan al interior de la Red TecnoParque.",
    ]

    for i, item in enumerate(compromisos_red, start=1):
        historia.append(crear_parrafo(f"{i}. {item}", estilo_normal))

    historia.append(
        crear_parrafo(
            "<b>PARÁGRAFO.</b> La Red TecnoParque SENA, NO financia ninguna clase de materiales, "
            "insumos, equipos, membresías, pagos, viajes, papelería, para el desarrollo de proyectos, "
            "construcción o comercialización de prototipos.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>DECIMOPRIMERO. COMPROMISOS DE LOS TITULARES Y/O SUS INTERLOCUTORES.</b> "
            "Por medio de la presente acta se comprometen a:",
            estilo_normal,
        )
    )

    compromisos_titulares = [
        "Elaborar los documentos de Inicio, planeación, ejecución y cierre avalados por la Red TecnoParque SENA, luego de la firma de la presente Acta.",
        "Entregar a tiempo todos los documentos y las evidencias solicitadas por los expertos del Nodo, utilizando las herramientas de gestión dispuestas para tal fin.",
        "Cumplir con un horario de asistencia mínimo de horas semanales de trabajo autónomo presencial en el Nodo, el cual es establecido en común acuerdo con el experto asignado. Adicionalmente deberá cumplir con dos (2) horas semanales de acompañamiento técnico por el/la experto/a para el desarrollo del proyecto.",
        "Asistir al comité de seguimiento del proyecto y presentar al experto asignado los avances en un informe, en donde se dará cumplimiento a los objetivos planteados al inicio del proceso y con las respectivas evidencias (fotos, videos, simulaciones, diseños, entre otras) que lo respalden, lo anterior como mecanismo de autoevaluación y seguimiento, por lo tanto, es de carácter obligatorio.",
        "Asistir a las reuniones programadas por el Nodo.",
        "Del comportamiento: a) Mantener en todos los momentos (eventos, talleres, seminarios, trabajo en laboratorios, etc.) y espacios institucionales del SENA, un trato de respeto y buena convivencia. b) Utilizar la indumentaria y los elementos de protección personal dispuestos y/o solicitados por el Experto a cargo del laboratorio. c) Conservar y mantener en buen estado, orden y aseo, las instalaciones físicas, equipos y herramientas de la entidad o que estén a cargo de ésta, respondiendo por los daños ocasionados a éstos intencionalmente o por descuido, debidamente comprobados.",
        "No realizar actividades diferentes a las requeridas por el proyecto dentro de instalaciones del Nodo o no avaladas por la Red, en caso de presentarse la necesidad deberá contar con la autorización del Experto asignado al proyecto.",
        "Una vez finalizado el proyecto, se firmará un Acta de Cierre, en donde la(s)/el(los) Titular(es) entregará(n) las evidencias de finalización como fotos, videos, simulaciones, diseños e informes correspondientes.",
        "La(s)/el(los) Titular(es) en contraprestación a los servicios recibidos por la Red, realizará promoción y difusión de la Imagen Red TecnoParque SENA, esto durante el desarrollo del proyecto y una vez finalizado. Para ello utilizará el Logo SENA/Tecnoparque, el cual estará acompañado de la siguiente frase: “Apoyado por la Red Tecnoparque”, impreso y pegado sobre el prototipo del producto/servicio. Nunca en productos comerciales ni en prototipos en proceso de patente.",
        "Una vez finalizado el proyecto, asistir a la rueda de iniciativas empresariales, evento programado por el Nodo para la muestra, proyección y difusión de las iniciativas gestadas con el apoyo de la institución, para ello se deberán tener en cuenta las pautas para la selección de las iniciativas empresariales a presentar en el evento, estas pautas son diseñadas acorde a las particularidades de la región y el Nodo en el que se desarrollaron los proyectos.",
        "Conocer, aceptar y dar cumplimiento a los términos para uso de infraestructura adecuado de los diferentes laboratorios y equipos de la Red Tecnoparque SENA, incluyendo las medidas de Bioseguridad pertinentes en cada nodo y laboratorio.",
        "Programar, coordinar y asegurar la asistencia y las actividades de trabajo del equipo de Talentos Ejecutores del Proyecto.",
        "Cumplir con todos los protocolos de bioseguridad de los diferentes espacios a utilizar en cada Nodo.",
    ]

    for i, item in enumerate(compromisos_titulares, start=1):
        historia.append(crear_parrafo(f"{i}. {item}", estilo_normal))

    historia.append(
        crear_parrafo(
            "<b>DECIMOSEGUNDO. COMPROMISOS DE LOS TALENTOS EJECUTORES.</b> Por medio de la presente acta se comprometen a:",
            estilo_normal,
        )
    )

    compromisos_ejecutores = [
        "Entregar a tiempo todos los documentos y las evidencias solicitadas por los expertos del Nodo, utilizando las herramientas de gestión dispuestas para tal fin.",
        "Cumplir con el horario de asistencia acordado con el Talento interlocutor.",
        "Asistir al comité de seguimiento del proyecto y presentar al experto asignado los avances en un informe, en donde se dará cumplimiento a los objetivos planteados al inicio del proceso y con las respectivas evidencias (fotos, videos, simulaciones, diseños, entre otras) que lo respalden, lo anterior como mecanismo de autoevaluación y seguimiento, por lo tanto, es de carácter obligatorio.",
        "Asistir a las reuniones programadas por el Nodo.",
        "Del comportamiento: a) Mantener en todos los momentos (eventos, talleres, seminarios, trabajo en laboratorios, etc.) y espacios institucionales del SENA, un trato de respeto y buena convivencia. b) Utilizar la indumentaria y los elementos de protección personal dispuestos y/o solicitados por el Experto a cargo del laboratorio. c) Conservar y mantener en buen estado, orden y aseo, las instalaciones físicas, equipos y herramientas de la entidad o que estén a cargo de ésta, respondiendo por los daños ocasionados a éstos intencionalmente o por descuido, debidamente comprobados.",
        "No realizar actividades diferentes a las requeridas por el proyecto dentro de instalaciones del Nodo o no avaladas por la Red, en caso de presentarse la necesidad deberá contar con la autorización del Experto asignado al proyecto.",
        "Conocer, aceptar y dar cumplimiento a los términos para uso de infraestructura adecuado de los diferentes laboratorios y equipos de la Red Tecnoparque SENA, incluyendo las medidas de Bioseguridad pertinentes en cada nodo y laboratorio.",
        "Cumplir con todos los protocolos de bioseguridad de los diferentes espacios a utilizar en cada Nodo.",
    ]

    for i, item in enumerate(compromisos_ejecutores, start=1):
        historia.append(crear_parrafo(f"{i}. {item}", estilo_normal))

    historia.append(
        crear_parrafo(
            "<b>DECIMOTERCERO. TRANSFERENCIA DE CONOCIMIENTO.</b> En contrapartida por haber "
            "recibido el servicio de Asesoría técnica especializada y usos de infraestructura, en el desarrollo "
            "de proyectos de Base Tecnológica, la(s)/el(los) Titular(es) cumplirá(n) con alguna(s) de las siguientes "
            "actividades de Transferencia de Conocimiento. Éstas se ejecutan dentro del tiempo en el que "
            "la(s)/el(los) Titular(es) está(n) recibiendo el servicio mencionado y se definen y cumplen de acuerdo "
            "con los cronogramas de trabajo que se construyen entre Experto y Titulares en la etapa de planeación "
            "del proyecto:",
            estilo_normal,
        )
    )

    transferencias = [
        "Desarrollar Charlas Informativas de casos de éxito.",
        "Participar como ponente en un evento de divulgación tecnológica hacia empresa o academia.",
        "Participar en eventos en representación del SENA.",
        "Apoyar procesos de actualización a Experto Tecnoparque a través de transferencias de conocimiento.",
    ]

    for i, item in enumerate(transferencias, start=1):
        historia.append(crear_parrafo(f"{i}. {item}", estilo_normal))

    historia.append(
        crear_parrafo(
            "La(s)/el(los) Titular(es) además debe(n) diligenciar en su totalidad los documentos entregables "
            "y entregar uno de los productos, de acuerdo con la fase en la que se encuentre el proyecto y con "
            "los formatos establecidos por la Red, para ello deberá entregar una cuenta de correo personal o "
            "empresarial y compartir los documentos solicitados por el Experto a cargo de las asesorías del proyecto. "
            "Los documentos en los que debe participar son:",
            estilo_normal,
        )
    )

    documentos_soporte = [
        "a. Acta de inicio",
        "b. Actas de ejecución",
        "c. Encuesta de satisfacción",
        "d. Acta de cierre",
        "e. Documentos soporte",
    ]

    for item in documentos_soporte:
        historia.append(crear_parrafo(item, estilo_normal))

    historia.append(
        crear_parrafo(
            "<b>DECIMOCUARTO. INCUMPLIMIENTO DE COMPROMISOS DEL TITULAR.</b> El incumplimiento "
            "de los compromisos adquiridos por la(s)/el(los) Titular(es) dará lugar a la aplicación de las medidas "
            "restrictivas que se contemplan a continuación, dependiendo de la naturaleza del incumplimiento.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>DECIMOQUINTO MEDIDAS RESTRICTIVAS.</b> Dependiendo de la naturaleza del incumplimiento "
            "de los compromisos adquiridos por la(s) persona(s) Titular(es), se aplicarán las siguientes medidas restrictivas:",
            estilo_normal,
        )
    )

    medidas = [
        "<b>Restricción de acceso:</b> restricción de acceso a herramientas, laboratorios, equipos especializados y asesorías y pérdida de privilegios de horarios, durante un (1) mes, cuando la(s) persona(s) Titular(es), incumplan reiteradamente las citas y horarios programados con los expertos, que estén incumpliendo con el plan de trabajo injustificadamente y/o cuando no presenten los informes con evidencias de avance.",
        "<b>Restricción temporal de eventos:</b> restricción de acceso a cierto tipo de eventos durante un periodo de tres (3) meses. Aplica para personas que se hayan inscrito en talleres, charlas y actividades complementarias y no hayan asistido quitándole el cupo o la oportunidad a otras personas de participar.",
        "<b>Suspensión temporal:</b> suspensión de todos los servicios ofrecidos por la Red TecnoParque SENA, durante un periodo igual 30 días hábiles, cuando la(s) persona(s) Titular(es) no asista(n) al comité de seguimiento o se ausente(n) por más de cuatro (4) semanas al Nodo sin previa notificación o justificación.",
        "<b>Cancelación del proyecto:</b> se presenta cuando la(s) persona(s) Titular(es) se ausente(n) de las actividades de la Red Tecnoparque por un periodo superior a un (1) mes sin previa notificación o justificación. Durante los seis (6) meses siguientes a la cancelación del proyecto, no se podrá prestar proyectos al Comité de Ideas.",
    ]

    for i, item in enumerate(medidas, start=1):
        historia.append(crear_parrafo(f"{i}. {item}", estilo_normal))

    historia.append(
        crear_parrafo(
            "<b>DECIMOSEXTO. MODIFICACIÓN O TERMINACIÓN.</b> Este acuerdo sólo podrá ser modificado "
            "o darse por terminado con el consentimiento expreso por escrito de ambas partes antes o en el Acta de Cierre del proyecto.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            "<b>DECIMOSÉPTIMO. FIRMA DEL DOCUMENTO.</b> Para la firma de este documento en los casos "
            "de los menores de edad, este deberá ser avalado y firmado por un acudiente mayor de edad quien "
            "también firmará el presente acuerdo, aceptando todas las políticas y manuales vigentes de la Red "
            "TecnoParque SENA.",
            estilo_normal,
        )
    )

    historia.append(
        crear_parrafo(
            f"Para constancia, se firma el documento en la ciudad de {ciudad} a los "
            f"{fecha_larga_espanol_con_del(fecha_documento)}, por las partes:",
            estilo_normal,
        )
    )

    historia.append(Spacer(1, 10))
    historia.append(crear_parrafo("Firmas", estilo_subtitulo))

    firmas = [
        {
            "firma": ruta_firma_talento,
            "nombre": nombre_talento,
            "cargo": "Nombre Talento Interlocutor",
            "cedula": f"C.C. {cedula_talento}",
        },
        {
            "firma": ruta_firma("fsergio.png"),
            "nombre": "Sergio Andrés Cabrera",
            "cargo": "Nombre del Experto Tecnoparque",
            "cedula": "C.C. 1.110.454.504",
        },
        {
            "firma": ruta_firma("fcaro.png"),
            "nombre": "Carolina Garcia Monje",
            "cargo": "Nombre del Experto Tecnoparque",
            "cedula": "C.C. 36.301.495",
        },
        {
            "firma": ruta_firma("fdiego.png"),
            "nombre": "Diego Alfonso Polania",
            "cargo": "Nombre del Experto Tecnoparque",
            "cedula": "C.C. 7.684.683",
        },
        {
            "firma": ruta_firma("fcesar.png"),
            "nombre": "Cesar Augusto Pérez Tafur",
            "cargo": "Nombre del Experto Tecnoparque",
            "cedula": "C.C. 7.728.013",
        },
        {
            "firma": ruta_firma("fmaria.png"),
            "nombre": "Maria Andrea Qimbaya",
            "cargo": "Nombre del Apoyo Administrativo",
            "cedula": "C.C. 1003.812.026",
        },
        {
            "firma": ruta_firma("ffelix.png"),
            "nombre": "Felix Augusto Reyes Gutierrez",
            "cargo": "Profesional Grado 10.",
            "cedula": "C.C. 93407279",
        },
    ]

    tabla_firmas_data = [
        [
            crear_parrafo("<b>Nombre, cargo y cédula</b>", estilo_tabla_negrita),
            crear_parrafo("<b>Firma</b>", estilo_tabla_negrita),
        ]
    ]

    for firmante in firmas:
        tabla_firmas_data.append(
            [
                crear_parrafo(
                    f"<b>{firmante['nombre']}</b><br/>{firmante['cargo']}<br/>{firmante['cedula']}",
                    estilo_tabla,
                ),
                firma_img(firmante["firma"], width=4.0 * cm, height=1.05 * cm),
            ]
        )

    tabla_firmas = Table(
        tabla_firmas_data,
        colWidths=[10.2 * cm, 5.0 * cm],
        rowHeights=[0.75 * cm] + [1.65 * cm for _ in firmas],
    )

    tabla_firmas.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    historia.append(tabla_firmas)

    doc.build(
        historia,
        onFirstPage=encabezado_pie,
        onLaterPages=encabezado_pie,
    )

    datos_json = dict(datos)
    datos_json["fecha_documento"] = fecha_documento.strftime("%d/%m/%Y")
    datos_json["ruta_pdf"] = ruta_pdf
    guardar_datos_json(datos_json, ruta="datos_confidencialidad_compromiso.json")

    return ruta_pdf



def render_confidencialidad(modo_prueba: bool = True, modelo_openai: str = "") -> None:
    st.markdown("---")
    st.subheader("Formulario de Confidencialidad y Compromiso")
    st.caption(VERSION_CONFIDENCIALIDAD)

    st.info(
        "Este documento no consume API de OpenAI. Genera el formato GIC-F-041 V03 "
        "con datos dinámicos del proyecto, talento y firmas institucionales."
    )

    if "datos_confidencialidad_generada" not in st.session_state:
        st.session_state.datos_confidencialidad_generada = None

    if "ruta_pdf_confidencialidad_generado" not in st.session_state:
        st.session_state.ruta_pdf_confidencialidad_generado = None

    with st.expander("Verificación de firmas institucionales"):
        firmas_requeridas = [
            "fcaro.png",
            "fcesar.png",
            "fdiego.png",
            "ffelix.png",
            "fmaria.png",
            "fsergio.png",
        ]
        for firma in firmas_requeridas:
            ruta = Path(CARPETA_FIRMAS) / firma
            if ruta.exists():
                st.success(f"Encontrada: {ruta}")
            else:
                st.warning(f"No encontrada: {ruta}")

    with st.form("form_confidencialidad"):
        col_a, col_b = st.columns(2)

        with col_a:
            codigo_proyecto = st.text_input(
                "Código del proyecto",
                placeholder="Ejemplo: I2025-1421611-24230",
            )

            nombre_proyecto = st.text_area(
                "Nombre del proyecto",
                placeholder="Ejemplo: Sistema de Bombeo Eléctrico para Mieles de Café",
                height=100,
            )

            nombre_talento = st.text_input(
                "Nombres completos del talento",
                placeholder="Ejemplo: José Lizardo Ninco Ibarra",
            )

            cedula_talento = st.text_input(
                "Cédula del talento",
                placeholder="Ejemplo: 7.728.013",
            )

        with col_b:
            ciudad_expedicion = st.text_input(
                "Ciudad de expedición de la cédula",
                placeholder="Ejemplo: Neiva-Huila",
            )

            ciudad = st.text_input(
                "Ciudad del documento",
                value="Campoalegre",
            )

            fecha_documento = st.date_input(
                "Fecha del documento",
                value=date.today(),
            )

            firma_talento_upload = st.file_uploader(
                "Firma del talento en PNG/JPG",
                type=["png", "jpg", "jpeg"],
                help="Opcional. Si no se carga, quedará el espacio de firma en blanco.",
            )

        generar_confidencialidad = st.form_submit_button(
            "Generar documento de confidencialidad y compromiso"
        )

    if generar_confidencialidad:
        campos_obligatorios = {
            "Código del proyecto": codigo_proyecto,
            "Nombre del proyecto": nombre_proyecto,
            "Nombres completos del talento": nombre_talento,
            "Cédula del talento": cedula_talento,
            "Ciudad de expedición de la cédula": ciudad_expedicion,
            "Ciudad del documento": ciudad,
        }

        if not validar_campos_obligatorios(campos_obligatorios):
            st.stop()

        try:
            ruta_firma_talento_tmp = guardar_archivo_subido(
                firma_talento_upload,
                "firma_talento",
            )
        except Exception as error:
            st.error(f"No se pudo procesar la firma del talento: {error}")
            st.stop()

        datos_confidencialidad = {
            "tipo_documento": "Confidencialidad y compromiso",
            "codigo_proyecto": codigo_proyecto,
            "nombre_proyecto": nombre_proyecto,
            "nombre_talento": nombre_talento,
            "cedula_talento": cedula_talento,
            "ciudad_expedicion": ciudad_expedicion,
            "ciudad": ciudad,
            "fecha_documento": fecha_documento,
            "fecha_corta": fecha_documento.strftime("%d/%m/%Y"),
            "fecha_iso": fecha_documento.strftime("%Y-%m-%d"),
            "ruta_firma_talento": ruta_firma_talento_tmp,
            "version": VERSION_CONFIDENCIALIDAD,
        }

        st.session_state.datos_confidencialidad_generada = datos_confidencialidad
        st.session_state.ruta_pdf_confidencialidad_generado = None

        st.success("Información registrada correctamente. Ahora puedes generar el PDF.")

    if st.session_state.datos_confidencialidad_generada:
        datos_confidencialidad = st.session_state.datos_confidencialidad_generada

        st.markdown("## Resumen para validación")
        st.write("**Tipo de documento:**", datos_confidencialidad["tipo_documento"])
        st.write("**Código del proyecto:**", datos_confidencialidad["codigo_proyecto"])
        st.write("**Nombre del proyecto:**", datos_confidencialidad["nombre_proyecto"])
        st.write("**Talento:**", datos_confidencialidad["nombre_talento"])
        st.write("**Cédula:**", datos_confidencialidad["cedula_talento"])
        st.write("**Ciudad de expedición:**", datos_confidencialidad["ciudad_expedicion"])
        st.write(
            "**Ciudad y fecha:**",
            f'{datos_confidencialidad["ciudad"]}, {datos_confidencialidad["fecha_corta"]}',
        )

        if datos_confidencialidad.get("ruta_firma_talento"):
            st.success("Firma del talento cargada correctamente.")
        else:
            st.warning(
                "No se cargó firma del talento. El documento se generará con el espacio en blanco."
            )

        col_json, col_pdf = st.columns(2)

        with col_json:
            datos_json_descarga = dict(datos_confidencialidad)
            if isinstance(datos_json_descarga.get("fecha_documento"), date):
                datos_json_descarga["fecha_documento"] = datos_json_descarga[
                    "fecha_documento"
                ].strftime("%d/%m/%Y")

            st.download_button(
                label="Descargar datos en JSON",
                data=json.dumps(datos_json_descarga, ensure_ascii=False, indent=4),
                file_name="datos_confidencialidad_compromiso.json",
                mime="application/json",
            )

        with col_pdf:
            if st.button("📄 Generar PDF de confidencialidad y compromiso"):
                try:
                    ruta_pdf = generar_pdf_confidencialidad(datos_confidencialidad)
                    st.session_state.ruta_pdf_confidencialidad_generado = ruta_pdf
                    st.success(f"PDF generado correctamente: {ruta_pdf}")
                except Exception as error:
                    st.error(f"No se pudo generar el PDF: {error}")

        if (
            st.session_state.ruta_pdf_confidencialidad_generado
            and Path(st.session_state.ruta_pdf_confidencialidad_generado).exists()
        ):
            ruta_pdf = st.session_state.ruta_pdf_confidencialidad_generado

            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar PDF de confidencialidad y compromiso",
                    data=f,
                    file_name=Path(ruta_pdf).name,
                    mime="application/pdf",
                )
