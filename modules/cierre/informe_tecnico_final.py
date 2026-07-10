# -*- coding: utf-8 -*-
"""
Módulo: Informe Técnico Final / Informe Final GCDTP-F-023 V01
Autor: Generado para TecnoParque
Descripción:
    Genera el informe final en Word usando la plantilla oficial, con:
    - Nombre del archivo basado en codigo_proyecto.
    - Propiedades internas del Word limitadas a 255 caracteres.
    - Tabla de contenido con campos automáticos de Word.
    - Títulos con estilos Heading 1 / Heading 2 para que Word actualice páginas reales.
    - Formulario sin carga de evidencias.
    - Actividades solicitadas únicamente como descripción.
    - Tabla de resultados con N.°, descripción de actividad ejecutada y evidencia.
    - Estado del arte con dos antecedentes reales, citas y referencias.
    - Propiedad intelectual sugerida automáticamente según el tipo de proyecto.
    - Conclusiones con recomendación de continuidad según TRL alcanzado.

Notas importantes:
    1. Si usas Streamlit Cloud, asegúrate de que la plantilla exista en:
       resources/GCDTP-F-023_V01_Formato_Informe_Final.docx

    2. Para que la tabla de contenido muestre los números de página reales,
       el documento queda configurado para actualizar campos al abrirse en Word.
       En algunos equipos Word pide confirmar la actualización de campos.
"""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    import streamlit as st
except Exception:  # Permite importar el módulo sin Streamlit en pruebas unitarias
    st = None

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import _Cell


# =========================================================
# RUTAS Y CONSTANTES
# =========================================================

BASE_DIR = Path(__file__).resolve().parents[2] if len(Path(__file__).resolve().parents) >= 3 else Path.cwd()
RESOURCES_DIR = BASE_DIR / "resources"
OUTPUT_DIR = BASE_DIR / "output"

CODIGO_FORMATO_INFORME = "GCDTP-F-023 V01"
NOMBRE_PLANTILLA_INFORME = "GCDTP-F-023_V01_Formato_Informe_Final.docx"
RUTA_PLANTILLA_INFORME = RESOURCES_DIR / NOMBRE_PLANTILLA_INFORME


# =========================================================
# UTILIDADES GENERALES
# =========================================================

def limpiar_nombre_archivo(texto: Any) -> str:
    """
    Limpia un texto para usarlo como nombre de archivo en Windows, Linux y macOS.
    """
    texto = str(texto or "SIN_CODIGO").strip()
    texto = re.sub(r'[<>:"/\\|?*\n\r\t]+', "-", texto)
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.strip(" .-")
    return texto or "SIN_CODIGO"


def limpiar_texto_propiedad_word(texto: Any, limite: int = 250) -> str:
    """
    Word limita varias propiedades internas del documento a 255 caracteres.
    Esta función evita el error: exceeded 255 char limit for property.
    """
    texto = str(texto or "").replace("\n", " ").replace("\r", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    if len(texto) > limite:
        return texto[: limite - 3] + "..."
    return texto


def normalizar_espacios(texto: Any) -> str:
    texto = str(texto or "").replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def obtener_dato(datos: Dict[str, Any], *claves: str, defecto: str = "") -> str:
    """
    Obtiene el primer valor disponible entre varias claves posibles.
    Esto ayuda cuando el formulario usa nombres diferentes para el mismo campo.
    """
    for clave in claves:
        valor = datos.get(clave)
        if valor not in (None, ""):
            return str(valor).strip()
    return defecto


def obtener_lista(datos: Dict[str, Any], *claves: str) -> List[str]:
    """
    Convierte un campo en lista. Acepta lista, tupla o texto con saltos de línea.
    """
    valor: Any = None
    for clave in claves:
        if datos.get(clave) not in (None, ""):
            valor = datos.get(clave)
            break

    if valor is None:
        return []

    if isinstance(valor, (list, tuple)):
        elementos = [str(x).strip() for x in valor if str(x).strip()]
    else:
        texto = str(valor).strip()
        elementos = [
            linea.strip(" -•\t")
            for linea in re.split(r"\n|;", texto)
            if linea.strip(" -•\t")
        ]

    return elementos


def extraer_trl(valor: Any) -> Optional[int]:
    texto = str(valor or "").upper()
    match = re.search(r"TRL\s*([1-9])", texto)
    if match:
        return int(match.group(1))

    match = re.search(r"\b([1-9])\b", texto)
    if match:
        return int(match.group(1))

    return None


# =========================================================
# CORRECCIÓN BÁSICA Y MEJORA DE REDACCIÓN
# =========================================================

REEMPLAZOS_BÁSICOS = {
    " poryecto ": " proyecto ",
    " proyeto ": " proyecto ",
    " docuemnto ": " documento ",
    " docuemento ": " documento ",
    " codigos ": " códigos ",
    " codigo ": " código ",
    " tecnologias ": " tecnologías ",
    " tecnologia ": " tecnología ",
    " implementacion ": " implementación ",
    " informacion ": " información ",
    " descripcion ": " descripción ",
    " avistamiento ": " avistamiento ",
    " diseno ": " diseño ",
    " senalizacion ": " señalización ",
    " ecologico ": " ecológico ",
    " interactivo ": " interactivo ",
    " multimedia ": " multimedia ",
    " hostin ": " hosting ",
    " hostingg ": " hosting ",
    " dominioo ": " dominio ",
    " pagian ": " página ",
    " paginas ": " páginas ",
    " sonido ": " sonido ",
    " sonidos ": " sonidos ",
}


def corregir_ortografia_basica(texto: Any) -> str:
    """
    Corrección básica sin depender de servicios externos.
    No reemplaza una revisión humana, pero mejora errores frecuentes de digitación.
    """
    texto = f" {normalizar_espacios(texto)} "
    texto_min = texto

    for incorrecto, correcto in REEMPLAZOS_BÁSICOS.items():
        texto_min = re.sub(re.escape(incorrecto), correcto, texto_min, flags=re.IGNORECASE)

    texto_min = normalizar_espacios(texto_min)
    return texto_min


def capitalizar_oracion(texto: str) -> str:
    texto = texto.strip()
    if not texto:
        return texto
    return texto[0].upper() + texto[1:]


def mejorar_redaccion_actividad(texto: Any) -> str:
    """
    Toma una actividad escrita por el usuario, corrige aspectos básicos
    y la convierte en una oración formal para el informe.
    """
    texto = corregir_ortografia_basica(texto)
    texto = texto.strip(" .;:-")
    if not texto:
        return ""

    texto = capitalizar_oracion(texto)

    # Evita frases excesivamente informales.
    reemplazos = {
        r"\bse hizo\b": "se realizó",
        r"\bse montó\b": "se implementó",
        r"\bse subió\b": "se cargó",
        r"\bse creó\b": "se desarrolló",
        r"\bse pusieron\b": "se instalaron",
    }

    for patron, reemplazo in reemplazos.items():
        texto = re.sub(patron, reemplazo, texto, flags=re.IGNORECASE)

    if not texto.endswith("."):
        texto += "."

    return texto


def mejorar_lista_actividades(actividades: Iterable[str]) -> List[str]:
    actividades_mejoradas = []
    for actividad in actividades:
        actividad_mejorada = mejorar_redaccion_actividad(actividad)
        if actividad_mejorada:
            actividades_mejoradas.append(actividad_mejorada)
    return actividades_mejoradas


# =========================================================
# GENERACIÓN DE TEXTOS POR APARTADO
# =========================================================

def redactar_descripcion_general(datos: Dict[str, Any]) -> str:
    nombre = obtener_dato(datos, "nombre_proyecto", "proyecto", defecto="el proyecto")
    descripcion = corregir_ortografia_basica(
        obtener_dato(datos, "descripcion_proyecto", "descripcion", "resumen", defecto="")
    )
    beneficiario = obtener_dato(datos, "beneficiario", "empresa", "unidad_productiva", defecto="la unidad beneficiaria")
    municipio = obtener_dato(datos, "municipio", "ciudad", "ubicacion", defecto="el territorio de intervención")
    sector = obtener_dato(datos, "sector", "area", defecto="el contexto productivo y territorial asociado")

    if not descripcion:
        descripcion = (
            "La iniciativa se orientó al desarrollo de una solución tecnológica aplicada, "
            "con énfasis en la integración de recursos digitales, contenidos organizados "
            "y herramientas de apoyo para mejorar la experiencia de los usuarios finales."
        )

    return (
        f"El proyecto denominado “{nombre}” se desarrolló como una iniciativa de base tecnológica "
        f"orientada a fortalecer la propuesta de valor de {beneficiario} en {municipio}. "
        f"Su propósito central fue transformar una necesidad identificada en {sector} en una solución "
        f"funcional, verificable y pertinente, articulando recursos técnicos, diseño, organización de "
        f"información y criterios de usabilidad para facilitar su implementación en un entorno real. "
        f"{descripcion} "
        f"Desde el enfoque de TecnoParque, el proyecto permitió acompañar la maduración de una idea "
        f"hacia un resultado tecnológico concreto, con énfasis en la apropiación de herramientas "
        f"digitales, la generación de capacidades en el beneficiario y la validación de una solución "
        f"útil para mejorar procesos, experiencias o servicios. La intervención se estructuró de manera "
        f"progresiva, iniciando con el reconocimiento de requerimientos, continuando con el diseño de "
        f"componentes, la construcción de recursos, la integración de contenidos y la verificación de "
        f"funcionamiento. Como alcance general, el proyecto no se limitó a la producción de elementos "
        f"aislados, sino que consolidó un sistema articulado entre el componente técnico, la experiencia "
        f"del usuario y las condiciones del entorno donde fue implementado. Esta descripción resume el "
        f"sentido del proyecto y sirve como marco general para comprender las secciones posteriores del "
        f"informe, en las cuales se detallan objetivos, metodología, actividades, resultados, propiedad "
        f"intelectual y recomendaciones de continuidad."
    )


def redactar_metodologia_y_desarrollo(datos: Dict[str, Any], actividades: List[str]) -> str:
    metodologia = obtener_dato(
        datos,
        "metodologia_utilizada",
        "metodologia",
        defecto="metodología aplicada de diseño, desarrollo, validación e implementación"
    )
    trl = obtener_dato(datos, "trl_alcanzado", "trl", defecto="el nivel de madurez tecnológica definido para el cierre")
    actividades_texto = " ".join(actividades[:6]) if actividades else (
        "Se ejecutaron actividades de levantamiento de información, diseño de la solución, "
        "construcción de componentes, pruebas funcionales e implementación final."
    )

    return (
        f"El desarrollo del proyecto se orientó bajo la metodología {metodologia}, entendida como una "
        f"ruta de trabajo organizada para pasar de la necesidad identificada a una solución funcional "
        f"y verificable. La metodología permitió estructurar el proceso en momentos sucesivos de "
        f"análisis, diseño, construcción, revisión e implementación, evitando que las actividades se "
        f"ejecutaran de manera aislada y facilitando la trazabilidad entre los objetivos propuestos y "
        f"los resultados alcanzados. En la fase inicial se precisaron los requerimientos del proyecto, "
        f"las condiciones del entorno de aplicación, los usuarios esperados y los recursos técnicos "
        f"necesarios para desarrollar la solución. Posteriormente, se definieron los componentes "
        f"funcionales y visuales, priorizando criterios de claridad, pertinencia, facilidad de uso y "
        f"posibilidad de actualización futura. Durante la etapa de construcción se desarrollaron los "
        f"elementos tecnológicos y de contenido necesarios, validando que cada componente respondiera "
        f"a la finalidad del proyecto. Entre las actividades ejecutadas se destacan: {actividades_texto} "
        f"Estas acciones fueron articuladas con la metodología seleccionada, de manera que cada actividad "
        f"aportara al avance del proyecto y no solo al cumplimiento operativo de tareas. En la fase de "
        f"validación se revisó el funcionamiento de la solución, la coherencia de los contenidos, la "
        f"interacción con los usuarios y la correspondencia con el alcance formulado. Finalmente, se "
        f"realizó la implementación o puesta en funcionamiento, dejando el proyecto en condiciones de "
        f"uso conforme al {trl}. La metodología aplicada permitió consolidar una solución ordenada, "
        f"documentada y alineada con las necesidades del beneficiario."
    )


def redactar_estado_del_arte(datos: Dict[str, Any]) -> str:
    nombre = obtener_dato(datos, "nombre_proyecto", "proyecto", defecto="el proyecto")
    enfoque = obtener_dato(datos, "sector", "area", defecto="la gestión de información digital y experiencias interactivas")

    return (
        f"El estado del arte del proyecto “{nombre}” se relaciona con soluciones digitales orientadas "
        f"a la consulta, organización y divulgación de información ambiental, turística o científica, "
        f"especialmente cuando estas integran contenidos multimedia, interacción desde dispositivos "
        f"móviles y datos provenientes de fuentes abiertas o colaborativas. Un antecedente relevante "
        f"es eBird, iniciativa del Cornell Lab of Ornithology que funciona como una plataforma global "
        f"para registrar observaciones de aves mediante listas aportadas por usuarios, generando una "
        f"base de datos utilizada en investigación, conservación y gestión ambiental (Cornell Lab of "
        f"Ornithology, s. f.-a). Este referente demuestra la importancia de estructurar información "
        f"biológica de forma consultable, georreferenciada y útil para diferentes públicos, desde "
        f"observadores aficionados hasta investigadores y gestores del territorio. Como segundo "
        f"antecedente se identifica Merlin Bird ID, también desarrollado por el Cornell Lab of "
        f"Ornithology, el cual integra fotografías, sonidos, mapas de distribución y asistencia para "
        f"la identificación de aves por medio de herramientas digitales orientadas al usuario final "
        f"(Cornell Lab of Ornithology, s. f.-b). La relación con {enfoque} permite evidenciar que las "
        f"soluciones actuales no se limitan a almacenar información, sino que buscan facilitar el "
        f"aprendizaje, la identificación, la navegación y la apropiación del conocimiento mediante "
        f"interfaces claras y experiencias interactivas. En ese sentido, el proyecto toma como referencia "
        f"la tendencia de articular datos, contenidos visuales, recursos sonoros, mapas y acceso móvil "
        f"para fortalecer experiencias educativas, ambientales y turísticas. Este apartado presenta una "
        f"síntesis de antecedentes pertinentes; el estado del arte completo se encuentra documentado en "
        f"la sección de documentos de planeación del proyecto."
    )


def redactar_resultados_generales(datos: Dict[str, Any], actividades: List[str]) -> str:
    resultado_usuario = corregir_ortografia_basica(
        obtener_dato(datos, "resultado_general", "resultados", "descripcion_resultados", defecto="")
    )
    trl = obtener_dato(datos, "trl_alcanzado", "trl", defecto="el TRL definido para el cierre")
    cantidad = len(actividades)

    if not resultado_usuario:
        resultado_usuario = (
            "Como resultado principal se obtuvo una solución funcional, organizada y verificable, "
            "implementada conforme al alcance definido y alineada con los objetivos del proyecto."
        )

    return (
        f"Los resultados del proyecto se evidencian en la consolidación de una solución funcional, "
        f"pertinente y alineada con las necesidades identificadas durante la etapa de formulación. "
        f"{resultado_usuario} La ejecución permitió pasar de una necesidad o idea inicial a un conjunto "
        f"de componentes implementados, validados y organizados para su uso por parte del beneficiario "
        f"o de los usuarios finales. En términos técnicos, los resultados muestran la integración de "
        f"los recursos diseñados, la disponibilidad de los contenidos o componentes desarrollados y la "
        f"verificación de su funcionamiento en el contexto previsto. También se logró fortalecer la "
        f"capacidad del beneficiario para comprender, usar y proyectar la solución hacia nuevas fases "
        f"de mejora o escalamiento. Las actividades registradas en la tabla siguiente permiten evidenciar "
        f"la trazabilidad entre lo ejecutado y los productos obtenidos durante el proceso. En total, se "
        f"relacionan {cantidad} actividad(es) principales, descritas de forma corregida y organizada para "
        f"facilitar su revisión dentro del documento. La columna de evidencia queda disponible para que "
        f"el usuario agregue posteriormente enlaces, imágenes o soportes directamente en Word, según "
        f"las evidencias finales disponibles. De acuerdo con el cierre del proyecto, los resultados son "
        f"coherentes con {trl} y permiten sustentar la continuidad técnica de la iniciativa sin afirmar "
        f"un nivel de madurez superior al alcanzado."
    )


def mecanismos_propiedad_intelectual(datos: Dict[str, Any], actividades: List[str]) -> List[str]:
    texto = " ".join([
        obtener_dato(datos, "nombre_proyecto", "proyecto"),
        obtener_dato(datos, "descripcion_proyecto", "descripcion"),
        obtener_dato(datos, "objetivo_general"),
        " ".join(obtener_lista(datos, "objetivos_especificos")),
        " ".join(actividades),
        obtener_dato(datos, "resultado_general", "resultados"),
    ]).lower()

    mecanismos: List[str] = []

    claves_software = [
        "software", "sistema web", "plataforma", "aplicación", "app", "código qr",
        "qr", "hosting", "dominio", "web", "base de datos", "dashboard", "scraping",
        "sitio web", "página web"
    ]
    claves_autor = [
        "contenido", "multimedia", "video", "audio", "sonoro", "texto", "cartilla",
        "manual", "diseño gráfico", "personaje", "ilustración", "fotografía", "animación"
    ]
    claves_marca = ["marca", "identidad visual", "logo", "nombre comercial", "emprendimiento"]
    claves_diseno = ["diseño 3d", "diseño industrial", "carcasa", "placa", "letrero", "producto físico", "prototipo físico"]
    claves_modelo = ["mecanismo", "dispositivo", "mejora técnica", "ensamble", "estructura funcional", "prototipo"]
    claves_secreto = ["algoritmo", "modelo predictivo", "fórmula", "know how", "proceso interno", "método propio"]

    if any(k in texto for k in claves_software):
        mecanismos.append(
            "registro de software ante la Dirección Nacional de Derecho de Autor, en caso de que exista código fuente, estructura funcional o desarrollo web propio"
        )

    if any(k in texto for k in claves_autor):
        mecanismos.append(
            "derecho de autor sobre contenidos, textos, interfaces, material gráfico, piezas multimedia, registros sonoros o recursos visuales desarrollados"
        )

    if any(k in texto for k in claves_marca):
        mecanismos.append(
            "registro de marca ante la Superintendencia de Industria y Comercio, si la identidad visual o denominación se proyecta para uso comercial"
        )

    if any(k in texto for k in claves_diseno):
        mecanismos.append(
            "diseño industrial, si los elementos físicos desarrollados presentan una apariencia nueva y diferenciable susceptible de protección"
        )

    if any(k in texto for k in claves_modelo):
        mecanismos.append(
            "modelo de utilidad, únicamente si el prototipo físico incorpora una mejora técnica funcional y verificable frente a soluciones existentes"
        )

    if any(k in texto for k in claves_secreto):
        mecanismos.append(
            "secreto empresarial, cuando existan métodos, algoritmos, procesos internos o conocimiento técnico no divulgado que generen ventaja competitiva"
        )

    # Patente solo se menciona si hay señales claras de invención técnica, para evitar listar opciones genéricas.
    if "invención" in texto or "nuevo procedimiento técnico" in texto or "solución técnica novedosa" in texto:
        mecanismos.append(
            "patente de invención, solo si se demuestra novedad, nivel inventivo y aplicación industrial conforme a la normativa colombiana"
        )

    if not mecanismos:
        mecanismos.append(
            "derecho de autor sobre la documentación técnica, contenidos y materiales generados durante el desarrollo del proyecto"
        )

    return mecanismos


def redactar_propiedad_intelectual(datos: Dict[str, Any], actividades: List[str]) -> str:
    mecanismos = mecanismos_propiedad_intelectual(datos, actividades)
    mecanismos_txt = "; ".join(mecanismos)

    return (
        f"A partir de la descripción, los objetivos, las actividades ejecutadas y los resultados del "
        f"proyecto, los mecanismos de protección más pertinentes en Colombia se relacionan con {mecanismos_txt}. "
        f"Esta recomendación no implica que la protección ya haya sido obtenida, sino que identifica las "
        f"rutas que podrían gestionarse según la naturaleza de los activos generados. Para avanzar en "
        f"cualquiera de estos mecanismos se recomienda conservar soportes de autoría, versiones de diseño, "
        f"evidencias de desarrollo, fechas de creación, archivos fuente, documentos técnicos y autorizaciones "
        f"de uso de contenidos de terceros cuando aplique. La transferencia tecnológica puede orientarse hacia "
        f"la apropiación de la solución por parte del beneficiario, su uso controlado en el entorno definido, "
        f"la capacitación de usuarios y la definición de condiciones para futuras actualizaciones, escalamiento "
        f"o reproducción en escenarios similares."
    )


def recomendacion_continuidad_trl(trl: Optional[int]) -> str:
    if trl is None:
        return (
            "Como recomendación de continuidad, se sugiere precisar el TRL alcanzado en el cierre del proyecto "
            "para formular una fase posterior coherente con el nivel real de madurez tecnológica."
        )

    if trl < 9:
        siguiente = trl + 1
        return (
            f"Como recomendación de continuidad, dado que el proyecto finaliza en TRL {trl}, es viable formular "
            f"una nueva fase orientada a avanzar hacia TRL {siguiente}, mediante validaciones adicionales, "
            f"mejoras técnicas, pruebas con usuarios o implementación en condiciones más cercanas al entorno "
            f"operativo, según corresponda. Esta recomendación no afirma que el TRL {siguiente} ya haya sido "
            f"alcanzado, sino que plantea una ruta posible para la maduración tecnológica."
        )

    return (
        "Como recomendación de continuidad, al tratarse de un proyecto ubicado en TRL 9, la fase siguiente "
        "debe orientarse a sostenibilidad, transferencia, escalamiento, mantenimiento, apropiación por usuarios "
        "y seguimiento de impacto, más que a afirmar un nuevo nivel de madurez tecnológica."
    )


def redactar_conclusiones(datos: Dict[str, Any], actividades: List[str]) -> str:
    trl = extraer_trl(obtener_dato(datos, "trl_alcanzado", "trl", defecto=""))
    continuidad = recomendacion_continuidad_trl(trl)

    return (
        "El proyecto permitió consolidar una solución tecnológica pertinente frente a la necesidad identificada, "
        "integrando actividades de análisis, diseño, desarrollo, validación e implementación de manera coherente "
        "con los objetivos formulados. Los resultados obtenidos evidencian que el proceso de acompañamiento "
        "contribuyó a transformar una idea o requerimiento inicial en una solución organizada, documentada y "
        "susceptible de ser utilizada por el beneficiario en el contexto definido. Además, el desarrollo favoreció "
        "la apropiación de capacidades técnicas, el reconocimiento de oportunidades de mejora y la generación de "
        "insumos útiles para futuras fases de maduración. La revisión de las actividades ejecutadas permite "
        "concluir que el proyecto avanzó de forma progresiva, manteniendo relación entre el alcance formulado, "
        "la metodología aplicada y los resultados presentados. También se identifican oportunidades de continuidad "
        "asociadas a validación con más usuarios, ampliación de funcionalidades, fortalecimiento documental, "
        "seguimiento de impacto y posibles mecanismos de protección de los activos generados. "
        f"{continuidad}"
    )


def referencias_estado_del_arte() -> List[str]:
    return [
        "Cornell Lab of Ornithology. (s. f.-a). eBird: An online database of bird distribution and abundance. https://ebird.org",
        "Cornell Lab of Ornithology. (s. f.-b). Merlin Bird ID. https://merlin.allaboutbirds.org",
    ]


# =========================================================
# FUNCIONES WORD
# =========================================================

def cargar_documento_base() -> DocumentObject:
    """
    Carga la plantilla oficial si existe. Si no existe, crea un documento vacío
    para no detener las pruebas locales.
    """
    if RUTA_PLANTILLA_INFORME.exists():
        return Document(str(RUTA_PLANTILLA_INFORME))

    documento = Document()
    documento.add_heading("Informe Final", level=0)
    return documento


def configurar_actualizacion_campos(documento: DocumentObject) -> None:
    """
    Configura Word para actualizar campos al abrir el documento.
    Esto ayuda a que la tabla de contenido actualice números de página reales.
    """
    settings = documento.settings.element
    existentes = settings.xpath("./w:updateFields")
    if existentes:
        existentes[0].set(qn("w:val"), "true")
        return

    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)


def agregar_campo_word(parrafo, instruccion: str) -> None:
    """
    Inserta un campo de Word, por ejemplo una tabla de contenido automática.
    """
    run = parrafo.add_run()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instruccion

    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr_text)
    run._r.append(fld_separate)
    run._r.append(fld_end)


def agregar_tabla_contenido(documento: DocumentObject) -> None:
    """
    Agrega tabla de contenido automática.
    Los números de página se actualizan en Word al abrir el archivo.
    """
    titulo = documento.add_paragraph()
    titulo.style = "Heading 1"
    titulo.add_run("Tabla de contenido")

    parrafo = documento.add_paragraph()
    agregar_campo_word(parrafo, r'TOC \o "1-3" \h \z \u')

    documento.add_page_break()


def agregar_titulo(documento: DocumentObject, texto: str, nivel: int = 1) -> None:
    """
    Agrega títulos usando estilos nativos de Word para que la tabla de contenido funcione.
    """
    nivel = max(1, min(3, int(nivel)))
    parrafo = documento.add_heading(texto, level=nivel)
    parrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT


def agregar_parrafo(documento: DocumentObject, texto: str) -> None:
    texto = normalizar_espacios(texto)
    if not texto:
        return

    parrafo = documento.add_paragraph()
    parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    parrafo.paragraph_format.space_after = Pt(6)
    parrafo.paragraph_format.line_spacing = 1.15
    run = parrafo.add_run(texto)
    run.font.name = "Arial"
    run.font.size = Pt(11)


def set_cell_text(cell: _Cell, text: str, bold: bool = False) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(10)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def agregar_tabla_identificacion(documento: DocumentObject, datos: Dict[str, Any]) -> None:
    codigo = obtener_dato(datos, "codigo_proyecto", "codigo", "codigo_tecnoparque", "codigo_sennova", "id_proyecto", defecto="SIN_CODIGO")
    nombre = obtener_dato(datos, "nombre_proyecto", "proyecto", defecto="")
    beneficiario = obtener_dato(datos, "beneficiario", "empresa", "unidad_productiva", defecto="")
    municipio = obtener_dato(datos, "municipio", "ciudad", "ubicacion", defecto="")
    trl = obtener_dato(datos, "trl_alcanzado", "trl", defecto="")

    tabla = documento.add_table(rows=5, cols=2)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.style = "Table Grid"

    filas = [
        ("Código del proyecto", codigo),
        ("Nombre del proyecto", nombre),
        ("Beneficiario / unidad atendida", beneficiario),
        ("Municipio / ubicación", municipio),
        ("TRL alcanzado al cierre", trl),
    ]

    for idx, (campo, valor) in enumerate(filas):
        set_cell_text(tabla.cell(idx, 0), campo, bold=True)
        set_cell_text(tabla.cell(idx, 1), valor)


def agregar_objetivos(documento: DocumentObject, datos: Dict[str, Any]) -> None:
    objetivo_general = corregir_ortografia_basica(obtener_dato(datos, "objetivo_general", defecto=""))
    objetivos_especificos = obtener_lista(datos, "objetivos_especificos", "objetivos")

    agregar_titulo(documento, "Objetivo general", 2)
    agregar_parrafo(documento, objetivo_general or "No se registró objetivo general.")

    agregar_titulo(documento, "Objetivos específicos", 2)
    if objetivos_especificos:
        for objetivo in objetivos_especificos:
            p = documento.add_paragraph(style=None)
            p.style = "List Bullet"
            p.add_run(capitalizar_oracion(corregir_ortografia_basica(objetivo).strip(" .")) + ".")
    else:
        agregar_parrafo(documento, "No se registraron objetivos específicos.")


def agregar_tabla_resultados_actividades(documento: DocumentObject, actividades: List[str]) -> None:
    """
    Tabla solicitada:
    N.° | Descripción de la actividad ejecutada | Evidencia
    La evidencia queda como espacio para agregar enlace posteriormente en Word.
    """
    if not actividades:
        actividades = ["Actividad principal ejecutada durante el desarrollo del proyecto."]

    tabla = documento.add_table(rows=1, cols=3)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.style = "Table Grid"

    encabezados = ["N.°", "Descripción de la actividad ejecutada", "Evidencia"]
    for i, encabezado in enumerate(encabezados):
        set_cell_text(tabla.cell(0, i), encabezado, bold=True)

    for idx, actividad in enumerate(actividades, start=1):
        row = tabla.add_row()
        set_cell_text(row.cells[0], str(idx))
        set_cell_text(row.cells[1], actividad)
        set_cell_text(row.cells[2], "Espacio para agregar enlace")

    # Ajuste aproximado de anchos
    for row in tabla.rows:
        row.cells[0].width = Cm(1.2)
        row.cells[1].width = Cm(12)
        row.cells[2].width = Cm(4)


def agregar_referencias(documento: DocumentObject) -> None:
    for referencia in referencias_estado_del_arte():
        p = documento.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.left_indent = Cm(0.5)
        run = p.add_run(referencia)
        run.font.name = "Arial"
        run.font.size = Pt(10)


def configurar_propiedades_documento(documento: DocumentObject, datos: Dict[str, Any]) -> str:
    """
    Configura las propiedades internas del Word usando codigo_proyecto, no nombre_proyecto.
    Retorna el código limpio para usarlo también en el nombre del archivo.
    """
    codigo_proyecto = (
        obtener_dato(datos, "codigo_proyecto", "codigo", "codigo_tecnoparque", "codigo_sennova", "id_proyecto", defecto="")
        or "SIN_CODIGO"
    )
    codigo_proyecto = limpiar_nombre_archivo(codigo_proyecto)

    documento.core_properties.title = limpiar_texto_propiedad_word(f"Informe Final - {codigo_proyecto}")
    documento.core_properties.subject = "Informe Final Tecnoparque"
    documento.core_properties.author = "Tecnoparque"
    documento.core_properties.keywords = "SENA, TecnoParque, Informe Final, GCDTP-F-023 V01"

    return codigo_proyecto


# =========================================================
# GENERADOR PRINCIPAL
# =========================================================

def generar_documento_informe_final(datos: Dict[str, Any], ruta_salida: Optional[Path | str] = None) -> Tuple[bytes, str]:
    """
    Genera el documento Word del Informe Final.

    Parámetros:
        datos: diccionario con la información del formulario.
        ruta_salida: ruta opcional para guardar el archivo.

    Retorna:
        (bytes_docx, nombre_archivo)
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    documento = cargar_documento_base()
    configurar_actualizacion_campos(documento)

    codigo_proyecto = configurar_propiedades_documento(documento, datos)
    nombre_archivo = f"Informe Final - {codigo_proyecto}.docx"

    actividades_originales = obtener_lista(datos, "actividades_ejecutadas", "actividades", "descripciones_actividades")
    actividades_mejoradas = mejorar_lista_actividades(actividades_originales)

    # Inicio del contenido generado.
    documento.add_page_break()

    agregar_titulo(documento, "Informe Final", 1)
    agregar_parrafo(documento, f"Formato institucional {CODIGO_FORMATO_INFORME}.")

    agregar_titulo(documento, "Identificación del proyecto", 1)
    agregar_tabla_identificacion(documento, datos)

    documento.add_page_break()
    agregar_tabla_contenido(documento)

    agregar_titulo(documento, "1. Descripción general del proyecto", 1)
    agregar_parrafo(documento, redactar_descripcion_general(datos))

    agregar_titulo(documento, "2. Objetivos", 1)
    agregar_objetivos(documento, datos)

    agregar_titulo(documento, "3. Metodología y actividades ejecutadas", 1)
    agregar_parrafo(documento, redactar_metodologia_y_desarrollo(datos, actividades_mejoradas))

    agregar_titulo(documento, "4. Estado del arte", 1)
    agregar_parrafo(documento, redactar_estado_del_arte(datos))

    agregar_titulo(documento, "5. Resultados", 1)
    agregar_parrafo(documento, redactar_resultados_generales(datos, actividades_mejoradas))
    agregar_titulo(documento, "5.1 Actividades ejecutadas y evidencias", 2)
    agregar_tabla_resultados_actividades(documento, actividades_mejoradas)

    agregar_titulo(documento, "6. Propiedad intelectual y transferencia tecnológica", 1)
    agregar_parrafo(documento, redactar_propiedad_intelectual(datos, actividades_mejoradas))

    agregar_titulo(documento, "7. Conclusiones y recomendación de continuidad", 1)
    agregar_parrafo(documento, redactar_conclusiones(datos, actividades_mejoradas))

    agregar_titulo(documento, "8. Referencias", 1)
    agregar_referencias(documento)

    agregar_titulo(documento, "9. Anexos y evidencias", 1)
    agregar_parrafo(
        documento,
        "Este espacio queda disponible para que el usuario inserte manualmente enlaces, imágenes, "
        "capturas, soportes o evidencias adicionales directamente en Word. El módulo no solicita ni "
        "carga archivos de evidencias."
    )

    buffer = io.BytesIO()
    documento.save(buffer)
    bytes_docx = buffer.getvalue()

    if ruta_salida is None:
        ruta_final = OUTPUT_DIR / nombre_archivo
    else:
        ruta_final = Path(ruta_salida)

    ruta_final.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta_final, "wb") as f:
        f.write(bytes_docx)

    return bytes_docx, nombre_archivo


# Alias para compatibilidad con nombres que pueden existir en el proyecto.
generar_informe_final = generar_documento_informe_final
generar_word_informe_final = generar_documento_informe_final


# =========================================================
# FORMULARIO STREAMLIT
# =========================================================

def render_formulario_informe_final(datos_base: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
    """
    Renderiza un formulario Streamlit sencillo.
    No incluye carga de archivos de evidencias.
    Las actividades se solicitan únicamente como descripción, una por línea.
    """
    if st is None:
        raise RuntimeError("Streamlit no está disponible en este entorno.")

    datos_base = datos_base or {}

    st.subheader("Informe Final GCDTP-F-023 V01")

    with st.form("form_informe_final_gcdtp_023"):
        codigo_proyecto = st.text_input(
            "Código del proyecto",
            value=datos_base.get("codigo_proyecto", ""),
            help="Este código será usado para nombrar el archivo: Informe Final - Código del proyecto.docx",
        )

        nombre_proyecto = st.text_area(
            "Nombre del proyecto",
            value=datos_base.get("nombre_proyecto", ""),
            height=90,
        )

        beneficiario = st.text_input(
            "Beneficiario / unidad atendida",
            value=datos_base.get("beneficiario", ""),
        )

        municipio = st.text_input(
            "Municipio / ubicación",
            value=datos_base.get("municipio", ""),
        )

        sector = st.text_input(
            "Sector o área del proyecto",
            value=datos_base.get("sector", ""),
        )

        descripcion_proyecto = st.text_area(
            "Descripción general del proyecto",
            value=datos_base.get("descripcion_proyecto", ""),
            height=180,
        )

        objetivo_general = st.text_area(
            "Objetivo general",
            value=datos_base.get("objetivo_general", ""),
            height=100,
        )

        objetivos_especificos = st.text_area(
            "Objetivos específicos",
            value=datos_base.get("objetivos_especificos", ""),
            height=130,
            help="Escribe un objetivo por línea.",
        )

        metodologia_utilizada = st.selectbox(
            "Metodología utilizada",
            [
                "Design Thinking",
                "Metodología ágil e iterativa",
                "Investigación aplicada y desarrollo experimental",
                "Prototipado y validación funcional",
                "Diseño centrado en el usuario",
                "Otra",
            ],
            index=0,
        )

        metodologia_otro = ""
        if metodologia_utilizada == "Otra":
            metodologia_otro = st.text_input("Especifique la metodología utilizada")

        actividades = st.text_area(
            "Actividades ejecutadas durante el proyecto",
            value=datos_base.get("actividades_ejecutadas", ""),
            height=180,
            help="Escribe una actividad por línea. Solo se solicita la descripción de la actividad.",
        )

        resultado_general = st.text_area(
            "Resultado general alcanzado",
            value=datos_base.get("resultado_general", ""),
            height=120,
        )

        trl_alcanzado = st.selectbox(
            "TRL alcanzado al cierre",
            ["TRL 1", "TRL 2", "TRL 3", "TRL 4", "TRL 5", "TRL 6", "TRL 7", "TRL 8", "TRL 9"],
            index=5,
        )

        generar = st.form_submit_button("Generar Word oficial GCDTP-F-023 V01")

    if not generar:
        return None

    metodologia_final = metodologia_otro.strip() if metodologia_utilizada == "Otra" and metodologia_otro.strip() else metodologia_utilizada

    return {
        "codigo_proyecto": codigo_proyecto,
        "nombre_proyecto": nombre_proyecto,
        "beneficiario": beneficiario,
        "municipio": municipio,
        "sector": sector,
        "descripcion_proyecto": descripcion_proyecto,
        "objetivo_general": objetivo_general,
        "objetivos_especificos": objetivos_especificos,
        "metodologia_utilizada": metodologia_final,
        "actividades_ejecutadas": actividades,
        "resultado_general": resultado_general,
        "trl_alcanzado": trl_alcanzado,
    }


def render_modulo_informe_final(datos_base: Optional[Dict[str, Any]] = None) -> None:
    """
    Función lista para llamar desde app.py o desde el router del módulo de cierre.
    """
    if st is None:
        raise RuntimeError("Streamlit no está disponible en este entorno.")

    datos = render_formulario_informe_final(datos_base)

    if datos:
        try:
            bytes_docx, nombre_archivo = generar_documento_informe_final(datos)

            st.success("Documento Word generado correctamente.")

            st.download_button(
                label="Descargar Word oficial GCDTP-F-023 V01",
                data=bytes_docx,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            st.info(
                "Nota: al abrir el documento en Word, actualiza la tabla de contenido "
                "para que los números de página coincidan con la ubicación real de cada apartado."
            )

        except Exception as exc:
            st.error(f"No se pudo generar el documento Word: {exc}")


# Alias frecuentes para integración con routers existentes.
mostrar_modulo_informe_final = render_modulo_informe_final
render = render_modulo_informe_final
main = render_modulo_informe_final


if __name__ == "__main__":
    # Prueba local simple sin Streamlit.
    datos_prueba = {
        "codigo_proyecto": "P2026-000001",
        "nombre_proyecto": "Diseño e implementación de un sistema web interactivo para el avistamiento de aves",
        "beneficiario": "La Morada del Viento",
        "municipio": "Santa María, Huila",
        "sector": "Turismo de naturaleza y educación ambiental",
        "descripcion_proyecto": "Sistema web interactivo con códigos QR, placas informativas, contenidos científicos, registros sonoros, mapas de avistamiento y fuentes validadas.",
        "objetivo_general": "Implementar un sistema web interactivo para fortalecer la experiencia de avistamiento de aves.",
        "objetivos_especificos": "Diseñar el sistema web interactivo.\nImplementar códigos QR y placas informativas.\nIntegrar contenidos científicos y registros sonoros.",
        "metodologia_utilizada": "Design Thinking",
        "actividades_ejecutadas": "Levantamiento de información del sendero.\nDiseño de páginas web interactivas.\nGeneración de códigos QR.\nImplementación en hosting y dominio.\nValidación funcional del sistema.",
        "resultado_general": "Se implementó una solución digital funcional para consulta de información de aves.",
        "trl_alcanzado": "TRL 6",
    }

    generar_documento_informe_final(datos_prueba)
# =========================================================
# ALIAS REQUERIDO POR app.py
# =========================================================

def render_informe_tecnico_final(*args, **kwargs):
    if "render_modulo_informe_final" in globals():
        return render_modulo_informe_final(*args, **kwargs)

    if "mostrar_modulo_informe_final" in globals():
        return mostrar_modulo_informe_final(*args, **kwargs)

    if "main" in globals():
        return main(*args, **kwargs)

    raise RuntimeError(
        "No se encontró una función de renderizado para el módulo de Informe Técnico Final."
    )
