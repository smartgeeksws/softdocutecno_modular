from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

RESOURCES_DIR = BASE_DIR / "resources"
OUTPUT_DIR = BASE_DIR / "output"
DATA_DIR = BASE_DIR / "data"

RUTA_LOGO_SENA = RESOURCES_DIR / "logo_sena.png"
RUTA_LOGO_TECNOPARQUE = RESOURCES_DIR / "logo_tecnoparque.png"
CARPETA_FIRMAS = RESOURCES_DIR / "firmas"

FORMATO_CODIGO = "GOR-F-084 V02"
FORMATO_CONFIDENCIALIDAD = "GIC-F-041 V03"
FORMATO_ACTA_CIERRE = "GOR-F-084 V02"

LUGAR_ENLACE_DEFAULT = "Tecnoparque Angostura - km 38 vía al sur de Neiva"

DIRECCION_REGIONAL_CENTRO_DEFAULT = (
    "Dirección de formación profesional / HUILA / Centro De Formación Agroindustrial"
)

DEPENDENCIA_TALENTO_DEFAULT = "EMPRENDEDOR"
DEPENDENCIA_EXPERTO_DEFAULT = "SENA"
ANEXOS_DEFAULT = "NO APLICA"

NIVELES_TRL = ["TRL 6", "TRL 7", "TRL 8"]

MODELOS_OPENAI = [
    "gpt-4.1-mini",
    "gpt-4.1",
]