import json

from config.constants import DATA_DIR


def guardar_datos_json(datos: dict, nombre_archivo: str) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    ruta = DATA_DIR / nombre_archivo

    ruta.write_text(
        json.dumps(datos, ensure_ascii=False, indent=4, default=str),
        encoding="utf-8",
    )