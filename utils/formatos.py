def formato_moneda_colombiana(valor: float | int) -> str:
    try:
        valor_int = int(round(float(valor)))
    except Exception:
        valor_int = 0

    return "$" + f"{valor_int:,.0f}".replace(",", ".")


def conteo_palabras(texto: str) -> int:
    return len(str(texto or "").split())