def safe_filename(texto: str) -> str:
    limpio = "".join(
        ch if ch.isalnum() or ch in "-_ " else ""
        for ch in str(texto)
    )

    limpio = "_".join(limpio.split())

    return limpio[:80] or "documento"