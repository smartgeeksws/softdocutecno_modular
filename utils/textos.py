import re


def limpiar_texto(texto: str) -> str:
    texto = str(texto or "")
    texto = texto.replace("…", "")
    texto = texto.replace("...", "")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def limitar_palabras(texto: str, max_palabras: int = 220) -> str:
    palabras = limpiar_texto(texto).split()

    if len(palabras) <= max_palabras:
        return " ".join(palabras)

    return " ".join(palabras[:max_palabras]).rstrip(".,;:") + "."