import json
import re

from openai import OpenAI

from config.settings import obtener_api_key


def limpiar_respuesta_json(texto: str) -> str:
    if not texto:
        return ""

    texto = str(texto).strip()
    texto = texto.replace("```json", "")
    texto = texto.replace("```JSON", "")
    texto = texto.replace("```", "")
    texto = texto.strip()

    if texto.startswith("{") and texto.endswith("}"):
        return texto

    inicio = texto.find("{")
    fin = texto.rfind("}")

    if inicio != -1 and fin != -1 and fin > inicio:
        return texto[inicio:fin + 1].strip()

    match = re.search(r"\{.*\}", texto, re.DOTALL)

    if match:
        return match.group(0).strip()

    return texto


def obtener_cliente_openai() -> OpenAI:
    api_key = obtener_api_key()

    if not api_key:
        raise ValueError("No se encontró OPENAI_API_KEY.")

    return OpenAI(api_key=api_key)


def generar_json_openai(
    instrucciones: str,
    entrada: str,
    modelo: str = "gpt-4.1-mini",
    temperature: float = 0.25,
) -> dict:
    client = obtener_cliente_openai()

    respuesta = client.responses.create(
        model=modelo,
        instructions=instrucciones,
        input=entrada,
        temperature=temperature,
    )

    texto = limpiar_respuesta_json(respuesta.output_text)
    return json.loads(texto)


def generar_texto_openai(
    instrucciones: str,
    entrada: str,
    modelo: str = "gpt-4.1-mini",
    temperature: float = 0.25,
) -> str:
    client = obtener_cliente_openai()

    respuesta = client.responses.create(
        model=modelo,
        instructions=instrucciones,
        input=entrada,
        temperature=temperature,
    )

    return str(respuesta.output_text or "").strip()