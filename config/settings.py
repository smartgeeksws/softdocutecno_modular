import os
import streamlit as st


def obtener_api_key() -> str | None:
    api_key = os.getenv("OPENAI_API_KEY")

    if api_key:
        return api_key

    try:
        return st.secrets.get("OPENAI_API_KEY")
    except Exception:
        return None