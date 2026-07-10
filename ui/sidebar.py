import streamlit as st

from config.constants import MODELOS_OPENAI
from config.settings import obtener_api_key


def render_sidebar() -> tuple[bool, str]:
    with st.sidebar:
        st.header("Configuración")

        modo_prueba = st.checkbox(
            "Activar modo prueba sin consumir API",
            value=False,
            help="Usa textos generados localmente para probar el flujo sin gastar saldo de OpenAI.",
        )

        modelo_openai = st.selectbox(
            "Modelo para generar textos",
            MODELOS_OPENAI,
            index=0,
        )

        if modo_prueba:
            st.info("Modo prueba activo: no se consumirá API")
        else:
            if obtener_api_key():
                st.success("API Key detectada")
            else:
                st.warning("No se detectó OPENAI_API_KEY")

        st.markdown("---")
        st.caption("Recursos: resources/")
        st.caption("Firmas institucionales: resources/firmas/")

    return modo_prueba, modelo_openai