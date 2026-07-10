import streamlit as st


def render_header() -> None:
    st.markdown(
        '<div class="main-title">Generador de Documentos Tecnoparque</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="subtitle">Sistema modular para generar documentos institucionales de proyectos de base tecnológica, innovación y desarrollo tecnológico.</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-box">
        Selecciona la fase del proceso documental y luego el documento que deseas generar.
        </div>
        """,
        unsafe_allow_html=True,
    )