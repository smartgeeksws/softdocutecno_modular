import streamlit as st


def inicializar_estado() -> None:
    if "fase_seleccionada" not in st.session_state:
        st.session_state.fase_seleccionada = None

    if "documento_seleccionado" not in st.session_state:
        st.session_state.documento_seleccionado = None


def seleccionar_fase(nombre_fase: str) -> None:
    st.session_state.fase_seleccionada = nombre_fase
    st.session_state.documento_seleccionado = None


def seleccionar_documento(nombre_documento: str) -> None:
    st.session_state.documento_seleccionado = nombre_documento


def render_main_menu() -> None:
    st.subheader("¿Qué deseas hacer?")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("📌 Fase de inicio"):
            seleccionar_fase("inicio")

    with col2:
        if st.button("🗓️ Fase de planeación"):
            seleccionar_fase("planeacion")

    with col3:
        if st.button("📝 Acta de ejecución"):
            seleccionar_fase("ejecucion")

    with col4:
        if st.button("✅ Documentos de cierre"):
            seleccionar_fase("cierre")


def render_menu_inicio() -> None:
    st.markdown("---")
    st.subheader("Documentos de la fase de inicio")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("📄 Acta de inicio"):
            seleccionar_documento("acta_inicio")

    with col2:
        if st.button("🏢 Uso de infraestructura"):
            seleccionar_documento("uso_infraestructura")

    with col3:
        if st.button("🔒 Confidencialidad"):
            seleccionar_documento("confidencialidad")


def render_menu_planeacion() -> None:
    st.markdown("---")
    st.subheader("Documentos de la fase de planeación")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("🗓️ Cronograma de actividades"):
            seleccionar_documento("cronograma")

    with col2:
        if st.button("📚 Estado del arte"):
            seleccionar_documento("estado_arte")


def render_menu_cierre() -> None:
    st.markdown("---")
    st.subheader("Documentos de la fase de cierre")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("📄 Acta de cierre y ficha de caracterización"):
            seleccionar_documento("acta_cierre_ficha")

    with col2:
        if st.button("🧩 Modelo de negocio Lean Canvas"):
            seleccionar_documento("modelo_negocio")

    with col3:
        if st.button("📘 Informe técnico final"):
            seleccionar_documento("informe_tecnico_final")