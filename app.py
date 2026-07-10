import streamlit as st

from ui.styles import cargar_estilos
from ui.layout import render_header
from ui.sidebar import render_sidebar
from ui.navigation import (
    inicializar_estado,
    render_main_menu,
    render_menu_inicio,
    render_menu_planeacion,
    render_menu_cierre,
)

from modules.cierre.modelo_negocio import render_modelo_negocio


def main() -> None:
    st.set_page_config(
        page_title="Generador de Documentos Tecnoparque",
        page_icon="📄",
        layout="wide",
    )

    inicializar_estado()
    cargar_estilos()

    modo_prueba, modelo_openai = render_sidebar()

    render_header()
    render_main_menu()

    fase = st.session_state.fase_seleccionada
    documento = st.session_state.documento_seleccionado

    if fase is None:
        st.info("Selecciona una fase para iniciar.")
        st.stop()

    if fase == "inicio":
        render_menu_inicio()

        if documento == "acta_inicio":
            st.warning("Módulo pendiente de migrar: Acta de inicio.")
        elif documento == "uso_infraestructura":
            st.warning("Módulo pendiente de migrar: Uso de infraestructura.")
        elif documento == "confidencialidad":
            st.warning("Módulo pendiente de migrar: Confidencialidad.")
        else:
            st.info("Selecciona un documento de inicio.")

    elif fase == "planeacion":
        render_menu_planeacion()

        if documento == "cronograma":
            st.warning("Módulo pendiente de migrar: Cronograma.")
        elif documento == "estado_arte":
            st.warning("Módulo pendiente de migrar: Estado del arte.")
        else:
            st.info("Selecciona un documento de planeación.")

    elif fase == "ejecucion":
        st.markdown("---")
        st.subheader("Acta de ejecución")
        st.warning("Módulo pendiente de migrar: Acta de ejecución.")

    elif fase == "cierre":
        render_menu_cierre()

        if documento == "modelo_negocio":
            render_modelo_negocio(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )
        elif documento == "acta_cierre_ficha":
            st.warning("Módulo pendiente de migrar: Acta de cierre y ficha de caracterización.")
        elif documento == "informe_tecnico_final":
            st.warning("Módulo pendiente de migrar: Informe técnico final.")
        else:
            st.info("Selecciona un documento de cierre.")


if __name__ == "__main__":
    main()