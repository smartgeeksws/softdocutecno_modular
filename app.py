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
from modules.cierre.acta_cierre import render_acta_cierre
from modules.cierre.informe_tecnico_final import (
    render_informe_tecnico_final,
)

from modules.inicio.acta_inicio import render_acta_inicio
from modules.inicio.confidencialidad import render_confidencialidad
from modules.inicio.uso_infraestructura import render_uso_infraestructura

from modules.planeacion.cronograma import render_cronograma
from modules.planeacion.estado_arte import render_estado_arte

from modules.ejecucion.acta_ejecucion import render_acta_ejecucion


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
            render_acta_inicio(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        elif documento == "uso_infraestructura":
            render_uso_infraestructura(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        elif documento == "confidencialidad":
            render_confidencialidad(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        else:
            st.info("Selecciona un documento de inicio.")

    elif fase == "planeacion":
        render_menu_planeacion()

        if documento == "cronograma":
            render_cronograma(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        elif documento == "estado_arte":
            render_estado_arte(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        else:
            st.info("Selecciona un documento de planeación.")

    elif fase == "ejecucion":
        render_acta_ejecucion(
            modo_prueba=modo_prueba,
            modelo_openai=modelo_openai,
        )

    elif fase == "cierre":
        render_menu_cierre()

        if documento == "modelo_negocio":
            render_modelo_negocio(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        elif documento == "acta_cierre_ficha":
            render_acta_cierre(
                modo_prueba=modo_prueba,
                modelo_openai=modelo_openai,
            )

        elif documento == "informe_tecnico_final":
                render_informe_tecnico_final(
                    modo_prueba=modo_prueba,
                    modelo_openai=modelo_openai,
                )

        else:
            st.info("Selecciona un documento de cierre.")


if __name__ == "__main__":
    main()