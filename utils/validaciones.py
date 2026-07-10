import streamlit as st


def validar_campos_obligatorios(campos: dict) -> bool:
    faltantes = []

    for nombre, valor in campos.items():
        if not str(valor or "").strip():
            faltantes.append(nombre)

    if faltantes:
        st.error("Faltan campos obligatorios: " + ", ".join(faltantes))
        return False

    return True