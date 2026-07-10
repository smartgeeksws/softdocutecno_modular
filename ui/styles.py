import streamlit as st


def cargar_estilos() -> None:
    st.markdown(
        """
        <style>
        .main-title {
            font-size: 2rem;
            font-weight: 700;
            color: #2e7d32;
            margin-bottom: 0.2rem;
        }

        .subtitle {
            font-size: 1rem;
            color: #555;
            margin-bottom: 1.5rem;
        }

        .info-box {
            background-color: #f4f9f4;
            border-left: 5px solid #39a935;
            padding: 1rem;
            border-radius: 0.5rem;
            margin-bottom: 1rem;
        }

        div.stButton > button {
            width: 100%;
            min-height: 3.2rem;
            border-radius: 0.7rem;
            font-weight: 600;
            border: 1px solid #39a935;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )