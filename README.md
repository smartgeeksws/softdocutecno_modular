# SoftDocuTecno Modular

Sistema modular para la generación de documentos institucionales de proyectos Tecnoparque.

## Objetivo

Organizar el software de generación documental en una arquitectura clara, mantenible y escalable, evitando concentrar toda la lógica en un solo archivo.

## Estructura general

- `app.py`: archivo principal de Streamlit.
- `config/`: configuración, rutas y constantes.
- `ui/`: interfaz, estilos, navegación y barra lateral.
- `modules/`: documentos organizados por fase.
- `services/`: servicios reutilizables como OpenAI, PDF, Word y JSON.
- `utils/`: funciones auxiliares.
- `resources/`: logos, firmas e imágenes.
- `output/`: documentos generados localmente.
- `data/`: respaldos temporales en JSON.

## Ejecutar localmente

```bash
streamlit run app.py