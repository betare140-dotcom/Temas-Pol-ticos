# Extractor de menciones Onclusive por tema

Aplicación Streamlit independiente del sistema de actores políticos.

## Archivos que debes subir al repositorio nuevo

- `app.py` — aplicación principal.
- `extractor_onclusive.py` — lógica de extracción y generación de Word/HTML/TXT.
- `requirements.txt` — dependencias.

## Despliegue

1. Crea un repositorio nuevo en GitHub, por ejemplo `extractor-temas-onclusive`.
2. Sube `app.py`, `extractor_onclusive.py` y `requirements.txt` en la raíz.
3. En Streamlit Community Cloud crea una app nueva apuntando a ese repositorio y selecciona `app.py` como archivo principal.
4. No necesitas `GEMINI_API_KEY` para esta aplicación.

## Flujo

1. Escribe el nombre del tema.
2. Sube el Excel o CSV de Onclusive.
3. Opcionalmente activa el intento de recuperación de usuarios faltantes de Instagram/TikTok.
4. Presiona **Extraer menciones**.
5. Descarga Word, HTML o TXT.

La aplicación procesa X, Facebook, Instagram y TikTok, elimina RT y duplicados, limpia contenido y organiza las menciones por fecha.
