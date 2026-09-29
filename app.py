import streamlit as st
from extractor_onclusive import render_extractor_onclusive

st.set_page_config(
    page_title="Extractor de menciones Onclusive",
    page_icon="🔎",
    layout="centered",
)

st.title("🔎 Extracción de menciones Onclusive")
st.write(
    "Herramienta exclusiva para extraer y depurar menciones por tema desde "
    "archivos Excel o CSV exportados por Onclusive."
)

render_extractor_onclusive()
