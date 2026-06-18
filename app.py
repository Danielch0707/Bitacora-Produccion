import streamlit as st

st.title("✅ TEST FUNCIONANDO")

archivo = st.file_uploader("Sube archivo")

if archivo is not None:
    st.success("Archivo cargado")
else:
    st.warning("Sube archivo")
