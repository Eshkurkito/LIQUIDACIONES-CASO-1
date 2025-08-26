import streamlit as st
import pandas as pd

st.set_page_config(page_title="LIQUIDACIONES (Casos 1-5)", page_icon="📊", layout="wide")

st.title("📊 LIQUIDACIONES Automáticas (Casos 1–5)")
st.caption("Sube un Excel de Avantio (.xlsx), detecta el caso, aplica reglas oficiales y filtra por período.")

file = st.file_uploader("Sube el archivo de reservas (.xlsx)", type=["xlsx"])

if file is not None:
    try:
        df = pd.read_excel(file, header=1)
    except Exception:
        df = pd.read_excel(file)
    st.write("Vista previa:")
    st.dataframe(df.head(10), use_container_width=True)
    st.success("✅ Archivo cargado correctamente. Aquí iría la lógica de cálculo de liquidación.")
