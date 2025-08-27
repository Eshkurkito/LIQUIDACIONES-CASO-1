import streamlit as st
import pandas as pd
import numpy as np
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment

st.set_page_config(page_title="LIQUIDACIONES (Casos 1-5)", page_icon="📊", layout="wide")

# =====================
# Utilidades
# =====================
def ensure_required(df, required, ctx=""):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas requeridas: {missing} en {ctx}. Ajusta cabecera o activa modo por letras.")
        st.stop()

def _first_existing(df, candidates):
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = str(cand).strip().lower()
        if k in norm_map:
            return norm_map[k]
    return None

# =====================
# Normalización
# =====================
def normalize_columns(df):
    out = df.copy()
    col_aloj = _first_existing(out, ["Nombre alojamiento","Alojamiento"])
    col_fent = _first_existing(out, ["Fecha entrada","Fecha de entrada"])
    col_fsal = _first_existing(out, ["Fecha salida","Fecha de salida"])
    col_noch = _first_existing(out, ["Noches","Noches ocupadas"])
    col_alq  = _first_existing(out, ["Alquiler con tasas","Ingreso alojamiento"])
    col_ext  = _first_existing(out, ["Extras con tasas","Ingreso limpieza","Limpieza"])
    col_tot  = _first_existing(out, ["Total reserva con tasas","Total ingresos"])
    col_port = _first_existing(out, ["Web origen","Portal"])
    col_comi = _first_existing(out, ["Comisión portal","Comisión Portal/Intermediario: Comisión calculada"])
    col_ivaal= _first_existing(out, ["IVA del alojamiento"])

    rename = {}
    if col_aloj: rename[col_aloj] = "Alojamiento"
    if col_fent: rename[col_fent] = "Fecha entrada"
    if col_fsal: rename[col_fsal] = "Fecha salida"
    if col_noch: rename[col_noch] = "Noches ocupadas"
    if col_alq:  rename[col_alq]  = "Ingreso alojamiento"
    if col_ext:  rename[col_ext]  = "Ingreso limpieza"
    if col_tot:  rename[col_tot]  = "Total ingresos"
    if col_port: rename[col_port] = "Portal"
    if col_comi: rename[col_comi] = "Comisión portal"
    if col_ivaal:rename[col_ivaal]= "IVA del alquiler"

    out.rename(columns=rename, inplace=True)

    # Tipado
    for c in ["Ingreso alojamiento","Ingreso limpieza","Total ingresos","Comisión portal","IVA del alquiler","Noches ocupadas"]:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)

    for c in ["Fecha entrada","Fecha salida"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    if "Alojamiento" in out.columns:
        out["Alojamiento"] = out["Alojamiento"].astype(str).str.strip().str.upper()

    return out

# =====================
# App principal
# =====================
st.title("📊 LIQUIDACIONES Automáticas (Casos 1–5)")
st.caption("Sube Excel de Avantio (.xlsx) • Selección de caso • Filtro por alojamiento • Período y prorrateo")

with st.sidebar:
    st.header("Período")
    col1, col2 = st.columns(2)
    with col1: start_date = st.date_input("Desde", value=date(date.today().year, date.today().month, 1))
    with col2: end_date   = st.date_input("Hasta",  value=date(date.today().year, date.today().month, 28))
    st.divider()
    case_choice = st.radio("Selecciona el caso", [1,2,3,4,5], horizontal=True)

file = st.file_uploader("Sube el archivo de reservas (.xlsx)", type=["xlsx"])

if file is not None:
    df_in = pd.read_excel(file, header=0)
    if df_in.columns.duplicated().any():
        st.warning("Había columnas duplicadas. Me quedo con la primera de cada nombre.")
        df_in = df_in.loc[:, ~df_in.columns.duplicated()]

    st.subheader("Vista previa (todas las filas)")
    st.dataframe(df_in, use_container_width=True)

    df_norm = normalize_columns(df_in)

    # ---- Filtro por Alojamiento ----
    all_props = sorted(df_norm["Alojamiento"].dropna().unique().tolist()) if "Alojamiento" in df_norm.columns else []
    if all_props:
        st.subheader("Filtro por alojamiento")
        selected_props = st.multiselect("Selecciona uno o varios alojamientos", all_props, default=all_props)
        df_norm = df_norm[df_norm["Alojamiento"].isin(selected_props)]

    if st.button("Generar liquidación"):
        st.success(f"Liquidación generada (Caso {case_choice}) • {start_date}–{end_date}")
        st.dataframe(df_norm, use_container_width=True)
else:
    st.info("Sube un archivo para comenzar.")
