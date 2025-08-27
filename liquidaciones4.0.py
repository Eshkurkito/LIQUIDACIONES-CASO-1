import streamlit as st
import pandas as pd
import numpy as np
from datetime import date

st.set_page_config(page_title="LIQUIDACIONES (Casos 1–5)", page_icon="📊", layout="wide")

# ========= Utilidades =========
def ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garantiza nombres de columna únicos sin eliminar ninguna.
    Si hay duplicadas, añade sufijos .1, .2, ...
    """
    counts = {}
    new_cols = []
    for c in df.columns:
        name = str(c)
        n = counts.get(name, 0)
        new_cols.append(name if n == 0 else f"{name}.{n}")
        counts[name] = n + 1
    out = df.copy()
    out.columns = new_cols
    return out

def _first_existing(df, candidates):
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = str(cand).strip().lower()
        if k in norm_map:
            return norm_map[k]
    return None

def ensure_required(df, required, ctx=""):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas requeridas: {missing} en {ctx}. Ajusta la cabecera o activa el modo por letras.")
        st.stop()

def fmt_es_num(x):
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

def show_table_es(df, title):
    st.subheader(title)
    view = ensure_unique_columns(df)  # blindaje anti-duplicadas
    view = view.copy()
    for c in view.columns:
        if pd.api.types.is_numeric_dtype(view[c]):
            view[c] = view[c].apply(fmt_es_num)
    st.dataframe(view, use_container_width=True)

# ========= Normalización =========
def normalize_columns(df):
    out = df.copy()

    col_aloj = _first_existing(out, ["Nombre alojamiento","Alojamiento","Nombre del alojamiento","Nombre Alojamiento"])
    col_fent = _first_existing(out, ["Fecha entrada","Fecha de entrada"])
    col_fsal = _first_existing(out, ["Fecha salida","Fecha de salida"])
    col_noch = _first_existing(out, ["Noches","noches","Noches ocupadas"])
    col_alq  = _first_existing(out, ["Alquiler con tasas","Ingreso alojamiento","Importe alojamiento"])
    col_ext  = _first_existing(out, ["Extras con tasas","Ingreso limpieza","Limpieza","Importe limpieza"])
    col_tot  = _first_existing(out, ["Total reserva con tasas","Total ingresos","Total"])
    col_port = _first_existing(out, ["Web origen","Portal","Canal","Fuente"])
    col_comi = _first_existing(out, ["Comisión Portal/Intermediario: Comisión calculada","Comisión portal","Comisión"])
    col_ivaal= _first_existing(out, ["IVA del alojamiento","IVA alojamiento"])

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

# --- Fallback por letras (si cambian nombres en Avantio) ---
LETTER_MAP_DEFAULT = {
    "W": "Alojamiento",
    "D": "Fecha entrada",
    "F": "Fecha salida",
    "H": "Noches ocupadas",
    "I": "Ingreso alojamiento",
    "J": "Ingreso limpieza",
    "O": "Total ingresos",
    "AP": "Portal",              # Web origen
    "AR": "Comisión portal",
    "AL": "IVA del alquiler",    # si existe
}
def letters_to_idx(letter):
    s = letter.upper()
    n = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'): return None
        n = n*26 + (ord(ch)-ord('A')+1)
    return n-1
def normalize_columns_by_letters(df, letter_map=LETTER_MAP_DEFAULT):
    out = df.copy()
    cols = list(out.columns)
    rename = {}
    for L, std in letter_map.items():
        i = letters_to_idx(L)
        if i is not None and i < len(cols):
            rename[cols[i]] = std
    out.rename(columns=rename, inplace=True)
    return normalize_columns(out)

# ========= Reglas de casos =========
# Caso 1
case1_percent_amenities = {
    "APOLO 180": (0.20, 12.04), "ALMIRANTE 01": (0.22, 11.33), "ALMIRANTE 02": (0.22, 11.33),
    "CADIZ": (0.20, 9.11), "DENIA 61": (0.20, 10.96), "DOLORES ALCAYDE 04": (0.20, 11.33),
    "DR.LLUCH": (0.20, 11.16), "ERUDITO": (0.20, 13.37), "GOZALBO": (0.20, 15.25),
    "LA ELIANA": (0.20, 15.25), "MORAIRA": (0.25, 11.33), "NAPOLES Y SICILIA": (0.25, 0.00),
    "OLIVERETA 5": (0.20, 0.00), "OVE 01": (0.18, 0.00), "OVE 02": (0.18, 0.00),
    "QUART I": (0.20, 9.09), "QUART II": (0.20, 9.09), "SAN LUIS": (0.20, 11.02),
    "SERRANOS": (0.20, 13.37), "SEVILLA": (0.18, 9.45), "TUNDIDORES": (0.20, 7.85),
    "VALLE": (0.20, 11.33),
}
case1_props = set(case1_percent_amenities.keys())

# Caso 2
case2_percent_amenities = {
    "VISITACION": (0.20, 14.88),
    "PADRE PORTA 6": (0.20, 12.09), "PADRE PORTA 7": (0.20, 12.09), "PADRE PORTA 8": (0.20, 12.09),
    "PADRE PORTA 9": (0.20, 12.09), "PADRE PORTA 10": (0.20, 12.09),
    "LLADRO Y MALLI 00": (0.20, 9.45), "LLADRO Y MALLI 01": (0.20, 9.45), "LLADRO Y MALLI 02": (0.20, 9.45),
    "LLADRO Y MALLI 03": (0.20, 9.45), "LLADRO Y MALLI 04": (0.20, 9.45),
    "APOLO 29": (0.20, 11.58), "APOLO 197": (0.20, 17.40),
}
case2_props = set(case2_percent_amenities.keys())

# Caso 3
case3_cleaning_amenities = {
    "ZAPATEROS 10-2": (0.20, 60.00, 15.24),
    "ZAPATEROS 10-6": (0.20, 75.00, 15.24),
    "ZAPATEROS 10-8": (0.20, 75.00, 15.24),
    "ZAPATEROS 12-5": (0.20, 60.00, 11.33),
    "ALFARO": (0.20, 80.00, 14.88),
}
case3_props = set(case3_cleaning_amenities.keys())

# Caso 4
case4_props = {
    "SERRERIA 04", "SERRERIA 05", "RETOR A", "RETOR B",
    "PASAJE ANGELES Y FEDERICO 01", "PASAJE ANGELES Y FEDERICO 02", "PASAJE ANGELES Y FEDERICO 03",
    "MALILLA 05", "MALILLA 06", "MALILLA 07", "MALILLA 08", "MALILLA 14", "MALILLA 15",
    "BENICALAP 01", "BENICALAP 02", "BENICALAP 03", "BENICALAP 04", "BENICALAP 05", "BENICALAP 06"
}

# Caso 5
case5_percent_amenities = {
    "HOMERO 01": (0.20, 0.00), "HOMERO 02": (0.20, 0.00),
    "CARCAIXENT 01": (0.20, 8.60), "CARCAIXENT 02": (0.20, 8.60)
}
case5_props = set(case5_percent_amenities.keys())

def props_for_case(case):
    if case == 1: return case1_props
    if case == 2: return case2_props
    if case == 3: return case3_props
    if case == 4: return case4_props
    if case == 5: return case5_props
    return set()

# ========= Reglas transversales =========
def apply_booking_vat_on_commission(df, portal_col="Portal", commission_col="Comisión portal"):
    out = df.copy()
    if portal_col in out.columns and commission_col in out.columns:
        ser = out[portal_col]
        if isinstance(ser, pd.DataFrame): ser = ser.iloc[:, 0]
        ser = ser.astype("string").fillna("")
        mask = ser.str.lower().str.contains("booking", na=False)
        out[commission_col] = pd.to_numeric(out[commission_col], errors="coerce").fillna(0.0)
        out.loc[mask, commission_col] = out.loc[mask, commission_col] * 1.21
    return out

def overlap_nights(ci, co, start, end):
    if pd.isna(ci) or pd.isna(co): return 0
    start = pd.to_datetime(start); end = pd.to_datetime(end) + pd.Timedelta(days=1)
    a = max(ci, start); b = min(co, end)
    return max(0, (b-a).days)

def apply_period(df, start_date, end_date, prorate=True, limpieza_mode="prorratear", amenities_mode="prorratear"):
    out = df.copy()
    fe = pd.to_datetime(out.get("Fecha entrada"), errors="coerce", dayfirst=True)
    fs = pd.to_datetime(out.get("Fecha salida"), errors="coerce", dayfirst=True)
    tot_noches = pd.to_numeric(out.get("Noches ocupadas"), errors="coerce").fillna(0)
    nights = [overlap_nights(a,b,start_date,end_date) for a,b in zip(fe,fs)]
    out["Noches periodo"] = pd.Series(nights, index=out.index).astype(float)

    if not prorate:
        mask = out["Noches periodo"] > 0
        out = out[mask].drop(columns=["Noches periodo"])
        return out

    ratio = np.where(tot_noches.to_numpy()>0, (out["Noches periodo"]/tot_noches).to_numpy(), 0.0)

    for col in ["Ingreso alojamiento","Total ingresos","Comisión portal","Honorarios Florit"]:
        if col in out.columns:
            out[col] = (pd.to_numeric(out[col], errors="coerce").fillna(0).to_numpy()*ratio).round(2)

    # Limpieza
    if {"Ingreso limpieza","Gasto limpieza"}.issubset(out.columns):
        if limpieza_mode == "prorratear":
            out["Ingreso limpieza"] = (pd.to_numeric(out["Ingreso limpieza"], errors="coerce").fillna(0).to_numpy()*ratio).round(2)
            out["Gasto limpieza"] = out["Ingreso limpieza"]
        elif limpieza_mode == "salida":
            mask = (fs >= pd.to_datetime(start_date)) & (fs <= pd.to_datetime(end_date))
            out["Ingreso limpieza"] = np.where(mask, out["Ingreso limpieza"], 0.0)
            out["Gasto limpieza"] = out["Ingreso limpieza"]
        elif limpieza_mode == "entrada":
            mask = (fe >= pd.to_datetime(start_date)) & (fe <= pd.to_datetime(end_date))
            out["Ingreso limpieza"] = np.where(mask, out["Ingreso limpieza"], 0.0)
            out["Gasto limpieza"] = out["Ingreso limpieza"]

    # Amenities
    if "Amenities" in out.columns:
        if amenities_mode == "prorratear":
            out["Amenities"] = (pd.to_numeric(out["Amenities"], errors="coerce").fillna(0).to_numpy()*ratio).round(2)
        elif amenities_mode == "salida":
            mask = (fs >= pd.to_datetime(start_date)) & (fs <= pd.to_datetime(end_date))
            out["Amenities"] = np.where(mask, out["Amenities"], 0.0)
        elif amenities_mode == "entrada":
            mask = (fe >= pd.to_datetime(start_date)) & (fe <= pd.to_datetime(end_date))
            out["Amenities"] = np.where(mask, out["Amenities"], 0.0)

    # Recalcular totales
    for c in ["Total Gastos","Pago al propietario","Pago recibido"]:
        if c in out.columns: out.drop(columns=[c], inplace=True, errors="ignore")

    if {"Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"}.issubset(out.columns):
        out["Total Gastos"] = out[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].sum(axis=1).round(2)
    if {"Total ingresos","Total Gastos"}.issubset(out.columns):
        out["Pago al propietario"] = (out["Total ingresos"] - out["Total Gastos"]).round(2)
    if {"Total ingresos","Comisión portal"}.issubset(out.columns):
        out["Pago recibido"] = (out["Total ingresos"] - out["Comisión portal"]).round(2)

    return out

# ========= Procesadores =========
def process_case1(df):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Ingreso limpieza","Total ingresos","Comisión portal","Portal"], "Caso 1")
    df = apply_booking_vat_on_commission(df, "Portal", "Comisión portal")

    def honorarios(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        pct = case1_percent_amenities.get(key,(0.20,0.0))[0]
        return float(r.get("Ingreso alojamiento",0.0)) * pct * 1.21

    def amenities(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        return float(case1_percent_amenities.get(key,(0.20,0.0))[1])

    out = df.copy()
    out["Honorarios Florit"] = out.apply(honorarios, axis=1)
    out["Gasto limpieza"]   = out.get("Ingreso limpieza", 0.0)
    out["Amenities"]        = out.apply(amenities, axis=1)
    out["Total Gastos"]     = out[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].sum(axis=1)
    out["Pago al propietario"] = out["Total ingresos"] - out["Total Gastos"]
    out["Pago recibido"]    = out["Total ingresos"] - out["Comisión portal"]

    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza",
            "Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities",
            "Total Gastos","Pago al propietario","Pago recibido"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(by=["Alojamiento","Fecha entrada"] if "Fecha entrada" in cols else ["Alojamiento"])

def process_case2(df):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Total ingresos","Comisión portal","Portal"], "Caso 2")

    # +21% comisión en Booking solo para APOLO 29/197
    mask_apolo = df["Alojamiento"].astype(str).str.upper().isin({"APOLO 29","APOLO 197"})
    mask_book  = df["Portal"].astype(str).str.lower().str.contains("booking", na=False)
    df.loc[mask_apolo & mask_book, "Comisión portal"] = pd.to_numeric(df.loc[mask_apolo & mask_book, "Comisión portal"], errors="coerce").fillna(0.0) * 1.21

    def honorarios(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        pct = case2_percent_amenities.get(key,(0.20,0.0))[0]
        base = float(r.get("Ingreso alojamiento",0.0)) - float(r.get("IVA del alquiler",0.0))
        return base * pct * 1.21

    def amenities(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        return float(case2_percent_amenities.get(key,(0.20,0.0))[1])

    out = df.copy()
    out["Honorarios Florit"] = out.apply(honorarios, axis=1)
    out["Gasto limpieza"]   = out.get("Ingreso limpieza", 0.0)
    out["Amenities"]        = out.apply(amenities, axis=1)
    out["Total Gastos"]     = out[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].sum(axis=1)
    out["Pago al propietario"] = out["Total ingresos"] - out["Total Gastos"]
    out["Pago recibido"]    = out["Total ingresos"] - out["Comisión portal"]

    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alquiler",
            "Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza",
            "Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(by=["Alojamiento","Fecha entrada"] if "Fecha entrada" in cols else ["Alojamiento"])

def process_case3(df):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Total ingresos","Comisión portal","Portal"], "Caso 3")
    df = apply_booking_vat_on_commission(df, "Portal", "Comisión portal")

    def honorarios(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        pct = case3_cleaning_amenities.get(key,(0.20,0.0,0.0))[0]
        base = float(r.get("Ingreso alojamiento",0.0)) - float(r.get("Comisión portal",0.0))
        return base * pct * 1.21

    def gasto_limpieza(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        return float(case3_cleaning_amenities.get(key,(0.20,0.0,0.0))[1])

    def amenities(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        return float(case3_cleaning_amenities.get(key,(0.20,0.0,0.0))[2])

    out = df.copy()
    out["Honorarios Florit"] = out.apply(honorarios, axis=1)
    out["Gasto limpieza"]   = out.apply(gasto_limpieza, axis=1)
    out["Amenities"]        = out.apply(amenities, axis=1)
    out["Total Gastos"]     = out[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].sum(axis=1)
    out["Pago al propietario"] = out["Total ingresos"] - out["Total Gastos"]
    out["Pago recibido"]    = out["Total ingresos"] - out["Comisión portal"]

    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza",
            "Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities",
            "Total Gastos","Pago al propietario","Pago recibido"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(by=["Alojamiento","Fecha entrada"] if "Fecha entrada" in cols else ["Alojamiento"])

def process_case4(df):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Comisión portal"], "Caso 4")

    def honorarios(r):
        base = float(r.get("Ingreso alojamiento",0.0)) - float(r.get("IVA del alquiler",0.0)) - float(r.get("Comisión portal",0.0))
        return base * 0.20

    out = df.copy()
    out["Honorarios Florit"] = out.apply(honorarios, axis=1)
    out["Gasto limpieza"]   = out.get("Ingreso limpieza", 0.0)
    out["Amenities"]        = 0.0
    out["Total Gastos"]     = out[["Comisión portal","Honorarios Florit"]].sum(axis=1) + out["Gasto limpieza"] + out["Amenities"]
    out["Pago al propietario"] = out["Ingreso alojamiento"] - out.get("IVA del alquiler",0.0) - out["Comisión portal"] - out["Honorarios Florit"]
    out["Pago recibido"]    = out["Ingreso alojamiento"] + out.get("Ingreso limpieza",0.0) - out["Comisión portal"]

    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alquiler",
            "Portal","Comisión portal","Honorarios Florit","Pago al propietario","Pago recibido","Ingreso limpieza"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(by=["Alojamiento","Fecha entrada"] if "Fecha entrada" in cols else ["Alojamiento"])

def process_case5(df):
    df = normalize_columns(df)
    ensure_required(df, ["Alojamiento","Ingreso alojamiento","Total ingresos","Comisión portal"], "Caso 5")

    def honorarios(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        pct = case5_percent_amenities.get(key,(0.20,0.0))[0]
        base = float(r.get("Ingreso alojamiento",0.0)) - float(r.get("IVA del alquiler",0.0)) - float(r.get("Comisión portal",0.0))
        return base * pct * 1.21

    def amenities(r):
        key = str(r.get("Alojamiento","")).strip().upper()
        return float(case5_percent_amenities.get(key,(0.20,0.0))[1])

    out = df.copy()
    out["Honorarios Florit"] = out.apply(honorarios, axis=1)
    out["Gasto limpieza"]   = out.get("Ingreso limpieza", 0.0)
    out["Amenities"]        = out.apply(amenities, axis=1)
    out["Total Gastos"]     = out[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].sum(axis=1)
    out["Pago al propietario"] = out["Total ingresos"] - out["Total Gastos"]
    out["Pago recibido"]    = out["Total ingresos"] - out["Comisión portal"]

    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alquiler",
            "Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza",
            "Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(by=["Alojamiento","Fecha entrada"] if "Fecha entrada" in cols else ["Alojamiento"])

processors = {1: process_case1, 2: process_case2, 3: process_case3, 4: process_case4, 5: process_case5}

# ========= UI (sin vista previa, botón en el menú) =========
st.title("📊 LIQUIDACIONES Automáticas (Casos 1–5)")
st.caption("Sube el Excel, elige período y caso, y pulsa Generar en la izquierda. Se mostrará directamente la liquidación.")

with st.sidebar:
    st.header("Parámetros")
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("Desde", value=date(date.today().year, date.today().month, 1))
    with c2:
        end_date   = st.date_input("Hasta",  value=date(date.today().year, date.today().month, 28))
    st.checkbox("Prorratear por noches del período", value=True, key="prorate")
    limpieza_mode = st.selectbox("Limpieza", ["prorratear","salida","entrada"], index=0)
    amenities_mode = st.selectbox("Amenities", ["prorratear","salida","entrada"], index=0)
    st.divider()
    case_choice = st.radio("Caso", [1,2,3,4,5], horizontal=True)
    st.checkbox("Lectura por letras (fallback)", value=False, key="by_letters")
    st.caption("Mapeo: W, D, F, H, I, J, O, AP, AR, AL (si Avantio cambia títulos).")
    generate = st.button("Generar liquidación")

file = st.file_uploader("Sube el archivo de reservas (.xlsx)", type=["xlsx"])

# ========= Generación directa =========
if generate:
    if not file:
        st.error("Sube primero el archivo de reservas (.xlsx).")
        st.stop()

    # Leer y asegurar columnas únicas
    df_in = pd.read_excel(file, header=0)
    df_in = ensure_unique_columns(df_in)

    # Normalizar (por nombre o por letras) y asegurar único
    df_norm = normalize_columns_by_letters(df_in) if st.session_state.by_letters else normalize_columns(df_in)
    df_norm = ensure_unique_columns(df_norm)

    # Filtrar a las propiedades del caso para que solo salga lo que toca
    props = props_for_case(case_choice)
    if props and "Alojamiento" in df_norm.columns:
        df_norm = df_norm[df_norm["Alojamiento"].isin(props)]

    # Procesar caso y volver a asegurar único
    base_df = processors[case_choice](df_norm.copy())
    base_df = ensure_unique_columns(base_df)

    # Aplicar periodo/prorrateo y asegurar único
    result_df = apply_period(
        base_df,
        start_date=start_date,
        end_date=end_date,
        prorate=st.session_state.prorate,
        limpieza_mode=limpieza_mode,
        amenities_mode=amenities_mode
    )
    result_df = ensure_unique_columns(result_df)

    st.success(f"Liquidación generada (Caso {case_choice}) • {start_date.strftime('%d/%m/%Y')}–{end_date.strftime('%d/%m/%Y')}")
    show_table_es(result_df, "Tabla de liquidaciones")

    if "Alojamiento" in result_df.columns and "Pago al propietario" in result_df.columns:
        pagos = (result_df.groupby("Alojamiento", as_index=False)["Pago al propietario"]
                 .sum().round(2).sort_values("Alojamiento"))
        show_table_es(pagos, "💸 Pagos por alojamiento (suma)")
        total_general = pagos["Pago al propietario"].sum()
        st.markdown(f"**Total general a transferir:** {fmt_es_num(total_general)} €")
