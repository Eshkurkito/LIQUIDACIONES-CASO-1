# -*- coding: utf-8 -*-
"""
liquidaciones_con_banco_v3_1.py

Versión 3.1 unificada (18/09/2025) con:
- Procesado de CASOS 1 a 5 (mismas reglas que liquidaciones_cases.py).
- Conciliación bancaria con tolerancia de fechas.
- Filtros por mes/año o por rango de fechas (from/to).
- Conmutadores: "portal vacío = Booking" y "aplicar +21% IVA a comisión de Booking" (all/apolo/none).
- Mapeo por nombres y fallback por letras (W, D, F, H, I, K, L/J, O, AP, AR).
- Exportación Excel con formato rico, agrupación por Alojamiento, encabezado repetido y subtotales.

Uso ejemplo:
python liquidaciones_con_banco_v3_1.py \
  --liq "RESERVAS.xlsx" --case 2 --month 7 --year 2025 \
  --bank "BBVA.xlsx" --bank-sheet "Historico" --bank-header 14 \
  --booking-apply-vat apolo --treat-empty-portal-as-booking \
  --out-liq "Reservas_Florit_CASO_2_julio2025.xlsx" \
  --out-recon "Conciliacion_Banco_vs_Liquidaciones_Julio2025.xlsx"
"""

import argparse
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

BOOKING_NAMES = {"Booking", "Booking.com", "BOOKING", "BOOKING.COM", "booking", "booking.com", ""}

# --------------------------
# Configuración Casos (tablas)
# --------------------------

CASE2_TABLE = {
    "VISITACION": (0.20, 14.88),
    "PADRE PORTA 6": (0.20, 12.09),
    "PADRE PORTA 7": (0.20, 12.09),
    "PADRE PORTA 8": (0.20, 12.09),
    "PADRE PORTA 9": (0.20, 12.09),
    "PADRE PORTA 10": (0.20, 12.09),
    "LLADRO Y MALLI 00": (0.20, 9.45),
    "LLADRO Y MALLI 01": (0.20, 9.45),
    "LLADRO Y MALLI 02": (0.20, 9.45),
    "LLADRO Y MALLI 03": (0.20, 9.45),
    "LLADRO Y MALLI 04": (0.20, 9.45),
    "APOLO 29": (0.20, 11.58),
    "APOLO 197": (0.20, 17.40),
}

CASE3_TABLE = {
    "ZAPATEROS 10-2": (0.20, 60.00, 15.24),
    "ZAPATEROS 10-6": (0.20, 75.00, 15.24),
    "ZAPATEROS 10-8": (0.20, 75.00, 15.24),
    "ZAPATEROS 12-5": (0.20, 60.00, 11.33),
    "ALFARO": (0.20, 80.00, 14.88),
}

CASE4_PERCENT = 0.20

CASE5_TABLE = {
    "HOMERO 01": (0.20, 0.00),
    "HOMERO 02": (0.20, 0.00),
    "CARCAIXENT 01": (0.20, 8.60),
    "CARCAIXENT 02": (0.20, 8.60),
}

# --------------------------
# Utilidades
# --------------------------

STANDARD_COLS = [
    ("Alojamiento", ["Alojamiento", "Nombre alojamiento", "Apartamento", "Propiedad", "Columna W", "W"]),
    ("Fecha entrada", ["Fecha entrada", "Fecha de entrada", "Check-in", "D entrada", "Columna D", "D"]),
    ("Fecha salida", ["Fecha salida", "Fecha de salida", "Check-out", "D salida", "Columna F", "F"]),
    ("Noches ocupadas", ["Noches ocupadas", "Noches", "Nº noches", "Columna H", "H"]),
    ("Ingreso alojamiento", ["Ingreso alojamiento", "Alquiler con tasas", "Alojamiento", "Columna I", "I"]),
    ("IVA del alojamiento", ["IVA del alojamiento", "Tasas del alquiler", "IVA alquiler", "Columna K", "K", "AL"]),
    ("Ingreso limpieza", ["Ingreso limpieza", "Extras con tasas", "Limpieza (ingreso)", "Columna L", "Columna J", "L", "J"]),
    ("Total ingresos", ["Total ingresos", "Total reserva", "Total reserva con tasas", "Columna O", "O"]),
    ("Portal", ["Portal", "Web origen", "Canal", "Columna AP", "AP"]),
    ("Comisión portal", ["Comisión portal", "Comisión del portal", "Fee portal", "Columna AR", "AR"]),
]

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    colmap = {}
    for std, candidates in STANDARD_COLS:
        for cand in candidates:
            if cand in df.columns:
                colmap[cand] = std
                break
    df = df.rename(columns=colmap).copy()

    for c in ["Fecha entrada", "Fecha salida"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in ["Noches ocupadas", "Ingreso alojamiento", "IVA del alojamiento", "Ingreso limpieza", "Total ingresos", "Comisión portal"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "Portal" in df.columns:
        df["Portal"] = df["Portal"].astype(str).str.strip()

    if "Alojamiento" in df.columns:
        df["Alojamiento"] = df["Alojamiento"].astype(str).str.strip().str.lower().str.title()

    return df

def month_name_es(m: int) -> str:
    return ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"][m-1]

def fmt_money(x):
    if pd.isna(x): 
        return np.nan
    return float(round(float(x), 2))

def apply_booking_vat(df: pd.DataFrame, scope: str, treat_empty_as_booking: bool):
    """scope = 'none' | 'all' | 'apolo'"""
    if "Portal" not in df or "Comisión portal" not in df or "Alojamiento" not in df:
        return df

    if treat_empty_as_booking:
        mask_booking = df["Portal"].isin(BOOKING_NAMES)
    else:
        mask_booking = df["Portal"].isin(BOOKING_NAMES - {""})

    if scope == "none":
        return df
    elif scope == "all":
        mask = mask_booking
    elif scope == "apolo":
        mask = mask_booking & df["Alojamiento"].str.upper().isin({"APOLO 29", "APOLO 197"})
    else:
        mask = pd.Series(False, index=df.index)

    df.loc[mask, "Comisión portal"] = df.loc[mask, "Comisión portal"] * 1.21
    return df

def ensure_cols(df: pd.DataFrame, cols: list[str]):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}")

# --------------------------
# Reglas por caso
# --------------------------

def case1(df: pd.DataFrame) -> pd.DataFrame:
    ensure_cols(df, ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal"])
    df["% Honorarios"] = 0.20
    df["Honorarios Florit"] = df["Ingreso alojamiento"] * df["% Honorarios"] * 1.21
    df["Gasto limpieza"] = df["Ingreso limpieza"]
    df["Amenities"] = 0.00
    df["Total Gastos"] = df[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].fillna(0).sum(axis=1)
    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]
    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    return df[cols]

def case2(df: pd.DataFrame) -> pd.DataFrame:
    ensure_cols(df, ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal"])
    def pct_am(a):
        key = str(a).strip().upper()
        return CASE2_TABLE.get(key, (0.20, 0.00))
    vals = df["Alojamiento"].apply(pct_am)
    df["% Honorarios"] = vals.apply(lambda t: t[0])
    df["Amenities"] = vals.apply(lambda t: t[1])
    df["Honorarios Florit"] = (df["Ingreso alojamiento"] - df["IVA del alojamiento"]) * df["% Honorarios"] * 1.21
    df["Gasto limpieza"] = df["Ingreso limpieza"]
    df["Total Gastos"] = df[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].fillna(0).sum(axis=1)
    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]
    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    return df[cols]

def case3(df: pd.DataFrame) -> pd.DataFrame:
    ensure_cols(df, ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal"])
    def pct_clean_am(a):
        key = str(a).strip().upper()
        return CASE3_TABLE.get(key, (0.20, 0.00, 0.00))
    vals = df["Alojamiento"].apply(pct_clean_am)
    df["% Honorarios"] = vals.apply(lambda t: t[0])
    df["Gasto limpieza"] = vals.apply(lambda t: t[1])
    df["Amenities"] = vals.apply(lambda t: t[2])
    df["Honorarios Florit"] = (df["Ingreso alojamiento"] - df["Comisión portal"]) * df["% Honorarios"] * 1.21
    df["Total Gastos"] = df[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].fillna(0).sum(axis=1)
    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]
    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    return df[cols]

def case4(df: pd.DataFrame) -> pd.DataFrame:
    ensure_cols(df, ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Portal","Comisión portal"])
    df["% Honorarios"] = CASE4_PERCENT
    df["Honorarios Florit"] = (df["Ingreso alojamiento"] - df["IVA del alojamiento"] - df["Comisión portal"]) * df["% Honorarios"]
    df["Pago al propietario"] = df["Ingreso alojamiento"] - df["IVA del alojamiento"] - df["Comisión portal"] - df["Honorarios Florit"]
    if "Ingreso limpieza" in df.columns:
        df["Pago recibido"] = df["Ingreso alojamiento"] + df["Ingreso limpieza"] - df["Comisión portal"]
        extra_cols = ["Ingreso limpieza"]
    else:
        df["Pago recibido"] = df["Ingreso alojamiento"] - df["Comisión portal"]
        extra_cols = []
    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Portal","Comisión portal","Honorarios Florit","Pago al propietario","Pago recibido"] + extra_cols
    return df[cols]

def case5(df: pd.DataFrame) -> pd.DataFrame:
    ensure_cols(df, ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal"])
    def pct_am(a):
        key = str(a).strip().upper()
        return CASE5_TABLE.get(key, (0.20, 0.00))
    vals = df["Alojamiento"].apply(pct_am)
    df["% Honorarios"] = vals.apply(lambda t: t[0])
    df["Amenities"] = vals.apply(lambda t: t[1])
    df["Gasto limpieza"] = df["Ingreso limpieza"]
    df["Honorarios Florit"] = (df["Ingreso alojamiento"] - df["IVA del alojamiento"] - df["Comisión portal"]) * df["% Honorarios"] * 1.21
    df["Total Gastos"] = df[["Comisión portal","Honorarios Florit","Gasto limpieza","Amenities"]].fillna(0).sum(axis=1)
    df["Pago al propietario"] = df["Total ingresos"] - df["Total Gastos"]
    df["Pago recibido"] = df["Total ingresos"] - df["Comisión portal"]
    cols = ["Alojamiento","Fecha entrada","Fecha salida","Noches ocupadas","Ingreso alojamiento","IVA del alojamiento","Ingreso limpieza","Total ingresos","Portal","Comisión portal","Honorarios Florit","Gasto limpieza","Amenities","Total Gastos","Pago al propietario","Pago recibido"]
    return df[cols]

# --------------------------
# Presentación Excel
# --------------------------

def export_with_groups(df: pd.DataFrame, outfile: str):
    df = df.sort_values(["Alojamiento","Fecha entrada"]).copy()
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

    # Columna vacía entre Pago al propietario y Pago recibido si existen consecutivos
    if "Pago al propietario" in df.columns and "Pago recibido" in df.columns:
        cols = list(df.columns)
        i_prop = cols.index("Pago al propietario")
        i_recv = cols.index("Pago recibido")
        if i_recv == i_prop + 1:
            cols = cols[:i_recv] + [""] + cols[i_recv:]
            df = df.reindex(columns=cols)

    writer = pd.ExcelWriter(outfile, engine="xlsxwriter")
    sheet = "Liquidación"
    df0 = df.iloc[0:0]
    df0.to_excel(writer, index=False, sheet_name=sheet)
    wb = writer.book
    ws = writer.sheets[sheet]

    fmt_header = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
    fmt_cell = wb.add_format({"border": 1})
    fmt_money2 = wb.add_format({"border": 1, "num_format": "#,##0.00"})
    fmt_date = wb.add_format({"border": 1, "num_format": "dd/mm/yyyy"})
    fmt_bold = wb.add_format({"bold": True, "border": 1})
    fmt_group_title = wb.add_format({"bold": True})

    row = 0
    for prop, chunk in df.groupby("Alojamiento", sort=False):
        ws.write(row, 0, str(prop), fmt_group_title)
        row += 1
        for j, col in enumerate(df.columns):
            ws.write(row, j, col if col != "" else "", fmt_header)
        row += 1

        for _, r in chunk.iterrows():
            for j, col in enumerate(df.columns):
                val = r[col] if col in r else None
                if isinstance(val, (pd.Timestamp, datetime)):
                    ws.write_datetime(row, j, pd.to_datetime(val).to_pydatetime(), fmt_date)
                elif pd.isna(val):
                    ws.write(row, j, "", fmt_cell)
                elif isinstance(val, (int, float, np.number)):
                    ws.write_number(row, j, float(val), fmt_money2)
                else:
                    ws.write(row, j, str(val), fmt_cell)
            row += 1

        ws.write(row, 0, "Subtotal →", fmt_bold)
        for j, col in enumerate(df.columns):
            if col in numeric_cols:
                # Construir rango Excel
                col_letter = chr(ord('A') + j)
                start = row - len(chunk)
                end = row - 1
                ws.write_formula(row, j, f"=SUM({col_letter}{start+1}:{col_letter}{end+1})", fmt_money2)
            else:
                if j != 0:
                    ws.write(row, j, "", fmt_bold)
        row += 2

    for j, col in enumerate(df.columns):
        width = 18 if col in ("Alojamiento","Portal","") else 12 if "Fecha" in col else 16
        ws.set_column(j, j, width)

    writer.close()

# --------------------------
# Conciliación
# --------------------------

def read_bank_excel(path: str, sheet: str|None, header_row: int):
    if sheet:
        dfb = pd.read_excel(path, sheet_name=sheet, header=header_row-1)
    else:
        dfb = pd.read_excel(path, header=header_row-1)
    return dfb

def normalize_bank(df: pd.DataFrame):
    # Intentar deducir columnas clave comunes: Fecha, Importe, Concepto
    cand_date = [c for c in df.columns if "fecha" in str(c).lower() or "date" in str(c).lower()]
    cand_amount = [c for c in df.columns if "importe" in str(c).lower() or "amount" in str(c).lower() or "cargo" in str(c).lower() or "abono" in str(c).lower() or "debit" in str(c).lower() or "credit" in str(c).lower()]
    cand_desc = [c for c in df.columns if "concept" in str(c).lower() or "conce" in str(c).lower() or "descrip" in str(c).lower() or "detalle" in str(c).lower() or "narrative" in str(c).lower()]

    date_col = cand_date[0] if cand_date else df.columns[0]
    # Importe: preferimos una sola columna 'Importe' (positivo cobro), si existen cargo/abono los combinamos
    if any("cargo" in str(c).lower() for c in df.columns) and any("abono" in str(c).lower() for c in df.columns):
        cargo = [c for c in df.columns if "cargo" in str(c).lower()][0]
        abono = [c for c in df.columns if "abono" in str(c).lower()][0]
        df["Importe"] = pd.to_numeric(df[abono], errors="coerce").fillna(0) - pd.to_numeric(df[cargo], errors="coerce").fillna(0)
    else:
        amount_col = cand_amount[0] if cand_amount else df.columns[1]
        df["Importe"] = pd.to_numeric(df[amount_col], errors="coerce")

    desc_col = cand_desc[0] if cand_desc else (df.columns[2] if len(df.columns) > 2 else df.columns[0])

    out = pd.DataFrame({
        "Fecha banco": pd.to_datetime(df[date_col], errors="coerce"),
        "Importe banco": df["Importe"],
        "Descripción banco": df[desc_col].astype(str)
    })
    out = out.dropna(subset=["Fecha banco","Importe banco"])
    return out

def reconcile(liq: pd.DataFrame, bank: pd.DataFrame, reconcile_by: str, date_window_days: int):
    """
    reconcile_by: 'pago_recibido' (default) o 'pago_propietario'
    date_window_days: -1 = sin límite; 0+ = tolerancia en días por ambos lados
    """
    if reconcile_by == "pago_propietario":
        amount_col = "Pago al propietario"
        date_col = "Fecha salida" if "Fecha salida" in liq.columns else "Fecha entrada"
    else:
        amount_col = "Pago recibido"
        date_col = "Fecha entrada"

    # Prepara claves
    liq2 = liq.copy()
    liq2["__amt__"] = liq2[amount_col].round(2)
    liq2["__date__"] = pd.to_datetime(liq2[date_col], errors="coerce")

    bank2 = bank.copy()
    bank2["__amt__"] = bank2["Importe banco"].round(2)
    bank2["__date__"] = pd.to_datetime(bank2["Fecha banco"], errors="coerce")

    liq2 = liq2.dropna(subset=["__amt__","__date__"])
    bank2 = bank2.dropna(subset=["__amt__","__date__"])

    matches = []
    used_bank_idx = set()

    for i, r in liq2.iterrows():
        target_amt = r["__amt__"]
        target_date = r["__date__"]

        # candidatos por importe exacto
        cands = bank2[bank2["__amt__"].eq(target_amt)].copy()
        if date_window_days >= 0:
            start = target_date - timedelta(days=date_window_days)
            end = target_date + timedelta(days=date_window_days)
            cands = cands[(cands["__date__"] >= start) & (cands["__date__"] <= end)]

        # tomar el más cercano en fecha que no esté usado
        cands = cands[~cands.index.isin(used_bank_idx)]
        if cands.empty:
            continue
        cands["abs_delta_days"] = (cands["__date__"] - target_date).abs()
        j = cands.sort_values("abs_delta_days").index[0]

        used_bank_idx.add(j)
        matches.append((i, j))

    # Construir salidas
    liq_matched = []
    for i, j in matches:
        row = {**liq2.loc[i].to_dict(), **{f"Banco::{k}": bank2.loc[j][k] for k in ["Fecha banco","Importe banco","Descripción banco"]}}
        liq_matched.append(row)
    df_matched = pd.DataFrame(liq_matched)

    liq_unmatched = liq2[~liq2.index.isin([i for i, _ in matches])].copy()
    bank_unmatched = bank2[~bank2.index.isin([j for _, j in matches])].copy()

    return df_matched, bank_unmatched, liq_unmatched

# --------------------------
# CLI
# --------------------------

def main():
    p = argparse.ArgumentParser(description="Liquidaciones por casos + Conciliación bancaria (v3.1)")
    p.add_argument("--liq", required=True, help="Excel de reservas/liquidaciones origen")
    p.add_argument("--liq-header", type=int, default=1, help="Fila (1-based) de encabezado en --liq")
    p.add_argument("--case", type=int, choices=[1,2,3,4,5], required=True, help="Número de caso a aplicar")

    # Fechas: mes/año o rango
    p.add_argument("--filter-by", choices=["entry","exit"], default="entry", help="Filtrar por entrada o salida")
    p.add_argument("--month", type=int, help="Mes (1-12)")
    p.add_argument("--year", type=int, help="Año (YYYY)")
    p.add_argument("--from-date", help="YYYY-MM-DD (prioridad si se indica)")
    p.add_argument("--to-date", help="YYYY-MM-DD (prioridad si se indica)")

    # Normalización portal
    p.add_argument("--portal-col", help="Nombre/Letra de columna Portal si quieres forzarla")
    p.add_argument("--treat-empty-portal-as-booking", action="store_true", help="Considerar portal vacío como Booking")
    p.add_argument("--booking-apply-vat", choices=["none","all","apolo"], default="none", help="Añadir +21% a Comisión si portal es Booking")

    # Banco
    p.add_argument("--bank", required=True, help="Excel de banco")
    p.add_argument("--bank-sheet", help="Hoja del banco")
    p.add_argument("--bank-header", type=int, default=1, help="Fila (1-based) de encabezado en --bank")
    p.add_argument("--date-window-days", type=int, default=-1, help="-1 sin límite; 0+ tolerancia en días")
    p.add_argument("--reconcile-by", choices=["pago_recibido","pago_propietario"], default="pago_recibido")

    # Salidas
    p.add_argument("--out-liq", help="Ruta Excel salida de liquidaciones por caso")
    p.add_argument("--out-recon", help="Ruta Excel salida conciliación (3 hojas)")

    args = p.parse_args()

    # Leer liquidaciones
    df = pd.read_excel(args.liq, header=args.liq_header-1)

    # Portal forzado si procede
    if args.portal_col and args.portal_col in df.columns:
        df["Portal"] = df[args.portal_col]

    # Normalizar columnas (nombres/letras)
    df = normalize_cols(df)

    # Filtro fechas
    date_col = "Fecha entrada" if args.filter_by == "entry" else "Fecha salida"
    if args.from_date and args.to_date:
        start = pd.to_datetime(args.from_date)
        end = pd.to_datetime(args.to_date)
        df = df[(df[date_col] >= start) & (df[date_col] <= end)]
        mes_txt = f"{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}"
    else:
        if args.year:
            df = df[df[date_col].dt.year == args.year]
        if args.month:
            df = df[df[date_col].dt.month == args.month]
        mes_txt = f"{month_name_es(args.month)}{args.year}" if (args.month and args.year) else datetime.now().strftime("%Y%m%d_%H%M%S")

    if df.empty:
        raise SystemExit("No hay filas tras el filtro de fechas.")

    # Aplicar IVA Booking según flag
    df = apply_booking_vat(df, scope=args.booking_apply_vat, treat_empty_as_booking=args.treat_empty_portal_as_booking)

    # Calcular caso
    if args.case == 1:
        out_df = case1(df)
    elif args.case == 2:
        out_df = case2(df)
    elif args.case == 3:
        out_df = case3(df)
    elif args.case == 4:
        out_df = case4(df)
    elif args.case == 5:
        out_df = case5(df)
    else:
        raise SystemExit("Caso no soportado.")

    # Redondeo monetario
    for c in out_df.columns:
        if pd.api.types.is_numeric_dtype(out_df[c]):
            out_df[c] = out_df[c].apply(fmt_money)

    # Nombre salida liquidaciones
    out_liq = args.out_liq or f"Reservas_Florit_CASO_{args.case}_{mes_txt}.xlsx"
    export_with_groups(out_df, out_liq)

    # Leer y normalizar banco
    dfb = read_bank_excel(args.bank, args.bank_sheet, args.bank_header)
    dfb = normalize_bank(dfb)

    # Conciliar
    matched, bank_only, liq_only = reconcile(out_df, dfb, reconcile_by=args.reconcile_by, date_window_days=args.date_window_days)

    # Export conciliación
    out_recon = args.out_recon or f"Conciliacion_Banco_vs_Liquidaciones_CASO_{args.case}_{mes_txt}.xlsx"
    with pd.ExcelWriter(out_recon, engine="xlsxwriter") as w:
        (matched if not matched.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name="Conciliados")
        (bank_only if not bank_only.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name="Banco_sin_match")
        (liq_only if not liq_only.empty else pd.DataFrame()).to_excel(w, index=False, sheet_name="Liq_sin_match")

    print(f"✅ Liquidaciones: {out_liq}")
    print(f"✅ Conciliación: {out_recon}")

if __name__ == "__main__":
    main()
