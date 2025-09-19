#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
liquidaciones_con_banco_v3_estricto.py

Conciliación estricta banco-reservas para entornos de apartamentos turísticos.
Lógica:
  1) Match 1:1 exacto por importe (en céntimos) con prioridad por portal y base de apartamento.
  2) Si no cuadra, buscar combinaciones (hasta K) que sumen exactamente el importe del banco.
  3) Exporta resultados en un Excel con varias hojas.

Pensado para:
  - Casos donde bancos agrupan pagos de varias reservas (p. ej., Booking fin de semana).
  - Varios portales (Booking, Airbnb, etc.).
  - Mismo apartamento (base) con diferentes números (TRAFALGAR 01, 02...) en un mismo pago.

Autor: Kai (ChatGPT)
Fecha: 2025-09-19
"""

from __future__ import annotations
import argparse
import itertools
import math
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd

# =========================
# CONFIGURACIÓN COLUMNAS
# =========================
# Ajusta estos nombres si tus columnas difieren.
CONFIG = {
    "reservas": {
        "id_reserva": "reserva_id",     # ID único de la reserva
        "apartamento": "apartamento",   # Nombre del apartamento (p.ej., "TRAFALGAR 01")
        "portal": "portal",             # Portal de origen (p.ej., "Booking", "Airbnb")
        "importe": "importe",           # Importe esperado a conciliar (numérico)
        "fecha": "fecha_reserva",       # Fecha de la reserva o de liquidación (opcional)
    },
    "banco": {
        "id_banco": "mov_id",           # ID único del movimiento bancario (si no existe, se generará)
        "fecha": "fecha",               # Fecha del movimiento
        "concepto": "concepto",         # Concepto / descrip. del banco
        "importe": "importe",           # Importe del movimiento (positivo para abonos)
        "portal": "portal",             # (Opcional) Portal si viene identificado en extracto
    },
    # Tolerancia en céntimos para considerar "igual" (0 = estrictos)
    "tolerancia_centimos": 0,
    # Tamaño máximo de combinación a buscar
    "max_k_default": 4,
    # Máximo de candidatos por grupo para limitar combinatoria
    "max_candidatos_por_grupo": 40,
    # Ventana de días opcional para filtrar candidatos por fecha respecto al movimiento del banco (None = sin filtro)
    "ventana_dias": None,
}


def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(path)
    elif ext in [".csv", ".txt"]:
        # Try sensible defaults; user can re-save if needed
        try:
            return pd.read_csv(path)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin1", sep=";")
    else:
        raise ValueError(f"Extensión no soportada: {ext}")


def to_cents(x) -> int:
    if pd.isna(x):
        return 0
    try:
        # Convert to int cents to avoid float issues
        v = float(x)
    except Exception:
        v = float(str(x).replace(",", "."))
    cents = int(round(v * 100))
    return cents


def cents_equal(a: int, b: int, tol: int = 0) -> bool:
    return abs(a - b) <= tol


def normalize_apto_base(name: str) -> str:
    """
    Normaliza el 'base' del apartamento quitando sufijos numéricos y espacios.
    Ej.: 'TRAFALGAR 01' -> 'TRAFALGAR'
         'SERRERIA 12'  -> 'SERRERIA'
    """
    if not isinstance(name, str):
        return ""
    s = name.strip().upper()
    # quitar números finales y separadores
    s = re.sub(r"[\s\-_/]*\d{1,3}$", "", s).strip()
    # colapsar espacios múltiples
    s = re.sub(r"\s+", " ", s)
    return s


def safe_upper(s):
    return s.upper() if isinstance(s, str) else s


@dataclass
class Reserva:
    idx: int
    reserva_id: str
    apartamento: str
    portal: str
    importe_cents: int
    fecha: Optional[pd.Timestamp] = None
    base_apto: str = field(init=False)

    def __post_init__(self):
        self.base_apto = normalize_apto_base(self.apartamento)


@dataclass
class MovimientoBanco:
    idx: int
    mov_id: str
    fecha: Optional[pd.Timestamp]
    concepto: str
    importe_cents: int
    portal: Optional[str] = None


def build_reservas(df_res: pd.DataFrame) -> List[Reserva]:
    R = CONFIG["reservas"]
    out = []
    for i, row in df_res.reset_index(drop=True).iterrows():
        rid = str(row.get(R["id_reserva"], f"RES_{i}"))
        apto = row.get(R["apartamento"], "")
        portal = row.get(R["portal"], "")
        imp = to_cents(row.get(R["importe"], 0))
        fecha_val = row.get(R["fecha"], None)
        fecha = pd.to_datetime(fecha_val) if pd.notna(fecha_val) else None
        out.append(Reserva(i, rid, str(apto), str(portal), imp, fecha))
    return out


def build_banco(df_bnk: pd.DataFrame) -> List[MovimientoBanco]:
    B = CONFIG["banco"]
    out = []
    for i, row in df_bnk.reset_index(drop=True).iterrows():
        bid = str(row.get(B["id_banco"], f"MOV_{i}"))
        fecha_val = row.get(B["fecha"], None)
        fecha = pd.to_datetime(fecha_val) if pd.notna(fecha_val) else None
        concepto = str(row.get(B["concepto"], ""))
        imp = to_cents(row.get(B["importe"], 0))
        portal = row.get(B.get("portal", "portal"), None)
        portal = str(portal) if pd.notna(portal) else None
        out.append(MovimientoBanco(i, bid, fecha, concepto, imp, portal))
    return out


def filter_candidates(reservas_pend: List[Reserva],
                      mov: MovimientoBanco,
                      ventana_dias: Optional[int]) -> List[Reserva]:
    """
    Heurística de filtrado para reducir candidatos:
      - Si mov.portal está definido, prioriza ese portal.
      - Si ventana_dias está configurada y la reserva tiene fecha, filtra por proximidad.
    """
    cand = reservas_pend
    if mov.portal:
        p = safe_upper(mov.portal)
        cand = [r for r in cand if safe_upper(r.portal) == p] or cand

    if ventana_dias is not None and mov.fecha is not None:
        lb = mov.fecha - pd.Timedelta(days=ventana_dias)
        ub = mov.fecha + pd.Timedelta(days=ventana_dias)
        # Mantener primero los que caen dentro de ventana; si vacío, no filtrar
        in_win = [r for r in cand if (r.fecha is not None and lb <= r.fecha <= ub)]
        cand = in_win or cand

    return cand


def group_candidates(cands: List[Reserva]) -> Dict[Tuple[str, str], List[Reserva]]:
    """
    Agrupa por (portal, base_apto) para priorizar combinaciones dentro del mismo 'bloque'.
    """
    groups: Dict[Tuple[str, str], List[Reserva]] = {}
    for r in cands:
        key = (safe_upper(r.portal or ""), r.base_apto)
        groups.setdefault(key, []).append(r)
    return groups


def find_exact_single(cands: List[Reserva], target_cents: int, tol: int) -> Optional[List[Reserva]]:
    for r in cands:
        if cents_equal(r.importe_cents, target_cents, tol):
            return [r]
    return None


def find_exact_combo(cands: List[Reserva],
                     target_cents: int,
                     tol: int,
                     max_k: int,
                     cap: int) -> Optional[List[Reserva]]:
    """
    Búsqueda exacta de combinación que suma target_cents.
    - Ordena por importe descendente para mejor poda.
    - Limita a 'cap' candidatos para evitar explosión combinatoria.
    - Explora combos de 2..max_k.
    """
    if not cands:
        return None

    # Preselección por cap (los más cercanos al target by value)
    # Orden por cercanía descendente
    cands_sorted = sorted(cands, key=lambda r: r.importe_cents, reverse=True)[:cap]

    values = [r.importe_cents for r in cands_sorted]

    # Rápido: probar pares con diccionario (k=2)
    if max_k >= 2:
        seen: Dict[int, int] = {}
        for i, v in enumerate(values):
            rem = target_cents - v
            if (rem in seen) and cents_equal(v + values[seen[rem]], target_cents, tol):
                return [cands_sorted[i], cands_sorted[seen[rem]]]
            seen[v] = i

    # k >= 3: DFS con poda
    n = len(values)

    def dfs(start: int, k_left: int, sum_so_far: int, chosen_idx: List[int]) -> Optional[List[int]]:
        # poda por suma excesiva
        if sum_so_far > target_cents + tol:
            return None
        if k_left == 0:
            if cents_equal(sum_so_far, target_cents, tol):
                return list(chosen_idx)
            return None
        # cota superior simple: sumar los mayores posibles restantes
        # (no muy estricta pero ayuda algo)
        max_possible = sum_so_far + sum(values[start:start + k_left])
        if max_possible < target_cents - tol:
            return None

        for i in range(start, n):
            chosen_idx.append(i)
            res = dfs(i + 1, k_left - 1, sum_so_far + values[i], chosen_idx)
            if res is not None:
                return res
            chosen_idx.pop()
        return None

    for k in range(3, max_k + 1):
        res_idx = dfs(0, k, 0, [])
        if res_idx is not None:
            return [cands_sorted[i] for i in res_idx]

    return None


def reconcile(reservas: List[Reserva],
              movimientos: List[MovimientoBanco],
              tol_cent: int,
              max_k: int,
              cap: int,
              ventana_dias: Optional[int]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Devuelve:
      - df_match_1to1
      - df_match_combo
      - df_unmatched_banco
      - df_unmatched_reservas
      - df_log (detalle línea a línea)
    """
    # Copias de trabajo (pendientes)
    pend_res: List[Reserva] = list(reservas)
    pend_by_id: Dict[str, Reserva] = {r.reserva_id: r for r in pend_res}

    # Logs
    log_rows = []
    matches_1to1 = []
    matches_combo = []

    # Iterar movimientos bancarios
    for mov in movimientos:
        # Filtrar candidatos heurísticos
        cands_all = filter_candidates(pend_res, mov, ventana_dias)

        # Intentar 1:1 exacto
        pick = find_exact_single(cands_all, mov.importe_cents, tol_cent)
        if pick:
            r = pick[0]
            matches_1to1.append((mov, [r]))
            # remover reserva escogida de pendientes
            pend_res.remove(r)
            pend_by_id.pop(r.reserva_id, None)
            log_rows.append({
                "mov_id": mov.mov_id,
                "mov_fecha": mov.fecha,
                "mov_concepto": mov.concepto,
                "mov_importe": mov.importe_cents / 100.0,
                "tipo_match": "1to1",
                "num_reservas": 1,
                "reserva_ids": r.reserva_id,
                "portales": r.portal,
                "apartamentos": r.apartamento,
                "importe_reservas_sum": r.importe_cents / 100.0,
                "observaciones": ""
            })
            continue

        # Si no 1:1, probar por grupos (portal, base_apto) para reducir combinaciones
        groups = group_candidates(cands_all)
        combo_found: Optional[List[Reserva]] = None
        chosen_key = None

        # Prioridad: grupos con más candidatos suelen dar mejor probabilidad
        for key, group in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            combo = find_exact_combo(group, mov.importe_cents, tol_cent, max_k, cap)
            if combo:
                combo_found = combo
                chosen_key = key
                break

        # Si no hubo suerte en grupo, probar sobre todos los candidatos (cap)
        if combo_found is None:
            combo_found = find_exact_combo(cands_all, mov.importe_cents, tol_cent, max_k, cap)

        if combo_found:
            for r in combo_found:
                if r in pend_res:
                    pend_res.remove(r)
                    pend_by_id.pop(r.reserva_id, None)
            matches_combo.append((mov, combo_found))
            log_rows.append({
                "mov_id": mov.mov_id,
                "mov_fecha": mov.fecha,
                "mov_concepto": mov.concepto,
                "mov_importe": mov.importe_cents / 100.0,
                "tipo_match": "combo",
                "num_reservas": len(combo_found),
                "reserva_ids": ", ".join([r.reserva_id for r in combo_found]),
                "portales": ", ".join([r.portal for r in combo_found]),
                "apartamentos": ", ".join([r.apartamento for r in combo_found]),
                "importe_reservas_sum": sum(r.importe_cents for r in combo_found) / 100.0,
                "observaciones": f"grupo={chosen_key}" if chosen_key else ""
            })
        else:
            # Sin match: dejará como pendiente de banco
            log_rows.append({
                "mov_id": mov.mov_id,
                "mov_fecha": mov.fecha,
                "mov_concepto": mov.concepto,
                "mov_importe": mov.importe_cents / 100.0,
                "tipo_match": "sin_match",
                "num_reservas": 0,
                "reserva_ids": "",
                "portales": "",
                "apartamentos": "",
                "importe_reservas_sum": 0.0,
                "observaciones": "No se encontraron candidatos exactos"
            })

    # Construir dataframes de salida
    def mov_to_row(mov: MovimientoBanco) -> Dict:
        return {
            "mov_id": mov.mov_id,
            "fecha": mov.fecha,
            "concepto": mov.concepto,
            "importe": mov.importe_cents / 100.0,
            "portal": mov.portal
        }

    def res_to_row(r: Reserva) -> Dict:
        return {
            "reserva_id": r.reserva_id,
            "apartamento": r.apartamento,
            "apartamento_base": r.base_apto,
            "portal": r.portal,
            "importe": r.importe_cents / 100.0,
            "fecha_reserva": r.fecha
        }

    df_log = pd.DataFrame(log_rows)

    df_match_1to1 = pd.DataFrame([
        dict(mov_to_row(mov), **{
            "num_reservas": 1,
            "reserva_ids": r[0].reserva_id,
            "apartamentos": r[0].apartamento,
            "portales": r[0].portal,
            "importe_reservas_sum": r[0].importe_cents / 100.0,
        })
        for (mov, r) in matches_1to1
    ])

    df_match_combo = pd.DataFrame([
        dict(mov_to_row(mov), **{
            "num_reservas": len(rs),
            "reserva_ids": ", ".join([x.reserva_id for x in rs]),
            "apartamentos": ", ".join([x.apartamento for x in rs]),
            "portales": ", ".join([x.portal for x in rs]),
            "importe_reservas_sum": sum(x.importe_cents for x in rs) / 100.0,
        })
        for (mov, rs) in matches_combo
    ])

    # Unmatched
    matched_mov_ids = set(df_log.loc[df_log["tipo_match"].isin(["1to1", "combo"]), "mov_id"].unique())
    all_mov_ids = {m.mov_id for m in movimientos}
    unmatched_mov_ids = list(all_mov_ids - matched_mov_ids)
    df_unmatched_banco = pd.DataFrame([mov_to_row(m) for m in movimientos if m.mov_id in unmatched_mov_ids])

    # Reservas pendientes
    df_unmatched_reservas = pd.DataFrame([res_to_row(r) for r in pend_res])

    # Ordenar por fecha si existe
    for df in [df_log, df_match_1to1, df_match_combo, df_unmatched_banco, df_unmatched_reservas]:
        if "fecha" in df.columns:
            df.sort_values("fecha", inplace=True, kind="stable")

    return df_match_1to1, df_match_combo, df_unmatched_banco, df_unmatched_reservas, df_log


def export_excel(out_path: str,
                 df_match_1to1: pd.DataFrame,
                 df_match_combo: pd.DataFrame,
                 df_unmatched_banco: pd.DataFrame,
                 df_unmatched_reservas: pd.DataFrame,
                 df_log: pd.DataFrame) -> None:
    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        df_match_1to1.to_excel(writer, index=False, sheet_name="MATCH_1to1")
        df_match_combo.to_excel(writer, index=False, sheet_name="MATCH_COMBO")
        df_unmatched_banco.to_excel(writer, index=False, sheet_name="BANCO_SIN_MATCH")
        df_unmatched_reservas.to_excel(writer, index=False, sheet_name="RESERVAS_PENDIENTES")
        df_log.to_excel(writer, index=False, sheet_name="LOG_DETALLE")

        # Formatos sencillos de verificación
        wb = writer.book
        fmt_money = wb.add_format({"num_format": "#,##0.00"})
        for sheet in ["MATCH_1to1", "MATCH_COMBO", "BANCO_SIN_MATCH", "RESERVAS_PENDIENTES", "LOG_DETALLE"]:
            ws = writer.sheets[sheet]
            # aplicar formato monetario si existen columnas de dinero
            for col_idx, col_name in enumerate(pd.read_excel(writer, sheet_name=sheet).columns):
                if "importe" in col_name.lower():
                    ws.set_column(col_idx, col_idx, 14, fmt_money)
                else:
                    ws.set_column(col_idx, col_idx, 18)


def main():
    parser = argparse.ArgumentParser(description="Conciliación estricta de liquidaciones con banco.")
    parser.add_argument("--banco", required=True, help="Archivo de movimientos de banco (CSV/XLSX).")
    parser.add_argument("--reservas", required=True, help="Archivo de reservas (CSV/XLSX).")
    parser.add_argument("--out", default="conciliacion_salida.xlsx", help="Salida Excel.")
    parser.add_argument("--max-k", type=int, default=CONFIG["max_k_default"], help="Tamaño máximo de combinación.")
    parser.add_argument("--cap-grupo", type=int, default=CONFIG["max_candidatos_por_grupo"], help="Máx. candidatos por grupo.")
    parser.add_argument("--ventana-dias", type=int, default=-1, help="Ventana de días para filtrar por fecha ( -1 = sin filtro ).")
    parser.add_argument("--tolerancia-centimos", type=int, default=CONFIG["tolerancia_centimos"], help="Tolerancia en céntimos.")
    parser.add_argument("--idcols-res", nargs="*", default=None, help="(Opcional) Nombres reales de columnas en reservas: id, apartamento, portal, importe, fecha (en ese orden).")
    parser.add_argument("--idcols-bnk", nargs="*", default=None, help="(Opcional) Nombres reales de columnas en banco: id, fecha, concepto, importe, portal (en ese orden).")

    args = parser.parse_args()

    # Aplicar overrides de columnas si se pasan
    if args.idcols_res:
        keys = ["id_reserva", "apartamento", "portal", "importe", "fecha"]
        for k, v in zip(keys, args.idcols_res):
            CONFIG["reservas"][k] = v
    if args.idcols_bnk:
        keys = ["id_banco", "fecha", "concepto", "importe", "portal"]
        for k, v in zip(keys, args.idcols_bnk):
            CONFIG["banco"][k] = v

    # Ventana de días
    ventana = None if args.ventana_dias is None or args.ventana_dias < 0 else int(args.ventana_dias)

    # Lectura
    df_bnk = read_table(args.banco)
    df_res = read_table(args.reservas)

    reservas = build_reservas(df_res)
    movimientos = build_banco(df_bnk)

    df_match_1to1, df_match_combo, df_unmatched_banco, df_unmatched_reservas, df_log = reconcile(
        reservas, movimientos, tol_cent=int(args.tolerancia_centimos),
        max_k=int(args.max_k), cap=int(args.cap_grupo), ventana_dias=ventana
    )

    export_excel(args.out, df_match_1to1, df_match_combo, df_unmatched_banco, df_unmatched_reservas, df_log)
    print(f"✅ Conciliación completada. Archivo generado: {args.out}")


if __name__ == "__main__":
    main()
