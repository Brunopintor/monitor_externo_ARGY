"""
Arma la base diaria de tipos de cambio, bandas y reservas (equivalente a la hoja BASE
del Excel "(2024.03.06) Reservas y tipos de cambio - nuevo") y la guarda en data/.

    python actualizar_base.py                 # base completa desde 2003
    python actualizar_base.py --desde 2024-01-01

Salidas:
    data/base_diaria.csv     una fila por día hábil, todas las series + brechas + variaciones
    data/base_diaria.xlsx    lo mismo, para abrir en Excel / Power BI
    data/reservas_netas.csv  reservas netas y todos sus componentes (desde 2024)
    data/ultimo_dato.json    resumen del último dato de cada serie (para el tablero)
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

import pandas as pd

from bcra_api import BCRA
import internacionales
import paralelos
import reservas_netas

# ---- Series de la API monetaria (id -> nombre de columna). Buscar más con BCRA().buscar("texto")
SERIES_MONETARIAS = {
    "a3500": 5,                    # Tipo de cambio mayorista de referencia (Com. A3500)
    "oficial_minorista": 4,        # Tipo de cambio minorista, promedio vendedor
    "banda_inferior": 1187,        # Régimen de bandas cambiarias. Límite inferior
    "banda_superior": 1188,        # Régimen de bandas cambiarias. Límite superior
    "reservas_brutas": 1,          # Reservas internacionales, millones USD
    "encajes_usd": 1243,           # Cuentas corrientes en moneda extranjera en el BCRA, millones USD
    "dep_tesoro_me_ars": 1266,     # Depósitos del gobierno en el BCRA en moneda extranjera, millones ARS
    "var_rrii_compra_divisas": 78, # Factores de variación de reservas (millones USD)
    "var_rrii_org_internac": 79,
    "var_rrii_sector_publico": 80,
    "var_rrii_efectivo_minimo": 81,
    "var_rrii_otros": 82,
}

# ---- Series de la API cambiaria: moneda -> (columna tipoPase, columna tipoCotizacion)
MONEDAS = {
    "CNY": ("usd_por_cny", "ars_por_cny"),
    "XDR": ("usd_por_deg", "ars_por_deg"),
    "XAU": ("usd_por_onza_oro", None),
    "BRL": ("usd_por_brl", "ars_por_brl"),
}

DATA = Path(__file__).parent / "data"


def construir(desde: str = "2003-01-01") -> pd.DataFrame:
    b = BCRA()

    print("BCRA monetarias...")
    mon = b.series(SERIES_MONETARIAS, desde=desde)

    print("BCRA cambiarias...")
    cam = []
    for moneda, (col_pase, col_cot) in MONEDAS.items():
        c = b.cotizacion(moneda, desde=desde)
        cols = {"tipoPase": col_pase, "tipoCotizacion": col_cot}
        c = c.rename(columns=cols)[[v for v in cols.values() if v]]
        cam.append(c)
    cam = pd.concat(cam, axis=1, sort=True)

    print("Paralelos (ArgentinaDatos)...")
    par = paralelos.historicos()
    par = par[["mep_venta", "ccl_venta", "blue_compra", "blue_venta", "tarjeta_venta", "cripto_venta"]]

    print("DXY (Yahoo Finance)...")
    dxy = internacionales.dxy(desde)

    df = pd.concat([mon, cam, par, dxy], axis=1, sort=True)
    df = df.loc[desde:pd.Timestamp(date.today())]          # las bandas vienen publicadas a futuro
    # Días hábiles: ArgentinaDatos repite el último valor en fines de semana y feriados.
    # Se dejan lunes a viernes, y se sacan los feriados (días sin A3500 anteriores al último A3500).
    # El día de hoy queda aunque el BCRA todavía no haya publicado el A3500.
    ult_a3500 = df["a3500"].last_valid_index()
    feriado = df["a3500"].isna() & (df.index <= ult_a3500)
    df = df[(df.index.dayofweek < 5) & ~feriado]
    df.index.name = "fecha"

    # ---- Derivadas
    df["dep_tesoro_me_usd"] = df["dep_tesoro_me_ars"] / df["a3500"]
    for c in ["mep_venta", "ccl_venta", "blue_venta"]:
        df[f"brecha_{c.split('_')[0]}"] = df[c] / df["a3500"] - 1
    df["brl_por_usd"] = 1 / df["usd_por_brl"]                     # reales por dólar
    df["canje_ccl_mep"] = df["ccl_venta"] / df["mep_venta"] - 1
    df["dist_banda_sup"] = df["banda_superior"] / df["a3500"] - 1   # cuánto falta para el techo
    df["posicion_en_banda"] = (df["a3500"] - df["banda_inferior"]) / (df["banda_superior"] - df["banda_inferior"])
    for c in ["a3500", "oficial_minorista", "mep_venta", "ccl_venta", "blue_venta"]:
        s = df[c].dropna()
        df[f"var_d_{c}"] = (s / s.shift(1) - 1).reindex(df.index)
    df["var_d_reservas_brutas"] = df["reservas_brutas"].dropna().diff().reindex(df.index)
    return df


def resumen(df: pd.DataFrame) -> dict:
    out = {}
    for c in df.columns:
        s = df[c].dropna()
        if len(s):
            out[c] = {"fecha": s.index[-1].date().isoformat(), "valor": round(float(s.iloc[-1]), 6)}
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", default="2003-01-01")
    a = ap.parse_args()
    DATA.mkdir(exist_ok=True)
    df = construir(a.desde)
    print("Reservas netas (API + insumos_reservas_netas.xlsx)...")
    rn = reservas_netas.calcular(desde="2024-01-01")[reservas_netas.SALIDA]
    rn.to_csv(DATA / "reservas_netas.csv", float_format="%.2f")
    netas = ["netas_fmi", "netas_ex_bopreal", "netas_ex_bopreal_tesoro", "liquidas", "var_vs_meta_rin"]
    df = df.join(rn[netas])
    df.to_csv(DATA / "base_diaria.csv", float_format="%.6f")
    df.to_excel(DATA / "base_diaria.xlsx", freeze_panes=(1, 1))
    (DATA / "ultimo_dato.json").write_text(json.dumps(resumen(df), indent=1, ensure_ascii=False))
    print(f"OK: {len(df)} filas, {df.index.min().date()} a {df.index.max().date()}")
    print(df.iloc[-3:, :8].T)
