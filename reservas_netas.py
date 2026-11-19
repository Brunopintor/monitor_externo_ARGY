"""
Reservas netas, replicando la hoja RRII del Excel "(2024.03.06) Reservas y tipos de cambio - nuevo".

Insumos:
  - API BCRA (automático): brutas (id 1), encajes USD (1243), depósitos del Tesoro en ME (1266),
    A3500 (5), yuan / DEG / oro (API cambiaria).
  - insumos_reservas_netas.xlsx (manual): pasivos que no están online + parámetros.

Fórmulas (letra = columna de la hoja RRII original):
  K  exceso_deg        = max(stock_deg − limite_deg × USD/DEG, 0)
  P  swap_china_usd    = swap_china_cny × ($/CNY) / A3500
  D  netas_fmi         = brutas − K − P − bis − encajes − swap_usa − organismos − repos(+SEDESA)
  E  netas_ex_bopreal  = D − bopreal
  F  netas_ex_bopreal_tesoro = E − depósitos del Tesoro en USD
  M,N,O ajustes de valuación a tipos de cambio de programa (fecha_ref_tmu):
        oro × (precio_ref/precio_t) − oro,  ídem DEG y swap China (CNY por USD)
  G  netas_sin_eff_tmu = AH + M + N + O
  AH rin_sin_fmi_2025  = E − desembolsos FMI EFF 2025 (cada uno desde su fecha)
  H  liquidas          = brutas − stock_deg + swap_china_activado − oro − P − sedesa
  L  var_vs_meta_rin   = E − rin_base_metas
  AI control_netas     = D + encajes − encajes (3 días hábiles antes)
  AJ control_super_netas = AI − depósitos Tesoro − bopreal

Uso:
    python reservas_netas.py              # calcula desde 2024 y guarda data/reservas_netas.csv
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from bcra_api import BCRA, hoy_ar

AQUI = Path(__file__).parent
INSUMOS = AQUI / "insumos_reservas_netas.xlsx"
DATA = AQUI / "data"


def leer_insumos(path: Path = INSUMOS) -> tuple[pd.DataFrame, dict, pd.Series, pd.DataFrame]:
    pas = pd.read_excel(path, sheet_name="pasivos").dropna(subset=["fecha_desde"])
    pas = pas.set_index(pd.to_datetime(pas["fecha_desde"])).drop(columns=["fecha_desde", "nota"], errors="ignore")
    pas = pas.sort_index().astype(float)

    par = pd.read_excel(path, sheet_name="parametros").set_index("parametro")["valor"].to_dict()
    for k in ("fecha_ref_tmu", "fecha_base_metas", "fecha_desembolso_fmi_1", "fecha_desembolso_fmi_2"):
        par[k] = pd.Timestamp(par[k])

    oro = pd.read_excel(path, sheet_name="oro_manual").dropna()
    oro = oro.set_index(pd.to_datetime(oro["fecha_desde"]))["oro_usd"].astype(float)

    bop = pd.read_excel(path, sheet_name="bopreal_cronograma").dropna(subset=["fecha_pago"])
    bop = bop.set_index(pd.to_datetime(bop["fecha_pago"])).drop(columns=["fecha_pago", "total"], errors="ignore")
    return pas, par, oro, bop.fillna(0).sum(axis=1)


def datos_api(desde: str, ref: pd.Timestamp) -> pd.DataFrame:
    b = BCRA()
    ini = min(pd.Timestamp(desde), ref - pd.Timedelta(days=10)).date().isoformat()
    d = b.series({"brutas": 1, "encajes_usd": 1243, "dep_tesoro_me_ars": 1266, "a3500": 5}, desde=ini)
    cny = b.cotizacion("CNY", desde=ini)["tipoCotizacion"].rename("ars_por_cny")
    deg = b.cotizacion("XDR", desde=ini)["tipoPase"].rename("usd_por_deg")
    oro = b.cotizacion("XAU", desde=ini)["tipoPase"].rename("usd_por_onza")
    d = pd.concat([d, cny, deg, oro], axis=1, sort=True)
    return d[d.index <= pd.Timestamp(hoy_ar())]


def vencimientos_bopreal_12m(cron: pd.Series, fechas: pd.DatetimeIndex) -> pd.Series:
    """Pagos de BOPREAL entre el inicio del mes de cada fecha y 12 meses después (referencia)."""
    out = []
    for f in fechas:
        ini = f.to_period("M").to_timestamp()
        out.append(cron[(cron.index >= ini) & (cron.index < ini + pd.DateOffset(months=12))].sum())
    return pd.Series(out, index=fechas)


def calcular(desde: str = "2024-01-01", api: pd.DataFrame | None = None) -> pd.DataFrame:
    pas, par, oro_man, cron = leer_insumos()
    ref = par["fecha_ref_tmu"]
    api = datos_api(desde, ref) if api is None else api

    # precios de referencia (TMU): último dato disponible a la fecha de referencia
    r = api.ffill().loc[:ref].iloc[-1]
    cny_usd_ref = r["a3500"] / r["ars_por_cny"]

    d = api[api["brutas"].notna()].loc[desde:].copy()
    # mercado: los tipos de cambio se completan con el último dato si falta el del día
    for c in ["a3500", "ars_por_cny", "usd_por_deg", "usd_por_onza"]:
        d[c] = api[c].ffill().reindex(d.index)

    # pasivos manuales: rige el último valor cargado a la fecha
    d = d.join(pas.reindex(d.index.union(pas.index)).ffill().reindex(d.index))
    d[pas.columns] = d[pas.columns].fillna(0)

    d["dep_tesoro_usd"] = d["dep_tesoro_me_ars"] / d["a3500"]
    if str(par.get("oro_metodo", "onzas")).strip().lower() == "manual":
        d["oro"] = oro_man.reindex(d.index.union(oro_man.index)).ffill().reindex(d.index)
    else:
        d["oro"] = float(par["oro_onzas_mm"]) * d["usd_por_onza"]

    d["exceso_deg"] = np.maximum(d["stock_deg"] - par["limite_deg"] * d["usd_por_deg"], 0)          # K
    d["swap_china_usd"] = par["swap_china_cny"] * d["ars_por_cny"] / d["a3500"]                       # P
    d["netas_fmi"] = (d["brutas"] - d["exceso_deg"] - d["swap_china_usd"] - d["bis"] - d["encajes_usd"]
                      - d["swap_usa"] - d["organismos"] - d["repos"])                                 # D
    d["netas_ex_bopreal"] = d["netas_fmi"] - d["bopreal"]                                             # E
    d["netas_ex_bopreal_tesoro"] = d["netas_ex_bopreal"] - d["dep_tesoro_usd"]                        # F

    d["ajuste_oro_tmu"] = d["oro"] * r["usd_por_onza"] / d["usd_por_onza"] - d["oro"]                 # M
    d["ajuste_deg_tmu"] = d["stock_deg"] * r["usd_por_deg"] / d["usd_por_deg"] - d["stock_deg"]       # N
    d["ajuste_yuan_tmu"] = d["swap_china_usd"] * cny_usd_ref / (d["a3500"] / d["ars_por_cny"]) - d["swap_china_usd"]  # O
    fmi = (par["desembolso_fmi_1"] * (d.index >= par["fecha_desembolso_fmi_1"])
           + par["desembolso_fmi_2"] * (d.index >= par["fecha_desembolso_fmi_2"]))
    d["rin_sin_fmi_2025"] = d["netas_ex_bopreal"] - fmi                                               # AH
    d["netas_sin_eff_tmu"] = d["rin_sin_fmi_2025"] + d["ajuste_oro_tmu"] + d["ajuste_deg_tmu"] + d["ajuste_yuan_tmu"]  # G

    d["liquidas"] = (d["brutas"] - d["stock_deg"] + par["swap_china_activado_usd"] - d["oro"]
                     - d["swap_china_usd"] - d["sedesa"])                                             # H
    d["var_vs_meta_rin"] = d["netas_ex_bopreal"] - par["rin_base_metas"]                              # L
    d["control_netas"] = d["netas_fmi"] + d["encajes_usd"] - d["encajes_usd"].shift(3)               # AI
    d["control_super_netas"] = d["control_netas"] - d["dep_tesoro_usd"] - d["bopreal"]               # AJ
    d["bopreal_venc_12m_auto"] = vencimientos_bopreal_12m(cron, d.index)                              # referencia
    d.index.name = "fecha"
    return d


SALIDA = ["brutas", "netas_fmi", "netas_ex_bopreal", "netas_ex_bopreal_tesoro", "netas_sin_eff_tmu",
          "rin_sin_fmi_2025", "liquidas", "var_vs_meta_rin", "control_netas", "control_super_netas",
          "exceso_deg", "swap_china_usd", "encajes_usd", "dep_tesoro_usd", "oro", "stock_deg", "swap_usa",
          "bis", "sedesa", "bopreal", "bopreal_venc_12m_auto", "organismos", "repos",
          "ajuste_oro_tmu", "ajuste_deg_tmu", "ajuste_yuan_tmu"]

if __name__ == "__main__":
    DATA.mkdir(exist_ok=True)
    rn = calcular()[SALIDA]
    rn.to_csv(DATA / "reservas_netas.csv", float_format="%.2f")
    print(rn.tail(5)[["brutas", "netas_fmi", "netas_ex_bopreal", "netas_ex_bopreal_tesoro", "liquidas"]].round(0))
