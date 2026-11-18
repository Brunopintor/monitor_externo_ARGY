"""
Variables internacionales que no están en las APIs del BCRA.

- DXY (índice del dólar contra una canasta de 6 monedas: EUR, JPY, GBP, CAD, SEK, CHF).
  Fuente: Yahoo Finance, símbolo DX-Y.NYB (ICE). Es una API pública pero NO oficial y sin
  documentación: si cambia o se cae, `dxy()` devuelve una serie vacía y avisa, para que el
  resto de la base se actualice igual.
  Plan B si Yahoo deja de andar: FRED DTWEXBGS (índice amplio de la Fed; no es el DXY y
  sale con ~1 semana de rezago): https://fred.stlouisfed.org/graph/fredgraph.csv?id=DTWEXBGS

El real (BRL) NO está acá: sale de la API cambiaria del BCRA (ver MONEDAS en actualizar_base.py).
"""
from __future__ import annotations

import warnings
from datetime import datetime, timezone

import pandas as pd
import requests

YAHOO = "https://query1.finance.yahoo.com/v8/finance/chart/{simbolo}"


def yahoo_diario(simbolo: str, desde: str = "2003-01-01") -> pd.Series:
    """Cierre diario de un símbolo de Yahoo Finance, indexado por fecha local del mercado."""
    p1 = int(datetime.fromisoformat(desde).replace(tzinfo=timezone.utc).timestamp())
    p2 = int(datetime.now(timezone.utc).timestamp())
    r = requests.get(YAHOO.format(simbolo=simbolo),
                     params={"period1": p1, "period2": p2, "interval": "1d"},
                     headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
    r.raise_for_status()
    res = r.json()["chart"]["result"][0]
    tz = res["meta"].get("exchangeTimezoneName", "America/New_York")
    fechas = pd.to_datetime(res["timestamp"], unit="s", utc=True).tz_convert(tz).normalize().tz_localize(None)
    s = pd.Series(res["indicators"]["quote"][0]["close"], index=fechas, dtype=float).dropna()
    return s[~s.index.duplicated(keep="last")].sort_index()


def dxy(desde: str = "2003-01-01") -> pd.Series:
    try:
        s = yahoo_diario("DX-Y.NYB", desde)
    except Exception as e:  # noqa: BLE001 — una fuente no oficial no tiene que frenar la actualización
        warnings.warn(f"No se pudo bajar el DXY de Yahoo Finance: {e}")
        s = pd.Series(dtype=float, index=pd.DatetimeIndex([]))
    return s.rename("dxy")


if __name__ == "__main__":
    print(dxy("2026-10-01").tail())
