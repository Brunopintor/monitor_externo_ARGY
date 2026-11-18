"""
Dólares paralelos y financieros (MEP, CCL, blue, tarjeta, cripto).

No hay fuente oficial con API, pero hay dos APIs comunitarias, gratuitas y sin token,
que ya hacen el scraping por nosotros:

- ArgentinaDatos  https://api.argentinadatos.com/v1/cotizaciones/dolares/{casa}
    HISTÓRICO diario completo (blue/oficial/mayorista desde 2011, CCL desde 2013, MEP desde 2018).
    casas: oficial, blue, bolsa (=MEP), contadoconliqui (=CCL), mayorista, tarjeta, cripto, solidario
- DolarApi        https://dolarapi.com/v1/dolares
    Cotización EN VIVO (se actualiza cada pocos minutos). Útil para el valor intradiario/cierre.

Ambas son proyectos open source del mismo ecosistema; toman los precios de fuentes
periodísticas/mercado. No son oficiales: conviene validar contra la fuente que se use
hoy en el Excel y dejar registro de cuál se usa.
"""
from __future__ import annotations

import pandas as pd
import requests

AD = "https://api.argentinadatos.com/v1/cotizaciones/dolares"
DOLARAPI = "https://dolarapi.com/v1/dolares"

CASAS = {"mep": "bolsa", "ccl": "contadoconliqui", "blue": "blue", "tarjeta": "tarjeta", "cripto": "cripto"}


def historico(casa: str) -> pd.DataFrame:
    """Histórico diario de una casa: columnas compra, venta."""
    d = pd.DataFrame(requests.get(f"{AD}/{casa}", timeout=60).json())
    d["fecha"] = pd.to_datetime(d["fecha"])
    return d.set_index("fecha")[["compra", "venta"]].astype(float).sort_index()


def historicos(casas: dict[str, str] = CASAS) -> pd.DataFrame:
    """Todas las casas en un DataFrame ancho: mep_compra, mep_venta, ccl_venta, ..."""
    out = []
    for nombre, casa in casas.items():
        h = historico(casa)
        h.columns = [f"{nombre}_{c}" for c in h.columns]
        out.append(h)
    return pd.concat(out, axis=1)


def en_vivo() -> pd.DataFrame:
    """Cotización actual de todas las casas (DolarApi)."""
    d = pd.DataFrame(requests.get(DOLARAPI, timeout=30).json())
    d["fechaActualizacion"] = pd.to_datetime(d["fechaActualizacion"]).dt.tz_convert("America/Argentina/Buenos_Aires")
    return d.set_index("casa")[["compra", "venta", "fechaActualizacion"]]


if __name__ == "__main__":
    print(en_vivo())
    print(historicos().tail())
