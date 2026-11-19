"""
Cliente reutilizable para las APIs públicas del BCRA.

Dos APIs, misma raíz (https://api.bcra.gob.ar), sin autenticación ni token:

1) Estadísticas monetarias v4.0  ->  /estadisticas/v4.0/monetarias
   - GET /monetarias                  catálogo: ~1.600 variables (id, descripción, categoría, unidad,
                                      periodicidad, primera/última fecha, último valor)
   - GET /monetarias/{id}?desde=&hasta=&limit=&offset=
                                      serie de una variable. limit máx. 3000.
   Respuesta: {"status":200,"metadata":{"resultset":{"count":N,...}},
               "results":[{"idVariable":5,"detalle":[{"fecha":"2026-11-17","valor":1552.26}, ...]}]}

2) Estadísticas cambiarias v1.0  ->  /estadisticascambiarias/v1.0
   - GET /Maestros/Divisas            lista de monedas (USD, CNY, XDR=DEG, XAU=oro, REF=Com. A3500...)
   - GET /Cotizaciones?fecha=AAAA-MM-DD
                                      todas las monedas de un día
   - GET /Cotizaciones/{moneda}?fechadesde=&fechahasta=&limit=&offset=
                                      serie de una moneda. limit entre 10 y 1000.
   Cada registro trae:
     tipoPase        -> dólares por unidad de la moneda (ej. CNY 0.15 USD; XAU 3.862 USD por onza)
     tipoCotizacion  -> pesos por unidad de la moneda (Com. A3500 para USD)

Uso rápido:
    from bcra_api import BCRA
    b = BCRA()
    b.buscar("banda")                       # busca en el catálogo
    b.serie(5, desde="2024-01-01")          # A3500 como pd.Series
    b.series({"a3500": 5, "brutas": 1})     # varias juntas en un DataFrame
    b.cotizacion("CNY", desde="2024-01-01") # DataFrame con tipoPase y tipoCotizacion
"""
from __future__ import annotations

import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

BASE = "https://api.bcra.gob.ar"
MON = f"{BASE}/estadisticas/v4.0/monetarias"
CAM = f"{BASE}/estadisticascambiarias/v1.0"



def hoy_ar() -> date:
    """Fecha de hoy en la Argentina.

    No usar date.today(): en GitHub Actions el servidor está en UTC y desde las 21 h de
    Argentina ya es "mañana"; la API cambiaria rechaza con 400 una fecha posterior a hoy.
    """
    return datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()

class BCRA:
    def __init__(self, timeout: int = 30, reintentos: int = 3, verify=True):
        # verify: en algunas PCs con Windows la cadena de certificados del BCRA falla;
        # en ese caso pasar la ruta a un bundle (ej. certifi.where()) en vez de desactivarlo.
        self.s = requests.Session()
        self.s.headers["Accept"] = "application/json"
        self.timeout, self.reintentos, self.verify = timeout, reintentos, verify
        self._catalogo: pd.DataFrame | None = None

    # ------------------------------------------------------------------ http
    def _get(self, url: str, params: dict | None = None) -> dict:
        for i in range(self.reintentos):
            try:
                r = self.s.get(url, params=params, timeout=self.timeout, verify=self.verify)
                if r.status_code == 404:          # sin datos para ese rango
                    return {"results": [], "metadata": {"resultset": {"count": 0}}}
                r.raise_for_status()
                return r.json()
            except requests.RequestException:
                if i == self.reintentos - 1:
                    raise
                time.sleep(2 * (i + 1))

    # ------------------------------------------------------ monetarias v4
    def catalogo(self, refrescar: bool = False) -> pd.DataFrame:
        """Todas las variables disponibles (≈1.600)."""
        if self._catalogo is None or refrescar:
            d = self._get(MON, {"limit": 3000})
            self._catalogo = pd.DataFrame(d["results"]).set_index("idVariable")
        return self._catalogo

    def buscar(self, texto: str) -> pd.DataFrame:
        """Busca variables por texto en la descripción (sin distinguir mayúsculas)."""
        c = self.catalogo()
        m = c["descripcion"].str.contains(texto, case=False, na=False)
        return c.loc[m, ["descripcion", "categoria", "unidadExpresion", "periodicidad",
                         "primerFechaInformada", "ultFechaInformada", "ultValorInformado"]]

    def serie(self, id_variable: int, desde: str = "2003-01-01", hasta: str | None = None,
              nombre: str | None = None) -> pd.Series:
        """Serie completa de una variable, paginando de a 3000 registros."""
        hasta = hasta or (hoy_ar() + timedelta(days=60)).isoformat()  # bandas vienen pre-publicadas
        filas, offset = [], 0
        while True:
            d = self._get(f"{MON}/{id_variable}",
                          {"desde": desde, "hasta": hasta, "limit": 3000, "offset": offset})
            res = d.get("results") or []
            det = res[0]["detalle"] if res else []
            filas += det
            total = d["metadata"]["resultset"]["count"]
            offset += 3000
            if not det or offset >= total:
                break
        s = pd.Series({pd.Timestamp(x["fecha"]): x["valor"] for x in filas},
                      name=nombre or id_variable, dtype="float64")
        if s.empty:   # índice de fechas aunque no haya datos, para poder juntarla con otras
            s.index = pd.DatetimeIndex([], name="fecha")
        return s.sort_index()

    def series(self, ids: dict[str, int], **kw) -> pd.DataFrame:
        """Varias variables en un DataFrame, columnas = claves del dict."""
        return pd.concat([self.serie(i, nombre=k, **kw) for k, i in ids.items()], axis=1)

    # ------------------------------------------------------ cambiarias v1
    def divisas(self) -> pd.DataFrame:
        return pd.DataFrame(self._get(f"{CAM}/Maestros/Divisas")["results"])

    def cotizaciones_dia(self, fecha: str) -> pd.DataFrame:
        r = self._get(f"{CAM}/Cotizaciones", {"fecha": fecha})["results"]
        return pd.DataFrame(r["detalle"]).assign(fecha=r.get("fecha")) if r else pd.DataFrame()

    def cotizacion(self, moneda: str, desde: str = "2003-01-01", hasta: str | None = None) -> pd.DataFrame:
        """Serie de una moneda: columnas tipoPase (USD por unidad) y tipoCotizacion ($ por unidad)."""
        hasta = hasta or hoy_ar().isoformat()   # la API rechaza (400) fechas posteriores a hoy en Argentina
        filas, offset = [], 0
        while True:
            d = self._get(f"{CAM}/Cotizaciones/{moneda}",
                          {"fechadesde": desde, "fechahasta": hasta, "limit": 1000, "offset": offset})
            res = d.get("results") or []
            for dia in res:
                for x in dia["detalle"]:
                    filas.append({"fecha": pd.Timestamp(dia["fecha"]),
                                  "tipoPase": x["tipoPase"], "tipoCotizacion": x["tipoCotizacion"]})
            total = d["metadata"]["resultset"]["count"]
            offset += 1000
            if not res or offset >= total:
                break
        if not filas:
            return pd.DataFrame(columns=["tipoPase", "tipoCotizacion"], dtype="float64",
                                index=pd.DatetimeIndex([], name="fecha"))
        return pd.DataFrame(filas).drop_duplicates("fecha").set_index("fecha").sort_index()


if __name__ == "__main__":
    b = BCRA()
    print(b.buscar("bandas"))
    print(b.serie(5, desde="2026-11-01").tail())
    print(b.cotizacion("CNY", desde="2026-11-01").tail())
