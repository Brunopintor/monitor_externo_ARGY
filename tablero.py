"""
Arma el tablero HTML (docs/index.html) a partir de lo que dejó actualizar_base.py en data/.

    python actualizar_base.py      # primero, baja y calcula todo
    python tablero.py              # después, arma el tablero

No llama a ninguna API: sólo lee data/base_diaria.csv y data/reservas_netas.csv,
prepara los datos en JSON y los inserta en tablero_plantilla.html (ECharts).
También genera la planilla de descarga (docs/monitor_externo_ARGY.xlsx, ver descarga_excel.py).
"""
from __future__ import annotations

import json
import math
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

import descarga_excel

AQUI = Path(__file__).parent
DATA = AQUI / "data"
PLANTILLA = AQUI / "tablero_plantilla.html"
SALIDA = AQUI / "docs" / "index.html"
MARCA = "/*__DATOS__*/null"

# Dólares que se grafican y van a la tablita: columna -> etiqueta.
# El orden fija el color de cada dólar en todos los gráficos (ver COLOR en la plantilla).
DOLARES = {
    "a3500": "Oficial (A3500)",
    "mep_venta": "MEP",
    "ccl_venta": "CCL",
    "blue_venta": "Blue",
    "oficial_minorista": "Minorista",
    "tarjeta_venta": "Tarjeta",
    "cripto_venta": "Cripto",
}
BRECHAS = {"brecha_mep": "MEP", "brecha_ccl": "CCL", "brecha_blue": "Blue"}
NETAS = {
    "netas_fmi": "RIN 1",
    "netas_ex_bopreal": "RIN 2",
    "netas_ex_bopreal_tesoro": "RIN 3",
}
NOTA_RIN = "RIN: reservas internacionales netas."
# Notas al pie de cada medida de reservas netas (van con asterisco en los boxes y bajo el gráfico)
NOTAS_NETAS = {
    "netas_fmi": "Metodología FMI: reservas brutas menos pasivos de corto plazo en moneda extranjera.",
    "netas_ex_bopreal": "RIN 1 menos vencimientos de capital e intereses del bono BOPREAL "
                        "en los próximos 12 meses.",
    "netas_ex_bopreal_tesoro": "RIN 2 menos los depósitos en dólares del Tesoro Nacional en el BCRA.",
}

# Color de los boxes y de las variaciones según qué significa que la variable suba:
#   "sube_malo"  -> suba en rojo, baja en verde (depreciación, brecha, dólar global fuerte)
#   "sube_bueno" -> suba en verde, baja en rojo (reservas)
#   "neutro"     -> gris siempre
# Para cambiar el criterio de una variable alcanza con cambiar la palabra acá.
SENTIDO = {
    "a3500": "sube_malo", "mep_venta": "sube_malo", "ccl_venta": "sube_malo", "blue_venta": "sube_malo",
    "oficial_minorista": "sube_malo", "tarjeta_venta": "sube_malo", "cripto_venta": "sube_malo",
    "brecha_ccl": "sube_malo", "posicion_en_banda": "sube_malo",
    "dxy": "sube_malo", "brl_por_usd": "sube_malo",
    "reservas_brutas": "sube_bueno", "netas_fmi": "sube_bueno", "netas_ex_bopreal": "sube_bueno",
    "netas_ex_bopreal_tesoro": "sube_bueno",
}
FACTORES = {
    "var_rrii_compra_divisas": "Compra de divisas",
    "var_rrii_org_internac": "Organismos internacionales",
    "var_rrii_sector_publico": "Sector público",
    "var_rrii_efectivo_minimo": "Efectivo mínimo",
    "var_rrii_otros": "Otros (incl. valuación)",
}
# Pasos de brutas a netas (hoja RRII): columna -> etiqueta. Todos restan.
PASOS_NETAS_FMI = {
    "exceso_deg": "Exceso DEG",
    "swap_china_usd": "Swap China",
    "bis": "BIS",
    "encajes_usd": "Encajes",
    "swap_usa": "Swap EE.UU.",
    "organismos": "Organismos",
    "repos": "REPOs + SEDESA",
}


def leer() -> tuple[pd.DataFrame, pd.DataFrame]:
    base = pd.read_csv(DATA / "base_diaria.csv", index_col=0, parse_dates=True)
    rn = pd.read_csv(DATA / "reservas_netas.csv", index_col=0, parse_dates=True)
    return base, rn


def _num(x, dec: int = 4):
    if x is None or (isinstance(x, float) and math.isnan(x)) or pd.isna(x):
        return None
    return round(float(x), dec) + 0.0   # + 0.0 evita "-0.0"


def serie_columnar(df: pd.DataFrame, cols: list[str], dec: int = 4) -> dict:
    """{'fechas': [...], 'col': [...]} sin filas totalmente vacías."""
    d = df[cols].dropna(how="all")
    out = {"fechas": [f.strftime("%Y-%m-%d") for f in d.index]}
    for c in cols:
        out[c] = [_num(v, dec) for v in d[c]]
    return out


def ultimo(s: pd.Series) -> tuple[pd.Timestamp | None, float | None]:
    s = s.dropna()
    if s.empty:
        return None, None
    return s.index[-1], float(s.iloc[-1])


def valor_al(s: pd.Series, fecha: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    """Último dato disponible en o antes de `fecha`."""
    return ultimo(s.loc[:fecha])


def variaciones(s: pd.Series) -> dict:
    """Último valor, variación contra el dato anterior y contra un mes antes."""
    f, v = ultimo(s)
    if f is None:
        return {}
    s = s.dropna()
    ant = s.iloc[-2] if len(s) > 1 else None
    _, v_mes = valor_al(s, f - pd.DateOffset(months=1))
    return {
        "fecha": f.strftime("%Y-%m-%d"),
        "valor": _num(v),
        "var_dia": _num(v / ant - 1, 6) if ant else None,
        "var_mes": _num(v / v_mes - 1, 6) if v_mes else None,
    }


# KPIs de la pestaña Resumen, por grupo: (columna, etiqueta, tipo, ayuda).
# tipo define el formato y cómo se mide la variación:
#   ars / indice / brl -> variación %   ·   pct / banda -> puntos porcentuales   ·   mm -> millones de USD
KPIS = {
    "Dólar en la Argentina": [
        ("a3500", "Oficial A3500", "ars", None),
        ("mep_venta", "MEP", "ars", None),
        ("ccl_venta", "CCL", "ars", None),
        ("blue_venta", "Blue", "ars", None),
        ("brecha_ccl", "Brecha CCL", "pct", "CCL / A3500 − 1"),
        ("posicion_en_banda", "Posición en la banda", "banda", "0 % = piso · 100 % = techo"),
    ],
    "Dólar en el mundo": [
        ("dxy", "DXY", "indice", "Índice del dólar contra euro, yen, libra, dólar canadiense, corona sueca y franco suizo"),
        ("brl_por_usd", "Real por dólar", "brl", "Reales brasileños por dólar (BCRA)"),
    ],
    "Reservas (MM USD)": [
        ("reservas_brutas", "Reservas Brutas", "mm", None),
        ("netas_fmi", NETAS["netas_fmi"], "mm", None),
        ("netas_ex_bopreal", NETAS["netas_ex_bopreal"], "mm", None),
        ("netas_ex_bopreal_tesoro", NETAS["netas_ex_bopreal_tesoro"], "mm", None),
    ],
}
DIAS_SPARK = 66   # ~3 meses hábiles para la minicurva de cada KPI


def kpis(base: pd.DataFrame) -> list[dict]:
    grupos = []
    for grupo, items in KPIS.items():
        out = []
        for col, etiqueta, tipo, ayuda in items:
            s = base[col].dropna() if col in base else pd.Series(dtype=float)
            if s.empty:
                continue
            f, v = s.index[-1], s.iloc[-1]
            _, v_mes = valor_al(s, f - pd.DateOffset(months=1))
            relativa = tipo in ("ars", "indice", "brl")
            dif = (lambda a, b: a / b - 1) if relativa else (lambda a, b: a - b)
            out.append({
                "id": col, "etiqueta": etiqueta, "tipo": tipo, "ayuda": ayuda,
                "sentido": SENTIDO.get(col, "neutro"), "nota": NOTAS_NETAS.get(col),
                "fecha": f.strftime("%Y-%m-%d"), "valor": _num(v, 6),
                "var_dia": _num(dif(v, s.iloc[-2]), 6) if len(s) > 1 else None,
                "var_mes": _num(dif(v, v_mes), 6) if v_mes is not None else None,
                "spark": [_num(x, 4) for x in s.iloc[-DIAS_SPARK:]],
            })
        if out:
            grupos.append({"grupo": grupo, "items": out})
    return grupos


def tabla_dolares(base: pd.DataFrame) -> list[dict]:
    filas = []
    for col, etq in DOLARES.items():
        v = variaciones(base[col])
        if v:
            filas.append({"id": col, "etiqueta": etq, "sentido": SENTIDO.get(col, "neutro"), **v})
    return filas


def waterfall_mensual(base: pd.DataFrame) -> dict:
    """Movimiento de las brutas en el último mes, abierto por factor de explicación.

    Ventana: (último dato de brutas − 1 mes, último dato de brutas]. Los días que todavía
    no tienen factores publicados quedan en "Sin desagregar" (la diferencia contra el total).
    """
    brutas = base["reservas_brutas"].dropna()
    f_fin, v_fin = ultimo(brutas)
    f_ini, v_ini = valor_al(brutas, f_fin - pd.DateOffset(months=1))
    ventana = base.loc[(base.index > f_ini) & (base.index <= f_fin)]
    pasos = [{"etiqueta": etq, "valor": _num(ventana[col].sum(min_count=1) or 0.0, 1)}
             for col, etq in FACTORES.items()]
    resto = (v_fin - v_ini) - sum(p["valor"] for p in pasos)
    if abs(resto) >= 1:
        f_fact = ventana[list(FACTORES)].dropna(how="all").index.max()
        pasos.append({"etiqueta": "Sin desagregar", "valor": _num(resto, 1),
                      "nota": f"Días posteriores al {f_fact:%d/%m} sin factores publicados todavía"
                      if pd.notna(f_fact) else None})
    return {"desde": f_ini.strftime("%Y-%m-%d"), "hasta": f_fin.strftime("%Y-%m-%d"),
            "inicio": _num(v_ini, 1), "fin": _num(v_fin, 1), "pasos": pasos}


def composicion_netas(rn: pd.DataFrame) -> dict:
    """De brutas a netas en la última fecha con netas calculadas."""
    f = rn["netas_fmi"].last_valid_index()
    r = rn.loc[f]
    pasos = [{"etiqueta": etq, "valor": _num(-r[col], 1)} for col, etq in PASOS_NETAS_FMI.items()]
    # Control: los pasos tienen que cerrar con netas_fmi (si no, falta un componente)
    cierre = r["brutas"] + sum(p["valor"] for p in pasos) - r["netas_fmi"]
    if abs(cierre) > 1:
        raise ValueError(f"La composición no cierra con netas_fmi el {f:%d/%m/%Y}: dif {cierre:.1f} MM")
    return {
        "fecha": f.strftime("%Y-%m-%d"),
        "brutas": _num(r["brutas"], 1),
        "pasos_fmi": pasos,
        "netas_fmi": _num(r["netas_fmi"], 1),
        # lo que efectivamente se resta (vencimientos 12m o stock, según bopreal_metodo)
        "bopreal": _num(-r.get("bopreal_restado", r["bopreal"]), 1),
        "netas_ex_bopreal": _num(r["netas_ex_bopreal"], 1),
        "tesoro": _num(-r["dep_tesoro_usd"], 1),
        "netas_ex_bopreal_tesoro": _num(r["netas_ex_bopreal_tesoro"], 1),
    }


def armar_datos() -> dict:
    base, rn = leer()
    ahora = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
    return {
        "generado": ahora.strftime("%Y-%m-%dT%H:%M"),
        "etiquetas": {**DOLARES, **BRECHAS, **NETAS, "reservas_brutas": "Reservas Brutas",
                      "banda_inferior": "Banda inferior", "banda_superior": "Banda superior"},
        "kpis": kpis(base),
        "tabla_dolares": tabla_dolares(base),
        "tc": serie_columnar(base, list(DOLARES) + ["banda_inferior", "banda_superior"], 2),
        "brechas": serie_columnar(base, list(BRECHAS), 5),
        "reservas": serie_columnar(base, ["reservas_brutas"] + list(NETAS), 1),
        "waterfall_mes": waterfall_mensual(base),
        "composicion": composicion_netas(rn),
        "notas_netas": [NOTAS_NETAS[k] for k in NETAS],
        "nota_rin": NOTA_RIN,
        "excel": descarga_excel.NOMBRE,
    }


def main():
    datos = armar_datos()
    html = PLANTILLA.read_text(encoding="utf-8")
    if MARCA not in html:
        raise ValueError(f"La plantilla no tiene la marca {MARCA}")
    js = json.dumps(datos, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    SALIDA.parent.mkdir(exist_ok=True)
    SALIDA.write_text(html.replace(MARCA, js), encoding="utf-8")
    print(f"OK: {SALIDA} ({SALIDA.stat().st_size / 1e6:.2f} MB)")
    xlsx = descarga_excel.generar(*leer(), SALIDA.parent / descarga_excel.NOMBRE)
    print(f"OK: {xlsx} ({xlsx.stat().st_size / 1e6:.2f} MB)")


if __name__ == "__main__":
    main()
