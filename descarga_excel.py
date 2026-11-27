"""
Planilla de descarga del tablero: docs/monitor_externo_ARGY.xlsx (botón "Excel" del tablero).

Dos hojas, Dólares y Reservas, con el formato de una base de series:
    fila 1: nombre de la variable · fila 2: unidad y fuente · desde la fila 3: datos
    columna A: fecha · una serie por columna

Se arma en cada corrida desde data/ (no llama a ninguna API). No se versiona en git
(está en .gitignore): el workflow la genera y la publica junto con el tablero.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NOMBRE = "monitor_externo_ARGY.xlsx"
FUENTE_LETRA = "Aptos"

ARS, PCT, MM, MM1, IDX, BRL = "#,##0.00", "0.0%", "#,##0", "#,##0.0", "0.00", "0.0000"

# (columna en data/, nombre, unidad, fuente, formato)
HOJA_DOLARES = [
    ("a3500", "Oficial mayorista (A3500)", "$ por USD", "BCRA", ARS),
    ("oficial_minorista", "Oficial minorista (promedio vendedor)", "$ por USD", "BCRA", ARS),
    ("mep_venta", "MEP (venta)", "$ por USD", "ArgentinaDatos / DolarApi", ARS),
    ("ccl_venta", "CCL (venta)", "$ por USD", "ArgentinaDatos / DolarApi", ARS),
    ("blue_venta", "Blue (venta)", "$ por USD", "ArgentinaDatos / DolarApi", ARS),
    ("tarjeta_venta", "Tarjeta (venta)", "$ por USD", "ArgentinaDatos / DolarApi", ARS),
    ("cripto_venta", "Cripto (venta)", "$ por USD", "ArgentinaDatos / DolarApi", ARS),
    ("banda_inferior", "Banda cambiaria: piso", "$ por USD", "BCRA", ARS),
    ("banda_superior", "Banda cambiaria: techo", "$ por USD", "BCRA", ARS),
    ("brecha_mep", "Brecha MEP", "% sobre el A3500", "Cálculo propio", PCT),
    ("brecha_ccl", "Brecha CCL", "% sobre el A3500", "Cálculo propio", PCT),
    ("brecha_blue", "Brecha blue", "% sobre el A3500", "Cálculo propio", PCT),
    ("posicion_en_banda", "Posición en la banda", "0 % = piso, 100 % = techo", "Cálculo propio", PCT),
    ("dxy", "DXY (dólar global)", "Índice", "Yahoo Finance (ICE)", IDX),
    ("brl_por_usd", "Real por dólar", "BRL por USD", "BCRA", BRL),
]

HOJA_RESERVAS = [
    ("reservas_brutas", "Reservas Brutas", "MM USD", "BCRA", MM),
    ("netas_fmi", "Reservas Netas 1 (metodología FMI)", "MM USD", "Cálculo propio", MM),
    ("netas_ex_bopreal", "Reservas Netas 2 (Netas 1 − venc. BOPREAL 12 meses)", "MM USD", "Cálculo propio", MM),
    ("netas_ex_bopreal_tesoro", "Reservas Netas 3 (Netas 2 − depósitos del Tesoro)", "MM USD", "Cálculo propio", MM),
    ("exceso_deg", "Exceso de DEG", "MM USD", "Cálculo propio", MM1),
    ("swap_china_usd", "Swap con China", "MM USD", "Cálculo propio (BCRA)", MM1),
    ("bis", "BIS", "MM USD", "Carga manual", MM1),
    ("encajes_usd", "Encajes en USD", "MM USD", "BCRA", MM1),
    ("swap_usa", "Swap con EE.UU.", "MM USD", "Carga manual", MM1),
    ("organismos", "Organismos internacionales", "MM USD", "Carga manual", MM1),
    ("repos", "REPOs (incl. SEDESA)", "MM USD", "Carga manual", MM1),
    ("bopreal", "BOPREAL: stock", "MM USD", "Carga manual", MM1),
    ("bopreal_venc_12m", "BOPREAL: vencimientos próximos 12 meses", "MM USD", "Cálculo propio (cronograma)", MM1),
    ("dep_tesoro_usd", "Depósitos del Tesoro en USD en el BCRA", "MM USD", "BCRA", MM1),
    ("var_rrii_compra_divisas", "Var. diaria: compra de divisas", "MM USD", "BCRA", MM1),
    ("var_rrii_org_internac", "Var. diaria: organismos internacionales", "MM USD", "BCRA", MM1),
    ("var_rrii_sector_publico", "Var. diaria: sector público", "MM USD", "BCRA", MM1),
    ("var_rrii_efectivo_minimo", "Var. diaria: efectivo mínimo", "MM USD", "BCRA", MM1),
    ("var_rrii_otros", "Var. diaria: otros (incl. valuación)", "MM USD", "BCRA", MM1),
]

_F_NOMBRE = Font(name=FUENTE_LETRA, bold=True, size=10, color="FFFFFF")
_F_UNIDAD = Font(name=FUENTE_LETRA, italic=True, size=9, color="52514E")
_F_DATO = Font(name=FUENTE_LETRA, size=10)
_RELLENO_NOMBRE = PatternFill("solid", fgColor="1C5CAB")
_RELLENO_UNIDAD = PatternFill("solid", fgColor="EEF4FC")
_BORDE = Border(bottom=Side(style="thin", color="C3C2B7"))
_AJUSTE = Alignment(wrap_text=True, vertical="center", horizontal="center")


def _hoja(wb: Workbook, titulo: str, df: pd.DataFrame, columnas: list[tuple]) -> None:
    columnas = [c for c in columnas if c[0] in df.columns]
    datos = df[[c[0] for c in columnas]].dropna(how="all")
    ws = wb.create_sheet(titulo)

    ws.cell(row=1, column=1, value="Fecha")
    ws.cell(row=2, column=1, value="dd/mm/aaaa")
    for j, (_, nombre, unidad, fuente, _) in enumerate(columnas, start=2):
        ws.cell(row=1, column=j, value=nombre)
        ws.cell(row=2, column=j, value=f"{unidad} · {fuente}")
    for j in range(1, len(columnas) + 2):
        c1, c2 = ws.cell(row=1, column=j), ws.cell(row=2, column=j)
        c1.font, c1.fill, c1.alignment = _F_NOMBRE, _RELLENO_NOMBRE, _AJUSTE
        c2.font, c2.fill, c2.alignment, c2.border = _F_UNIDAD, _RELLENO_UNIDAD, _AJUSTE, _BORDE
        ws.column_dimensions[get_column_letter(j)].width = 12 if j == 1 else 16
    ws.row_dimensions[1].height = 45
    ws.row_dimensions[2].height = 30

    formatos = [c[4] for c in columnas]
    for i, (fecha, fila) in enumerate(zip(datos.index, datos.itertuples(index=False)), start=3):
        c = ws.cell(row=i, column=1, value=fecha.to_pydatetime())
        c.number_format, c.font = "dd/mm/yyyy", _F_DATO
        for j, (v, fmt) in enumerate(zip(fila, formatos), start=2):
            if pd.notna(v):
                c = ws.cell(row=i, column=j, value=float(v))
                c.number_format, c.font = fmt, _F_DATO
    ws.freeze_panes = "B3"


def generar(base: pd.DataFrame, rn: pd.DataFrame, destino: Path) -> Path:
    """base = data/base_diaria.csv, rn = data/reservas_netas.csv (ambos indexados por fecha)."""
    wb = Workbook()
    wb.remove(wb.active)
    _hoja(wb, "Dólares", base, HOJA_DOLARES)
    # Reservas: brutas y factores desde la base (desde 2003); netas y componentes desde rn (desde 2024)
    res = base[["reservas_brutas"] + [c for c in base if c.startswith("var_rrii_")]].join(
        rn.drop(columns=["brutas"], errors="ignore"), how="outer")
    _hoja(wb, "Reservas", res, HOJA_RESERVAS)
    wb.properties.creator = "Bruno Pintor · monitor_externo_ARGY"
    destino.parent.mkdir(exist_ok=True)
    wb.save(destino)
    return destino
