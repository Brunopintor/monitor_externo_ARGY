# Monitor externo de Argentina: tipos de cambio y reservas

Automatización de una base de datos diaria sobre el panorama externo de la Argentina:
tipos de cambio (oficial, financieros y paralelos), brechas, banda cambiaria, dólar global (DXY),
real brasileño y reservas internacionales del BCRA, brutas y netas.

Todos los días hábiles, un proceso automático baja los datos de fuentes públicas (BCRA, ArgentinaDatos,
DolarApi y Yahoo Finance), calcula las variables derivadas (brechas, posición en la banda y reservas netas
con metodología FMI) y actualiza un tablero online:

**📊 Tablero: https://brunopintor.github.io/monitor_externo_ARGY/**

Autor: **Bruno Pintor** · [X / Twitter](https://x.com/PintorBruno) · [LinkedIn](https://www.linkedin.com/in/bruno-pintor-ab7438131/)

## Cómo correrlo

```
pip install pandas requests openpyxl
python actualizar_base.py            # base completa desde 2003 (~45 s)
python actualizar_base.py --desde 2024-01-01
```

Genera `data/base_diaria.csv`, `data/base_diaria.xlsx`, `data/reservas_netas.csv` y `data/ultimo_dato.json`.

Después, para el tablero:

```
python tablero.py                    # arma docs/index.html con los datos de data/
```

`docs/index.html` se abre con doble clic (necesita internet para cargar ECharts y la fuente).

## Archivos

| Archivo | Qué hace |
|---|---|
| `bcra_api.py` | Cliente genérico de las dos APIs del BCRA. Sirve para cualquier informe futuro. |
| `paralelos.py` | MEP, CCL, blue, tarjeta, cripto (ArgentinaDatos histórico + DolarApi en vivo). |
| `reservas_netas.py` | Reservas netas (metodología FMI, ex BOPREAL, ex Tesoro, líquidas, metas). |
| `actualizar_base.py` | Arma la base diaria del informe: series, brechas, posición en banda, variaciones. |
| `tablero.py` | Lee `data/`, prepara los datos y los inserta en `tablero_plantilla.html` → `docs/index.html`. |
| `tablero_plantilla.html` | Diseño del tablero (HTML + ECharts). Los cambios de diseño van acá, nunca en `docs/index.html`. |

---

## Cómo funciona la API del BCRA

Es pública: no pide registro, token ni clave. Son pedidos `GET` a `https://api.bcra.gob.ar` que devuelven JSON.

### 1. Estadísticas monetarias (v4.0)

**El catálogo** (`/estadisticas/v4.0/monetarias`) lista unas 1.600 variables. Cada una tiene un `idVariable`, una descripción, una categoría, una unidad, una periodicidad (D/M) y su último dato.

**La serie** (`/estadisticas/v4.0/monetarias/{id}?desde=AAAA-MM-DD&hasta=AAAA-MM-DD&limit=3000&offset=0`) trae como máximo 3.000 registros por pedido. Si hay más, se pagina con `offset`; `bcra_api.py` lo hace solo.

```python
from bcra_api import BCRA
b = BCRA()
b.buscar("reservas")            # encontrar el id de lo que necesito
b.serie(5, desde="2024-01-01")  # A3500
```

Categorías disponibles: Principales Variables, Series.xlsm (las mismas series del Excel de series diarias del BCRA), Informe Monetario Diario, Activos y pasivos del BCRA, agregados monetarios, depósitos y préstamos por tipo de titular y por línea, y tasas.

IDs que usa este informe:

| id | Serie |
|---|---|
| 5 | Tipo de cambio mayorista de referencia (A3500) |
| 4 | Tipo de cambio minorista (promedio vendedor) |
| 1187 / 1188 | Banda cambiaria, límite inferior / superior (se publica por adelantado) |
| 1 | Reservas internacionales brutas (MM USD) |
| 78–82 | Factores de variación diaria de reservas |
| 1243 | Cuentas corrientes en moneda extranjera en el BCRA (encajes USD) |
| 1266 | Depósitos del gobierno en el BCRA en moneda extranjera (MM ARS) |

### 2. Estadísticas cambiarias (v1.0)

- `/estadisticascambiarias/v1.0/Maestros/Divisas` lista las monedas.
- `/estadisticascambiarias/v1.0/Cotizaciones?fecha=AAAA-MM-DD` trae todas las monedas de un día.
- `/estadisticascambiarias/v1.0/Cotizaciones/{moneda}?fechadesde=&fechahasta=&limit=1000&offset=` trae la serie de una moneda. Aquí `limit` va de 10 a 1000.

Cada registro tiene dos campos:

- `tipoPase`: dólares por unidad. Para el oro (XAU) es el precio en USD de la onza.
- `tipoCotizacion`: pesos por unidad.

Códigos útiles: `USD`, `CNY` (yuan), `XDR` (DEG), `XAU` (oro), `EUR`, `BRL`.

### Detalles a tener en cuenta

- **Rezago:** el A3500 llega el mismo día o al siguiente. Reservas y factores de variación llegan con 1 o 2 días hábiles de rezago.
- **Fechas futuras:** las bandas vienen publicadas hacia adelante. El script corta en la fecha de hoy.
- **Días sin dato:** hay días hábiles sin A3500 en la API (por ejemplo, del 6 al 10 de noviembre de 2026). En el Excel también faltan, así que es un hueco de la fuente.
- **Error 404:** significa que no hay datos en ese rango, no que el endpoint esté caído.
- **Versiones viejas:** la API "Principales variables" v1, v2 y v3 está discontinuada o en proceso de baja. Usar la v4.0.

---

## Dólares paralelos y financieros

No hay una fuente oficial con API, pero no hace falta *scrapear*. Dos APIs comunitarias, gratuitas y sin token, ya lo hacen:

- **ArgentinaDatos** (`api.argentinadatos.com/v1/cotizaciones/dolares/{casa}`) tiene el histórico diario completo. Las casas son: `bolsa` (MEP, desde 2018), `contadoconliqui` (desde 2013), `blue` (desde 2011), `tarjeta`, `cripto` y `mayorista`.
- **DolarApi** (`dolarapi.com/v1/dolares`) da la cotización en vivo, actualizada cada pocos minutos.

**Validación contra el Excel (2025–2026, 450 días):**

| Serie | Diferencia mediana | Percentil 90 |
|---|---|---|
| A3500 | 0,00 % | 0,00 % |
| Reservas brutas | 0,00 % | 0,00 % |
| MEP | 0,2 % | 0,7 % |
| CCL | 0,2 % | 1,0 % |
| Blue | 0,3 % | 1,2 % |

Las diferencias en los paralelos vienen de la fuente y del horario de corte, no de un error. Conviene definir una sola fuente y quedarse con ella.

Plan B si alguna de estas APIs se cae: calcular el MEP y el CCL a partir de los precios de bonos, AL30 / AL30D (MEP) y AL30 / AL30C (CCL).

---

## Reservas netas

`reservas_netas.py` replica la hoja RRII. Combina dos fuentes:

- **De la API:** brutas, encajes, depósitos del Tesoro, yuan, DEG y oro.
- **De `insumos_reservas_netas.xlsx` (carga manual):** pasivos y parámetros.

Para el mapa de fórmulas y la validación ver `CLAUDE.md`.

### Cómo cargar los insumos

Abrir `insumos_reservas_netas.xlsx`, hoja `pasivos`. Cuando cambia algo (vence un REPO, se actualiza el BOPREAL, cambia el stock de DEG), agregar una fila con la fecha y todos los valores vigentes. No hace falta cargar todos los días: el código repite el último valor.

## Lo que la API no cubre (sigue siendo carga manual)

**Reservas netas.** Ningún organismo las publica: hay que calcularlas, como hace hoy la hoja `RRII`. La API sí trae las reservas brutas, los encajes en USD y los depósitos del Tesoro en moneda extranjera.

Siguen siendo carga manual los pasivos: el swap con China, el BOPREAL, los REPOs, las obligaciones con organismos internacionales y el BIS. Como cambian pocas veces, la propuesta es un CSV chico (`pasivos.csv` con fecha, concepto y monto) que se actualiza cuando hay novedades. Con ese archivo, el script calcula las netas solo.

**Hojas analíticas.** Hojas como TCRM, PIES, Esquema IMPO y FMI-pagos quedan fuera de este primer paso.

## Actualización diaria

El tablero se publica en **https://brunopintor.github.io/monitor_externo_ARGY/** y se actualiza solo
de lunes a viernes a las 19:00 (hora argentina) con GitHub Actions (`.github/workflows/actualizar.yml`):
baja los datos, arma el tablero, commitea `data/` y `docs/` y publica en GitHub Pages.

- Para correrlo en el momento: pestaña **Actions** del repo → *Actualizar tablero* → **Run workflow**.
- Para cambiar los insumos manuales: editar `insumos_reservas_netas.xlsx`, commit y push.
- Antes de tocar algo en la PC: `git pull` (el bot sube datos nuevos todos los días).
