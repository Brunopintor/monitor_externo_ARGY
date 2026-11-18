# CLAUDE.md: Automatización Reservas y Tipos de Cambio

Contexto para Claude Code. Leer entero antes de tocar código.

## Qué es este proyecto

Bruno Pintor es asesor económico en el Ministerio de Economía de la Provincia de Buenos Aires. Su base diaria de tipo de cambio, banda cambiaria y reservas internacionales se mantenía a mano en el Excel `../(2024.03.06) Reservas y tipos de cambio - nuevo.xlsx` (hojas BASE y RRII).

Este proyecto la reemplaza con Python y la API del BCRA. Más adelante va a alimentar un tablero HTML que se actualice solo todos los días.

- Idioma: español rioplatense, en el código, los comentarios y las respuestas.
- Fuente preferida para entregables visuales (Excel, HTML, gráficos): **Aptos**.

## Estado actual (lo que ya está hecho y validado)

| Archivo | Qué hace |
|---|---|
| `bcra_api.py` | Cliente genérico de las APIs del BCRA (monetarias v4.0 y cambiarias v1.0), con paginación y reintentos. **Es reutilizable para otros informes; no meterle lógica específica de este informe.** |
| `paralelos.py` | MEP, CCL, blue, tarjeta y cripto. Histórico desde ArgentinaDatos; en vivo desde DolarApi. |
| `reservas_netas.py` | Replica la hoja RRII: netas metodología FMI, ex BOPREAL, ex depósitos del Tesoro, sin EFF a precios TMU, líquidas, metas RIN y controles. |
| `internacionales.py` | DXY desde Yahoo Finance (DX-Y.NYB; no oficial: si falla, avisa y deja la serie vacía). El real sale del BCRA (`BRL` en `MONEDAS`). |
| `actualizar_base.py` | Script principal. Arma `data/base_diaria.csv/.xlsx` (equivale a la hoja BASE + netas), `data/reservas_netas.csv` y `data/ultimo_dato.json`. |
| `insumos_reservas_netas.xlsx` | Lo carga Bruno a mano: pasivos que no están en la API (BOPREAL, REPOs+SEDESA, organismos, swap USA, BIS, stock DEG) y parámetros. |
| `tablero.py` | Lee `data/` (no llama APIs), arma el JSON del tablero y lo inserta en la plantilla → `docs/index.html`. |
| `tablero_plantilla.html` | HTML + ECharts 5.5.1 (CDN). Tiene la marca `/*__DATOS__*/null` que reemplaza `tablero.py`. Los cambios de diseño van acá. |

Para correrlo: `pip install pandas requests openpyxl` y después `python actualizar_base.py` (tarda ~1 minuto) y `python tablero.py`.

Para previsualizar el tablero en Claude Code: servidor `tablero` en `.claude/launch.json` (http.server sobre `docs/`, puerto 8765).

### Validación hecha

- **A3500 y reservas brutas:** coinciden al 100 % con el Excel original.
- **MEP, CCL y blue:** diferencia mediana de 0,2–0,3 % con el Excel. Viene de la fuente y del horario de corte.
- **Banda cambiaria:** se usa la oficial de la API (ids 1187/1188). El Excel tenía un cálculo propio que difería en 0,35 %.
- **Reservas netas FMI, ex BOPREAL, ex Tesoro, metas y RIN sin FMI:** coinciden exacto con el Excel en 2025–2026 (más del 99 % de los días). En 2024, la mediana de la diferencia es 0 y el 90 % de los días están dentro de ±27 MM.
- **Líquidas:** difiere hasta ~1.700 MM en 2026. No es un error del código: el Excel tiene el oro congelado, y DEG y SEDESA con el error de autollenado (ver abajo).
- **"Netas a precios TMU":** en el Excel original está rota (da ~166.000 MM). El código la calcula bien.

## Mapa de fórmulas (hoja RRII → código)

Todo en millones de USD. Las letras son las columnas de la hoja RRII.

```
K exceso_deg       = max(stock_deg − 3166 × USD/DEG, 0)
P swap_china_usd   = 130.000 MM CNY × ($/CNY) / A3500
D netas_fmi        = brutas − K − P − bis − encajes − swap_usa − organismos − repos(+SEDESA)
E netas_ex_bopreal = D − bopreal
F netas_ex_bopreal_tesoro = E − dep. Tesoro USD
AH rin_sin_fmi_2025 = E − 12.151 (desde 15/04/25) − 2.028 (desde 04/08/25)
G netas_sin_eff_tmu = AH + ajustes de valuación de oro, DEG y yuan a precios del 31/01/25
H liquidas         = brutas − stock_deg + 4.500 − oro − P − sedesa
L var_vs_meta_rin  = E − (−5.058,32)   [E al 30/12/2024]
AI control_netas   = D + encajes − encajes(t−3)
AJ control_super   = AI − dep. Tesoro − bopreal
```

### Series de la API

- **Brutas:** id 1. Rezago de 1–2 días hábiles.
- **Encajes USD:** id 1243. Coincide exacto con la columna X del Excel.
- **Depósitos del Tesoro en moneda extranjera:** id 1266 (en MM ARS) / A3500. Coincide exacto con la columna Y.
- **A3500:** id 5.
- **Bandas:** ids 1187 y 1188. Vienen publicadas a futuro; el código corta en la fecha de hoy.
- **Factores de variación de reservas:** ids 78–82.
- **Cambiarias:** `CNY` (tipoCotizacion = $/CNY), `XDR` (tipoPase = USD/DEG), `XAU` (tipoPase = USD por onza).

## Decisiones tomadas (no revertir sin consultar)

1. **Fuentes de paralelos:** ArgentinaDatos para el histórico y DolarApi en vivo. No se hace scraping propio. Plan B si se caen: calcular MEP y CCL con AL30 / AL30D / AL30C.
2. **Insumos manuales:** van en `insumos_reservas_netas.xlsx`, hoja `pasivos`. Cada fila es una "fecha desde" con todos los valores vigentes, y el código repite el último valor hacia adelante. Bruno agrega una fila cuando algo cambia.
3. **REPOs incluye SEDESA:** es el criterio del Excel desde enero de 2026, y se aplicó hacia atrás sumando SEDESA a repos antes de 2026. Por eso la serie es continua.
4. **Oro:** se calcula como 1,984 MM onzas × precio XAU, porque el Excel lo tenía congelado desde el 23/12/2025. Se puede volver a carga manual con `oro_metodo = manual` en la hoja parametros.
5. **Errores del Excel original corregidos en los insumos:**
   - Stock DEG fijo en 492,97 desde el 19/08/2026 (antes subía +1 por día).
   - SEDESA fijo en 3.287,70 desde el 29/12/2025.
   - **Pendiente:** Bruno tiene que verificar ambos valores contra el balance del BCRA.
6. **Histórico:** las netas se calculan desde 2024. Antes de esa fecha, si hace falta, tomar los valores congelados de la hoja RRII original. No intentar replicar fórmulas viejas.

## Repositorio y publicación (desde el 18/11/2026)

- **Repo:** https://github.com/Brunopintor/monitor_externo_ARGY (público). Rama `main`.
- **Tablero publicado:** https://brunopintor.github.io/monitor_externo_ARGY/ (GitHub Pages sobre `main` / `docs`).
- **Carpeta de trabajo:** `C:\Users\bruno\proyectos\monitor_externo_ARGY`. Se sacó de Google Drive porque un repo de git adentro de una carpeta sincronizada da problemas. La carpeta `Automatizacion` de Drive quedó como copia vieja; no se trabaja más ahí.
- **Git:** autor "Bruno Pintor" con el email privado de GitHub (`91093548+Brunopintor@users.noreply.github.com`), configurado sólo en este repo. `gh` está instalado y logueado (cuenta Brunopintor).
- **`.gitignore`:** deja afuera `data/base_diaria.xlsx` (binario, se regenera) y `desktop.ini`.
- **Para publicar a mano:** `python actualizar_base.py`, `python tablero.py`, y después commit y push de `data/` y `docs/`.

## Actualización diaria (andando desde el 18/11/2026)

- **Workflow:** `.github/workflows/actualizar.yml`. Corre de lunes a viernes a las 22:00 UTC (19:00 ART) y también a mano (Actions → "Actualizar tablero" → Run workflow, o `gh workflow run actualizar.yml`).
- **Pasos:** instala `requirements.txt`, corre `actualizar_base.py` y `tablero.py`, commitea `data/` y `docs/` como github-actions[bot] ("Datos al AAAA-MM-DD") y publica `docs/` con `actions/deploy-pages`.
- **Pages se publica desde Actions** (`build_type = workflow`), no desde la rama: así se publica en la misma corrida.
- **Probado el 18/11/2026:** tardó ~4,5 min (la base, ~3,5 min). Desde GitHub la API del BCRA anduvo sin problemas de certificados y Yahoo respondió.
- **Antes de trabajar local:** hacer `git pull`, porque el bot commitea `data/` y `docs/` todos los días.
- **`insumos_reservas_netas.xlsx`:** Bruno lo edita en esta carpeta, hace commit y push, y entra en la próxima corrida; para verlo antes, se corre el workflow a mano.
- **Dependencias:** `requirements.txt` con rangos acotados (`pandas>=3,<4`, etc.) para que una versión mayor nueva no rompa la corrida.
- **Si falla una corrida:** GitHub le manda un mail a Bruno. Lo más probable es una caída de la API del BCRA o de ArgentinaDatos; se reintenta a mano. Si el DXY falla, la corrida no se corta: avisa y deja la serie vacía.
- **Fuentes vacías:** si una serie viene vacía (Yahoo o una moneda del BCRA), la corrida sigue, imprime `AVISO: vinieron vacías estas series: ...` y el KPI no aparece. Sólo corta con error si faltan `a3500` o `reservas_brutas`. El 18/11/2026 una corrida falló por esto, antes del arreglo; la siguiente anduvo.
- **Avisos pendientes de GitHub** (no frenan nada): las actions v4/v5 usan Node 20, que está deprecado, y `ubuntu-latest` pasa a Ubuntu 26. Cuando salgan versiones nuevas de las actions, subirlas.

## Tablero (hecho el 18/11/2026)

- **Autoría:** "Por Bruno Pintor" arriba a la derecha y, en el pie, X (https://x.com/PintorBruno), LinkedIn (https://www.linkedin.com/in/bruno-pintor-ab7438131/) y el repo.
- **Fuentes y notas:** el pie dice sólo "Fuente: BCRA, ArgentinaDatos, DolarApi y Yahoo Finance". Las aclaraciones van debajo de cada gráfico (brecha, metodología de netas, conceptos en cero en la composición). No se aclara cómo se calculan las variaciones del día y del mes: Bruno lo considera obvio.

Decisiones de Bruno, no revertir sin consultar:

- **Público:** el equipo del Ministerio, pero puede ser público en GitHub Pages (todo es información pública). Además va a ir a su portafolio/CV, así que tiene que verse prolijo.
- **Librería:** ECharts (Bruno no va a tocar los gráficos a mano).
- **Fuente:** Aptos si está instalada; si no, Inter (Google Fonts), que es la más parecida.
- **Pestañas:** Resumen (KPIs agrupados: dólar en la Argentina, dólar en el mundo —DXY y real— y reservas, con variación del día, del mes y minicurva de 3 meses), Dólares y Reservas. Cada una tiene su ancla (`#resumen`, `#dolares`, `#reservas`) y sus gráficos se dibujan al abrirla.
- **Gráficos:**
  - Tipos de cambio (A3500, MEP, CCL, blue; minorista, tarjeta y cripto se prenden desde la leyenda) con la banda sombreada. **Por defecto, últimos 4 años.** Al costado, una tablita con la variación del día y del mes de cada dólar.
  - Brechas MEP, CCL y blue contra el A3500 (4 años por defecto).
  - Reservas brutas y netas: **el histórico más largo posible** (brutas desde 2003, netas desde 2024).
  - Waterfall mensual: cómo se movieron las brutas en el último mes, por factor (ids 78–82). Lo que todavía no tiene factores publicados va a "Sin desagregar".
  - Waterfall de brutas a netas en el último día con netas. Omite los conceptos que dan 0 y lo aclara en el pie.
- **Convenciones:**
  - "Variación del mes" = contra el último dato en o antes de (fecha − 1 mes). El waterfall mensual usa la misma ventana.
  - Formato argentino (1.552,26). Cada KPI muestra la fecha de su propio último dato, porque las series tienen rezagos distintos.
  - Cada dólar tiene un color fijo en todos los gráficos (orden de `DOLARES` en `tablero.py` = `COLOR` en la plantilla).
  - Zoom con Ctrl + ruedita o con la barra de abajo; la ruedita sola baja por la página.
- **Se descartó:** las barras apiladas de factores diarios, porque las reemplaza el waterfall mensual.

## Próximos pasos (en este orden)

1. **Tests mínimos** (pytest), que se podrían correr también en el workflow antes de publicar:
   - Que el catálogo de la API siga teniendo los ids usados.
   - Que `reservas_netas` reproduzca un día conocido. Por ejemplo, 2026-11-05: brutas 49.586 y netas_fmi 8.453,31.
2. Opcional, más adelante: portar otras hojas del Excel (TCRM, PIES, Efectivo Mínimo).

## Cuidados

- **La API del BCRA:**
  - A veces devuelve 404 cuando no hay datos en el rango; eso no es un error.
  - Monetarias acepta `limit` hasta 3000; cambiarias, de 10 a 1000.
  - Hay días hábiles sin A3500 (por ejemplo, del 6 al 10 de noviembre de 2026); también faltan en el Excel.
- **Certificados en Windows:** en algunas PCs falla la cadena SSL del BCRA. En ese caso usar `BCRA(verify=certifi.where())`; nunca usar `verify=False`.
- **Archivos originales:** no modificar los Excel originales de la carpeta padre; son sólo referencia.
- **Credenciales:** no subir credenciales. Hoy no hay ninguna (todas las APIs son públicas).
