"""
Dashboard Consolidado — versión LOCAL
======================================
Uso:
    python dashboard_local.py --anio 2026 --mes marzo

Estructura de carpetas esperada (mismo directorio que el script):
    2026_01_Informe_*.xlsx
    2026_02_Informe_*.xlsx
    ...
    2025_00_Dashboard_Consolidado.xlsx   <- año anterior (opcional, para comparativa)

Salidas generadas en la misma carpeta:
    2026_00_Dashboard_Consolidado.xlsx
    2026_00_Dashboard_Final.html

Dependencias:
    pip install pandas openpyxl xlsxwriter
"""

import argparse
import json
import os
import sys

import pandas as pd
import openpyxl
import xlsxwriter

# =============================================================================
# CONFIGURACIÓN GENERAL
# =============================================================================

MESES_NOMBRE = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]
NOMBRE_A_COD = {n: f"{i+1:02d}" for i, n in enumerate(MESES_NOMBRE)}
COD_A_NOMBRE = {v: k for k, v in NOMBRE_A_COD.items()}

MAPA_ORDEN = [
    "Búsqueda de datos",
    "Bajada de datos",
    "Otras consultas",
    "Cargas en Data Manager",
]

MAPA_COLOR_RGB = {
    "Búsqueda de datos":       (81,  165, 225),
    "Bajada de datos":         (90,  107, 125),
    "Otras consultas":         (35,  169, 133),
    "Cargas en Data Manager":  (231,  84,  55),
}

MAPA_COLOR_HEX = {
    "Búsqueda de datos":       "#51a5e1",
    "Bajada de datos":         "#5a6b7d",
    "Otras consultas":         "#23a985",
    "Cargas en Data Manager":  "#e75437",
}

CAT_PREFIX = {
    "1-": MAPA_ORDEN[0],
    "3-": MAPA_ORDEN[1],
    "5-": MAPA_ORDEN[2],
    "7-": MAPA_ORDEN[3],
}

# =============================================================================
# HELPERS
# =============================================================================

def rgb_to_hex(r, g, b):
    return f"{r:02X}{g:02X}{b:02X}"


def tiempo_natural_sin_findes(row):
    inicio = pd.to_datetime(row['Submit Date'], errors='coerce')
    fin    = pd.to_datetime(row['Completed Date'], errors='coerce')
    if pd.isnull(inicio) or pd.isnull(fin) or fin < inicio:
        return 0.0
    segundos_totales = (fin - inicio).total_seconds()
    segundos_finde   = 0.0
    current = inicio
    while current < fin:
        next_day = (current + pd.Timedelta(days=1)).normalize()
        if next_day > fin:
            next_day = fin
        if current.weekday() >= 5:
            segundos_finde += (next_day - current).total_seconds()
        current = next_day
    return (segundos_totales - segundos_finde) / 3600.0


# =============================================================================
# PASO 1: ARGUMENTOS Y DERIVACIÓN DE MESES
# =============================================================================

def parsear_args():
    parser = argparse.ArgumentParser(description="Genera el Dashboard Consolidado anual en local.")
    parser.add_argument("--anio", required=True, type=int, help="Año a procesar (ej: 2026)")
    parser.add_argument("--mes",  required=True, type=str,
                        help="Mes hasta el que procesar inclusive (ej: marzo)")
    return parser.parse_args()


def construir_meses(anio, mes_nombre):
    """Devuelve lista [{mes, cod, anio}] desde Enero hasta mes_nombre inclusive."""
    mes_cap = mes_nombre.strip().capitalize()
    if mes_cap not in NOMBRE_A_COD:
        print(f"❌ Mes '{mes_nombre}' no reconocido. Valores válidos: {list(NOMBRE_A_COD.keys())}")
        sys.exit(1)
    cod_limite = NOMBRE_A_COD[mes_cap]
    meses = []
    for i, nombre in enumerate(MESES_NOMBRE):
        cod = f"{i+1:02d}"
        meses.append({"mes": nombre, "cod": cod, "anio": str(anio)})
        if cod == cod_limite:
            break
    return meses


# =============================================================================
# PASO 2: LECTURA DE INFORMES XLSX LOCALES
# =============================================================================

def leer_informes(carpeta, meses_a_procesar):
    codigos_ordenados = [m["cod"] for m in meses_a_procesar]
    acum_volumen     = {cod: {cat: 0   for cat in MAPA_ORDEN} for cod in codigos_ordenados}
    acum_horas_total = {cod: {cat: 0.0 for cat in MAPA_ORDEN} for cod in codigos_ordenados}
    acum_horas_count = {cod: {cat: 0   for cat in MAPA_ORDEN} for cod in codigos_ordenados}

    archivos = sorted(f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx'))

    for nombre_archivo in archivos:
        cod_mes = None
        for item in meses_a_procesar:
            if f"{item['anio']}_{item['cod']}" in nombre_archivo:
                cod_mes = item['cod']
                break
        if not cod_mes:
            continue

        ruta = os.path.join(carpeta, nombre_archivo)
        print(f"📂 Procesando: {nombre_archivo} → MES: {cod_mes}")

        try:
            wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        except Exception as e:
            print(f"  ⚠️ No se pudo abrir {nombre_archivo}: {e}")
            continue

        for ws in wb.worksheets:
            prefijo = ws.title[:2]
            if prefijo not in CAT_PREFIX:
                continue
            cat = CAT_PREFIX[prefijo]

            filas = list(ws.iter_rows(values_only=True))
            if len(filas) < 2:
                continue

            cabecera = [str(c).strip() if c is not None else "" for c in filas[0]]
            n_cols = len(cabecera)

            # Normalizar cada fila al mismo nº de columnas que la cabecera.
            # openpyxl puede devolver filas más largas (celdas combinadas) o más
            # cortas (filas incompletas al final), lo que rompe pd.DataFrame.
            filas_norm = []
            for fila in filas[1:]:
                fila = list(fila)
                if len(fila) < n_cols:
                    fila += [None] * (n_cols - len(fila))
                else:
                    fila = fila[:n_cols]
                filas_norm.append(fila)

            df = pd.DataFrame(filas_norm, columns=cabecera)

            if 'Work Order ID' not in df.columns:
                continue
            df = df[df['Work Order ID'].astype(str).str.strip().str.startswith('WO')]
            if 'Status' in df.columns:
                df = df[df['Status'].astype(str).str.strip().str.lower().isin(['cerrado', 'terminado'])]
            df = df.dropna(subset=['Submit Date', 'Completed Date'])

            if len(df) == 0:
                continue

            df['Horas_Resolucion'] = df.apply(tiempo_natural_sin_findes, axis=1)

            n_inv = (df['Horas_Resolucion'] == 0).sum()
            if n_inv > 0:
                print(f"  ⚠️ {ws.title}: {n_inv} ticket(s) con fecha invertida (contabilizan con 0 h).")

            acum_volumen[cod_mes][cat]     += len(df)
            acum_horas_total[cod_mes][cat] += df['Horas_Resolucion'].sum()
            acum_horas_count[cod_mes][cat] += len(df)
            print(f"  ✅ {ws.title}: {len(df)} tickets procesados.")

        wb.close()

    return acum_volumen, acum_horas_total, acum_horas_count


# =============================================================================
# PASO 3: CÁLCULO DE MATRICES
# =============================================================================

def calcular_matrices(codigos_ordenados, acum_volumen, acum_horas_total, acum_horas_count):
    matriz_vol = {}
    matriz_tie = {}
    for cod in codigos_ordenados:
        matriz_vol[cod] = {cat: acum_volumen[cod][cat] for cat in MAPA_ORDEN}
        matriz_tie[cod] = {}
        for cat in MAPA_ORDEN:
            ht = acum_horas_total[cod][cat]
            hc = acum_horas_count[cod][cat]
            matriz_tie[cod][cat] = round(ht / hc, 2) if hc > 0 else 0.0

    df_vol = pd.DataFrame.from_dict(matriz_vol, orient='index').reindex(columns=MAPA_ORDEN)
    df_tie = pd.DataFrame.from_dict(matriz_tie, orient='index').reindex(columns=MAPA_ORDEN)
    return df_vol, df_tie


# =============================================================================
# PASO 4: XLSX LOCAL CON XLSXWRITER
# =============================================================================

def generar_xlsx(ruta_salida, etiqueta_periodo, codigos_ordenados, mapa_meses, df_vol, df_tie):
    wb  = xlsxwriter.Workbook(ruta_salida)
    ws  = wb.add_worksheet("Dashboard")

    def c(r, g, b):
        return f"#{rgb_to_hex(r,g,b)}"

    fmt_titulo  = wb.add_format({'bold': True, 'font_size': 11, 'font_color': '#1A2647'})
    fmt_mes_hdr = wb.add_format({'bold': True, 'font_size': 10, 'font_color': 'white',
                                  'bg_color': '#262729', 'align': 'center', 'border': 1})

    def fmt_cat_hdr(r, g, b):
        return wb.add_format({'bold': True, 'font_size': 10, 'font_color': 'white',
                               'bg_color': c(r,g,b), 'align': 'center', 'border': 1})

    def fmt_data(par, dec=False):
        bg = '#F5F6F8' if par else '#FFFFFF'
        f = {'bg_color': bg, 'border': 1, 'border_color': '#E1E8F0', 'align': 'right'}
        if dec: f['num_format'] = '0.00'
        return wb.add_format(f)

    def fmt_mes_data(par):
        bg = '#F5F6F8' if par else '#FFFFFF'
        return wb.add_format({'bg_color': bg, 'border': 1, 'border_color': '#E1E8F0', 'align': 'left'})

    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 3, 16)
    ws.set_column(4, 4, 24)

    n_vol     = len(codigos_ordenados)
    n_tie     = len(codigos_ordenados)
    start_tie = n_vol + 4   # fila 0-indexed donde empieza tabla tiempos

    # Tabla volumen
    ws.write(0, 0, f"VOLUMEN DE PETICIONES ({etiqueta_periodo})", fmt_titulo)
    ws.write(1, 0, "Mes", fmt_mes_hdr)
    for i, cat in enumerate(MAPA_ORDEN):
        ws.write(1, i+1, cat, fmt_cat_hdr(*MAPA_COLOR_RGB[cat]))

    for ri, cod in enumerate(codigos_ordenados):
        er  = 2 + ri
        par = (ri % 2 == 0)
        ws.write(er, 0, mapa_meses[cod], fmt_mes_data(par))
        for ci, cat in enumerate(MAPA_ORDEN):
            ws.write(er, ci+1, int(df_vol.loc[cod, cat]), fmt_data(par))

    # Tabla tiempos
    ws.write(start_tie,   0, f"TIEMPOS MEDIOS DE RESOLUCIÓN EN HORAS ({etiqueta_periodo})", fmt_titulo)
    ws.write(start_tie+1, 0, "Mes", fmt_mes_hdr)
    for i, cat in enumerate(MAPA_ORDEN):
        ws.write(start_tie+1, i+1, cat, fmt_cat_hdr(*MAPA_COLOR_RGB[cat]))

    for ri, cod in enumerate(codigos_ordenados):
        er  = start_tie + 2 + ri
        par = (ri % 2 == 0)
        ws.write(er, 0, mapa_meses[cod], fmt_mes_data(par))
        for ci, cat in enumerate(MAPA_ORDEN):
            ws.write(er, ci+1, float(df_tie.loc[cod, cat]), fmt_data(par, dec=True))

    # Gráfico volumen
    ch_vol = wb.add_chart({'type': 'column', 'subtype': 'stacked'})
    ch_vol.set_title({'name': f"Volumen de Peticiones por Categoría ({etiqueta_periodo})"})
    ch_vol.set_x_axis({'name': 'Meses'})
    ch_vol.set_y_axis({'name': 'Peticiones'})
    ch_vol.set_legend({'position': 'bottom'})
    for i, cat in enumerate(MAPA_ORDEN):
        r, g, b = MAPA_COLOR_RGB[cat]
        ch_vol.add_series({
            'name':       cat,
            'categories': ['Dashboard', 2, 0, 2+n_vol-1, 0],
            'values':     ['Dashboard', 2, i+1, 2+n_vol-1, i+1],
            'fill':       {'color': c(r,g,b)},
            'data_labels': {'value': True, 'font': {'bold': True, 'color': 'white', 'size': 9}},
        })
    ws.insert_chart('G1', ch_vol, {'x_scale': 1.8, 'y_scale': 1.4})

    # Gráfico tiempos
    ch_tie = wb.add_chart({'type': 'column'})
    ch_tie.set_title({'name': f"Tiempos Medios de Resolución en Horas ({etiqueta_periodo})"})
    ch_tie.set_x_axis({'name': 'Meses'})
    ch_tie.set_y_axis({'name': 'Horas Promedio'})
    ch_tie.set_legend({'position': 'bottom'})
    for i, cat in enumerate(MAPA_ORDEN):
        r, g, b = MAPA_COLOR_RGB[cat]
        ch_tie.add_series({
            'name':       cat,
            'categories': ['Dashboard', start_tie+2, 0, start_tie+2+n_tie-1, 0],
            'values':     ['Dashboard', start_tie+2, i+1, start_tie+2+n_tie-1, i+1],
            'fill':       {'color': c(r,g,b)},
            'data_labels': {'value': True, 'num_format': '0.00',
                            'font': {'bold': True, 'color': '#334155', 'size': 9},
                            'position': 'outside_end'},
        })
    ws.insert_chart('G16', ch_tie, {'x_scale': 1.8, 'y_scale': 1.4})

    wb.close()
    print(f"✅ Excel guardado: {ruta_salida}")


# =============================================================================
# PASO 5: AÑO ANTERIOR (xlsx local)
# =============================================================================

def leer_anio_anterior(carpeta, anio_anterior):
    nombre = f"{anio_anterior}_00_Dashboard_Consolidado.xlsx"
    ruta   = os.path.join(carpeta, nombre)
    vol_ant, tie_ant = {}, {}

    if not os.path.exists(ruta):
        print(f"ℹ️  Archivo del año anterior no encontrado: {nombre}")
        otros = [f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx')]
        print(f"   .xlsx en la carpeta: {otros}")
        return vol_ant, tie_ant

    print(f"📅 Leyendo año anterior: {nombre}")
    try:
        wb  = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        ws  = wb.active
        raw = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()

        tabla_inicio = [r for r, row in enumerate(raw)
                        if row and str(row[0]).strip() == "Mes"]

        def leer_tabla(raw, hrow):
            cats   = [str(c).strip() for c in raw[hrow][1:] if c is not None and str(c).strip()]
            result = {}
            for row in raw[hrow+1:]:
                if not row or not str(row[0]).strip():
                    continue
                mes = str(row[0]).strip()
                cod = NOMBRE_A_COD.get(mes)
                if not cod:
                    break
                result[cod] = {}
                for i, cat in enumerate(cats):
                    try:
                        v = row[i+1]
                        result[cod][cat] = float(str(v).replace(',','.')) if v not in (None,'') else 0.0
                    except (ValueError, IndexError):
                        result[cod][cat] = 0.0
            return result

        if len(tabla_inicio) >= 1: vol_ant = leer_tabla(raw, tabla_inicio[0])
        if len(tabla_inicio) >= 2: tie_ant = leer_tabla(raw, tabla_inicio[1])

        meses_cargados = [COD_A_NOMBRE.get(m, m) for m in sorted(vol_ant.keys())]
        print(f"   Meses cargados: {meses_cargados}")

    except Exception as e:
        print(f"⚠️ Error leyendo año anterior: {e}")

    return vol_ant, tie_ant


# =============================================================================
# PASO 6: HTML
# =============================================================================

def generar_html(ruta_salida, etiqueta_periodo, codigos_ordenados, mapa_meses,
                 df_vol, df_tie, vol_anterior, tie_anterior, anio_anterior):

    meses_codigos = list(codigos_ordenados)
    meses_nombres = [mapa_meses[c] for c in meses_codigos]

    js_colors    = {}
    lista_colores = []
    for cat in MAPA_ORDEN:
        r, g, b  = MAPA_COLOR_RGB[cat]
        rgba = f"rgba({r},{g},{b},1)"
        js_colors[cat] = rgba
        lista_colores.append(rgba)

    datos_vol = {cat: [int(df_vol.loc[c, cat]) for c in codigos_ordenados] for cat in MAPA_ORDEN}
    datos_tie = {cat: [float(df_tie.loc[c, cat]) for c in codigos_ordenados] for cat in MAPA_ORDEN}

    def diff_cell_vol(actual, ant):
        if ant is None: return '<td class="diff-na">—</td>'
        d = int(round(actual)) - int(round(ant))
        if d > 0:   return f'<td class="diff-pos">▲ +{d}</td>'
        elif d < 0: return f'<td class="diff-neg">▼ {d}</td>'
        else:       return '<td class="diff-neu">= 0</td>'

    def diff_cell_tie(actual, ant):
        if ant is None: return '<td class="diff-na">—</td>'
        d = actual - ant
        if d < 0:   return f'<td class="diff-pos">▼ {d:+.2f} h</td>'
        elif d > 0: return f'<td class="diff-neg">▲ {d:+.2f} h</td>'
        else:       return '<td class="diff-neu">= 0</td>'

    hay_av = bool(vol_anterior)
    hay_at = bool(tie_anterior)

    def tabla_vol():
        h = '<table><thead><tr><th style="background-color:black;color:white;">Mes</th>'
        for cat in MAPA_ORDEN:
            hx = MAPA_COLOR_HEX[cat]
            h += f'<th style="background-color:{hx};color:white;">{cat}</th>'
            if hay_av:
                h += f'<th style="background-color:{hx};color:white;opacity:0.7;">vs {anio_anterior}</th>'
        h += "</tr></thead><tbody>"
        for cod in codigos_ordenados:
            h += f"<tr><td><strong>{mapa_meses[cod]}</strong></td>"
            for cat in MAPA_ORDEN:
                actual = int(df_vol.loc[cod, cat])
                h += f"<td>{actual}</td>"
                if hay_av:
                    h += diff_cell_vol(actual, vol_anterior.get(cod, {}).get(cat))
            h += "</tr>"
        return h + "</tbody></table>"

    def tabla_tie():
        h = '<table><thead><tr><th style="background-color:black;color:white;">Mes</th>'
        for cat in MAPA_ORDEN:
            hx = MAPA_COLOR_HEX[cat]
            h += f'<th style="background-color:{hx};color:white;">{cat}</th>'
            if hay_at:
                h += f'<th style="background-color:{hx};color:white;opacity:0.7;">vs {anio_anterior}</th>'
        h += "</tr></thead><tbody>"
        for cod in codigos_ordenados:
            h += f"<tr><td><strong>{mapa_meses[cod]}</strong></td>"
            for cat in MAPA_ORDEN:
                actual = float(df_tie.loc[cod, cat])
                h += f"<td>{actual:.2f}</td>"
                if hay_at:
                    h += diff_cell_tie(actual, tie_anterior.get(cod, {}).get(cat))
            h += "</tr>"
        return h + "</tbody></table>"

    opts = "".join(f'<option value="{c}">{n}</option>'
                   for c, n in zip(meses_codigos, meses_nombres))

    html_template = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>BI Control Panel Operational __PERIODO__</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
  <style>
    :root{--bg-main:#f4f6f9;--bg-card:#ffffff;--text-primary:#0f172a;--text-secondary:#64748b;--border-color:#e2e8f0;--accent-blue:#2563eb;}
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg-main);color:var(--text-primary);padding:16px;}
    .container{max-width:1400px;margin:0 auto;display:flex;flex-direction:column;gap:20px;}
    .header-panel{background:var(--bg-card);padding:16px 20px;border-radius:12px;border:1px solid var(--border-color);box-shadow:0 4px 8px rgba(0,0,0,.05);display:flex;flex-direction:column;gap:14px;}
    @media(min-width:640px){.header-panel{flex-direction:row;justify-content:space-between;align-items:center;}}
    .header-panel h1{font-size:1.4rem;font-weight:700;}
    .header-panel p{font-size:.85rem;color:var(--text-secondary);margin-top:2px;}
    .filter-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
    .filter-row label{font-size:.85rem;color:var(--text-secondary);white-space:nowrap;}
    .filter-row select{padding:6px 10px;border-radius:8px;border:1px solid var(--border-color);font-size:.85rem;background:white;cursor:pointer;}
    .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;}
    .kpi-card{background:var(--bg-card);border-radius:12px;padding:14px 18px;border:1px solid var(--border-color);border-left:5px solid var(--accent-blue);box-shadow:0 2px 6px rgba(0,0,0,.04);}
    .kpi-label{font-size:.78rem;color:var(--text-secondary);font-weight:600;text-transform:uppercase;letter-spacing:.05em;}
    .kpi-value{font-size:1.7rem;font-weight:700;margin-top:4px;}
    .kpi-desc{font-size:.78rem;color:var(--text-secondary);margin-top:4px;}
    .charts-row{display:grid;grid-template-columns:2fr 1fr;gap:16px;}
    @media(max-width:900px){.charts-row{grid-template-columns:1fr;}}
    .card{background:var(--bg-card);border-radius:12px;padding:16px 20px;border:1px solid var(--border-color);box-shadow:0 2px 6px rgba(0,0,0,.04);}
    .card-title{font-size:.95rem;font-weight:700;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;}
    .chart-wrap{position:relative;height:280px;}
    .btn-group{display:flex;gap:6px;}
    .btn-group button{padding:4px 12px;border-radius:8px;border:1px solid var(--border-color);background:white;font-size:.8rem;cursor:pointer;transition:background .15s;}
    .btn-group button.active{background:var(--accent-blue);color:white;border-color:var(--accent-blue);}
    table{border-collapse:collapse;width:100%;font-size:.82rem;}
    thead th{padding:7px 10px;text-align:left;position:sticky;top:0;z-index:1;}
    tbody td{padding:6px 10px;border-bottom:1px solid var(--border-color);}
    tbody tr:hover{background:#f8fafc;}
    table td:not(:first-child){text-align:right;}
    td.diff-pos{color:#16a34a;font-weight:700;text-align:right;}
    td.diff-neg{color:#dc2626;font-weight:700;text-align:right;}
    td.diff-neu{color:#94a3b8;font-weight:600;text-align:right;}
    td.diff-na{color:#cbd5e1;text-align:right;}
    .table-wrap{overflow-x:auto;max-height:340px;overflow-y:auto;}
  </style>
</head>
<body>
<div class="container">
  <div class="header-panel">
    <div><h1>BI Control Panel Operational</h1><p>Periodo: __PERIODO__</p></div>
    <div class="filter-row">
      <label>Desde:</label><select id="selDesde" onchange="render()">__OPTS__</select>
      <label>Hasta:</label><select id="selHasta" onchange="render()">__OPTS__</select>
    </div>
  </div>
  <div class="kpi-grid">
    <div class="kpi-card" id="kpiTotal">
      <div class="kpi-label">Total Periodo</div>
      <div class="kpi-value" id="valTotal">—</div>
      <div class="kpi-desc">peticiones</div>
    </div>
    <div class="kpi-card" id="kpiEficiencia">
      <div class="kpi-label">Mejor Tiempo Medio</div>
      <div class="kpi-value" id="valEficiencia">—</div>
      <div class="kpi-desc" id="descEficiencia">—</div>
    </div>
    <div class="kpi-card" id="kpiLatencia">
      <div class="kpi-label">Peor Tiempo Medio</div>
      <div class="kpi-value" id="valLatencia">—</div>
      <div class="kpi-desc" id="descLatencia">—</div>
    </div>
  </div>
  <div class="charts-row">
    <div class="card">
      <div class="card-title">Evolución de Cargas de Trabajo
        <div class="btn-group">
          <button id="btnApilado" class="active" onclick="setModo(true)">Apilado</button>
          <button id="btnParalelo" onclick="setModo(false)">En paralelo</button>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="canvasVolumen"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Mix Global de Peticiones (Periodo)</div>
      <div class="chart-wrap"><canvas id="canvasAnillo"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title">Tiempos Medios de Resolución en Horas</div>
    <div class="chart-wrap"><canvas id="canvasTiempos"></canvas></div>
  </div>
  <div class="card">
    <div class="card-title">Detalle — Volumen de Peticiones</div>
    <div class="table-wrap" id="tablaVolumenContenedor">__TABLA_VOL__</div>
  </div>
  <div class="card">
    <div class="card-title">Detalle — Tiempos Medios (horas)</div>
    <div class="table-wrap" id="tablaTiemposContenedor">__TABLA_TIE__</div>
  </div>
</div>
<script>
  if(typeof ChartDataLabels!=='undefined') Chart.register(ChartDataLabels);

  var codigosMeses   = __CODIGOS_MESES__;
  var nombresMeses   = __NOMBRES_MESES__;
  var categorias     = __CATEGORIAS__;
  var colores        = __COLORES__;
  var listaColores   = __LISTA_COLORES__;
  var rawVol         = __RAW_VOL__;
  var rawTie         = __RAW_TIE__;
  var rawVolAnt      = __VOL_ANT__;
  var rawTieAnt      = __TIE_ANT__;
  var anioAnt        = "__ANIO_ANT__";

  var chVol=null, chAnillo=null, chTie=null, apilado=true;
  var selD=document.getElementById('selDesde'), selH=document.getElementById('selHasta');
  selH.selectedIndex = selH.options.length-1;

  function setModo(v){
    apilado=v;
    document.getElementById('btnApilado').classList.toggle('active',v);
    document.getElementById('btnParalelo').classList.toggle('active',!v);
    render();
  }

  function render(){
    var iD=selD.selectedIndex, iH=selH.selectedIndex;
    if(iH<iD){iH=iD; selH.selectedIndex=iH;}
    var meses=nombresMeses.slice(iD,iH+1);
    var dsVol=[], dsTie=[], colAnillo=[], totAnillo=[], catsAnillo=[];

    categorias.forEach(function(cat,i){
      var vd=rawVol[cat].slice(iD,iH+1), td=rawTie[cat].slice(iD,iH+1);
      var tot=vd.reduce(function(a,b){return a+b;},0);
      dsVol.push({label:cat,data:vd,backgroundColor:colores[cat]});
      dsTie.push({label:cat,data:td,backgroundColor:colores[cat],borderColor:colores[cat],fill:false});
      if(tot>0){totAnillo.push(tot);catsAnillo.push(cat);colAnillo.push(listaColores[i]);}
    });

    var total=totAnillo.reduce(function(a,b){return a+b;},0);
    document.getElementById('valTotal').innerText=total.toLocaleString();

    var mn=Infinity,mx=-1,mnCat='',mxCat='',mnMes='',mxMes='';
    categorias.forEach(function(cat){
      rawTie[cat].slice(iD,iH+1).forEach(function(v,idx){
        if(v>0&&v<mn){mn=v;mnCat=cat;mnMes=meses[idx];}
        if(v>mx)     {mx=v;mxCat=cat;mxMes=meses[idx];}
      });
    });
    document.getElementById('valEficiencia').innerText=(mn===Infinity?'0.00':mn.toFixed(2))+' Horas';
    document.getElementById('descEficiencia').innerText=mnCat?mnCat+' durante '+mnMes:'Sin datos';
    if(mnCat) document.getElementById('kpiEficiencia').style.borderLeftColor=colores[mnCat];
    document.getElementById('valLatencia').innerText=(mx===-1?'0.00':mx.toFixed(2))+' Horas';
    document.getElementById('descLatencia').innerText=mxCat?mxCat+' durante '+mxMes:'Sin datos';
    if(mxCat) document.getElementById('kpiLatencia').style.borderLeftColor=colores[mxCat];

    if(chVol) chVol.destroy();
    chVol=new Chart(document.getElementById('canvasVolumen'),{
      type:'bar',data:{labels:meses,datasets:dsVol},
      options:{responsive:true,maintainAspectRatio:false,
        scales:{x:{stacked:apilado},y:{stacked:apilado,ticks:{precision:0}}},
        plugins:{legend:{position:'bottom'},datalabels:{
          display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},
          formatter:function(v){return v;},font:{weight:'bold',size:11},
          anchor:function(ctx){
            var mi=ctx.dataIndex,tot=dsVol.reduce(function(s,d){return s+(d.data[mi]||0);},0);
            return(apilado&&ctx.dataset.data[mi]/tot<0.03)?'end':'center';
          },
          align:function(ctx){
            var mi=ctx.dataIndex,tot=dsVol.reduce(function(s,d){return s+(d.data[mi]||0);},0);
            if(!apilado||ctx.dataset.data[mi]/tot>=0.03) return 'center';
            return(ctx.datasetIndex%2===0)?'right':'left';
          },
          color:'#ffffff',
          textStrokeColor:function(ctx){return ctx.dataset.backgroundColor;},
          textStrokeWidth:3,
          offset:function(ctx){
            var mi=ctx.dataIndex,tot=dsVol.reduce(function(s,d){return s+(d.data[mi]||0);},0);
            return(apilado&&ctx.dataset.data[mi]/tot<0.03)?6:0;
          }
        }}
      }
    });

    if(chAnillo) chAnillo.destroy();
    chAnillo=new Chart(document.getElementById('canvasAnillo'),{
      type:'doughnut',
      data:{labels:catsAnillo,datasets:[{data:totAnillo,backgroundColor:colAnillo}]},
      options:{responsive:true,maintainAspectRatio:false,cutout:'70%',
        plugins:{legend:{position:'bottom'},datalabels:{
          display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},
          color:'#ffffff',
          textStrokeColor:function(ctx){return colAnillo[ctx.dataIndex];},
          textStrokeWidth:3,font:{weight:'bold',size:12},
          formatter:function(v){return v;}
        }}
      }
    });

    if(chTie) chTie.destroy();
    chTie=new Chart(document.getElementById('canvasTiempos'),{
      type:'bar',data:{labels:meses,datasets:dsTie},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{position:'bottom'},datalabels:{
          display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},
          anchor:'end',align:'end',color:'#334155',font:{weight:'bold',size:11},
          formatter:function(v){return v.toFixed(2);}
        }}
      }
    });

    tablaHtml('tablaVolumenContenedor',iD,meses,dsVol,false,rawVolAnt);
    tablaHtml('tablaTiemposContenedor',iD,meses,dsTie,true,rawTieAnt);
  }

  function tablaHtml(id,iD,meses,ds,dec,ant){
    var h='<table><thead><tr><th style="background-color:black;color:white;">Mes</th>';
    ds.forEach(function(d){
      h+='<th style="background-color:'+colores[d.label]+';color:white;">'+d.label+'</th>';
      if(ant&&Object.keys(ant).length>0)
        h+='<th style="background-color:'+colores[d.label]+';color:white;opacity:0.7;">vs '+anioAnt+'</th>';
    });
    h+='</tr></thead><tbody>';
    meses.forEach(function(mes,idx){
      var cod=codigosMeses[iD+idx];
      h+='<tr><td><strong>'+mes+'</strong></td>';
      ds.forEach(function(d){
        var v=d.data[idx];
        h+='<td>'+(dec?v.toFixed(2):v.toLocaleString())+'</td>';
        if(ant&&Object.keys(ant).length>0){
          var va=(ant[cod]&&ant[cod][d.label]!==undefined)?ant[cod][d.label]:null;
          if(va===null){h+='<td class="diff-na">—</td>';}
          else{
            var df=v-va;
            if(dec){
              if(df<0) h+='<td class="diff-pos">▼ '+df.toFixed(2)+' h</td>';
              else if(df>0) h+='<td class="diff-neg">▲ +'+df.toFixed(2)+' h</td>';
              else h+='<td class="diff-neu">= 0</td>';
            }else{
              var di=Math.round(df);
              if(di>0) h+='<td class="diff-pos">▲ +'+di+'</td>';
              else if(di<0) h+='<td class="diff-neg">▼ '+di+'</td>';
              else h+='<td class="diff-neu">= 0</td>';
            }
          }
        }
      });
      h+='</tr>';
    });
    h+='</tbody></table>';
    document.getElementById(id).innerHTML=h;
  }

  render();
</script>
</body>
</html>"""

    html = (html_template
        .replace("__PERIODO__",      etiqueta_periodo)
        .replace("__OPTS__",         opts)
        .replace("__TABLA_VOL__",    tabla_vol())
        .replace("__TABLA_TIE__",    tabla_tie())
        .replace("__CODIGOS_MESES__", json.dumps(meses_codigos))
        .replace("__NOMBRES_MESES__", json.dumps(meses_nombres))
        .replace("__CATEGORIAS__",   json.dumps(MAPA_ORDEN))
        .replace("__COLORES__",      json.dumps(js_colors))
        .replace("__LISTA_COLORES__", json.dumps(lista_colores))
        .replace("__RAW_VOL__",      json.dumps(datos_vol))
        .replace("__RAW_TIE__",      json.dumps(datos_tie))
        .replace("__VOL_ANT__",      json.dumps({k:{c:int(v) for c,v in d.items()} for k,d in vol_anterior.items()}))
        .replace("__TIE_ANT__",      json.dumps({k:{c:round(float(v),2) for c,v in d.items()} for k,d in tie_anterior.items()}))
        .replace("__ANIO_ANT__",     str(anio_anterior))
    )

    with open(ruta_salida, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ HTML guardado: {ruta_salida}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    args = parsear_args()

    meses_a_procesar  = construir_meses(args.anio, args.mes)
    codigos_ordenados = [m["cod"] for m in meses_a_procesar]
    mapa_meses        = {m["cod"]: m["mes"] for m in meses_a_procesar}
    etiqueta_periodo  = str(args.anio)
    anio_anterior     = str(args.anio - 1)

    # sys.frozen es True cuando se ejecuta como .exe compilado con PyInstaller.
    # En ese caso __file__ apunta a la carpeta temporal, no al .exe real.
    if getattr(sys, "frozen", False):
        carpeta = os.path.dirname(os.path.abspath(sys.executable))
    else:
        carpeta = os.path.dirname(os.path.abspath(__file__))

    print(f"\n🚀 Iniciando — {etiqueta_periodo}, hasta {mapa_meses[codigos_ordenados[-1]]}")
    print(f"   Meses: {[mapa_meses[c] for c in codigos_ordenados]}")
    print(f"   Carpeta: {carpeta}\n")

    acum_vol, acum_ht, acum_hc = leer_informes(carpeta, meses_a_procesar)
    df_vol, df_tie = calcular_matrices(codigos_ordenados, acum_vol, acum_ht, acum_hc)

    nombre_xlsx = f"{args.anio}_00_Dashboard_Consolidado.xlsx"
    generar_xlsx(os.path.join(carpeta, nombre_xlsx), etiqueta_periodo,
                 codigos_ordenados, mapa_meses, df_vol, df_tie)

    vol_ant, tie_ant = leer_anio_anterior(carpeta, anio_anterior)

    nombre_html = f"{args.anio}_00_Dashboard_Final.html"
    generar_html(os.path.join(carpeta, nombre_html), etiqueta_periodo,
                 codigos_ordenados, mapa_meses, df_vol, df_tie,
                 vol_ant, tie_ant, anio_anterior)

    print(f"\n🎉 Listo. Archivos generados:")
    print(f"   📊 {nombre_xlsx}")
    print(f"   🌐 {nombre_html}")


if __name__ == "__main__":
    main()
