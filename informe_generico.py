"""
Informe HTML con gráficos interactivos (Plotly) para cualquier archivo de texto delimitado.
Corre primero el análisis de analisis_generico.py (deben estar en la misma carpeta) y arma el informe
con lo que ese análisis detectó: no hace falta correr los dos por separado.

Uso:   python informe_generico.py archivo.txt [--monto COLUMNA] [--sep "|"]
Requiere: pip install pandas openpyxl plotly
Salida: salida_analisis/informe.html (se abre con cualquier navegador, sin internet)
"""
import html as _html
import os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from plotly.offline import get_plotlyjs

import analisis_generico as A          # ejecuta el análisis completo y deja sus resultados disponibles

df, crudo, N, MONTO, rol = A.df, A.crudo, A.N, A.MONTO, A.rol
SALIDA = A.ruta_salida("informe.html")
MAX_PUNTOS = 20000        # máximo de puntos en los gráficos de dispersión
MAX_ATIPICOS = 300        # máximo de atípicos dibujados por caja
MAX_CAMPOS = 6            # máximo de categorías (y de fechas) que se grafican
UMBRAL_ALTO = 1.0         # % de registros desde el cual un control se marca en rojo (criterio orientativo)

AZUL, NARANJA, VERDE, GRIS = "#2a78d6", "#eb6834", "#1baf7a", "#898781"
SERIES = [AZUL, NARANJA, VERDE, "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]

esc = lambda x: _html.escape(str(x))
fmt = lambda x: f"{x:,.0f}".replace(",", ".")
fmt2 = lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
pct_txt = lambda p: f"{p:.2f}".replace(".", ",") + "%"
corto = lambda x: "(vacío)" if pd.isna(x) else (str(x) if len(str(x)) <= 30 else str(x)[:29] + "…")

positivos = df.loc[df[MONTO] > 0, MONTO] if MONTO else pd.Series(dtype=float)
HAY_MONTO = len(positivos) > 0
# Campos que identifican a un registro: se muestran al pasar el mouse sobre un punto
CLAVES = (A.ids + A.codigos)[:3] or list(crudo.columns[:2])
hover_claves = "<br>".join(f"{esc(c)}: %{{customdata[{i}]}}" for i, c in enumerate(CLAVES))

graficos = []   # (titulo, explicacion, figura)
def agregar(titulo, explicacion, fig, alto=430):
    fig.update_layout(template="plotly_white", height=alto, separators=",.", barcornerradius=4,
                      font=dict(family="Segoe UI, Arial, sans-serif", color="#52514e", size=13),
                      paper_bgcolor="white", plot_bgcolor="white", colorway=SERIES,
                      hoverlabel=dict(bgcolor="white", font_size=13),
                      margin=dict(l=60, r=20, t=40, b=60))
    fig.update_xaxes(gridcolor="#e1e0d9", linecolor="#c3c2b7", zerolinecolor="#c3c2b7")
    fig.update_yaxes(gridcolor="#e1e0d9", linecolor="#c3c2b7", zerolinecolor="#c3c2b7")
    graficos.append((titulo, explicacion, fig))

# ---------------------------------------------------------------- 1. GRÁFICOS DEL IMPORTE
# Todos se arman con datos ya resumidos (conteos, cuartiles, muestras): el HTML pesa lo mismo
# con 5 mil que con 5 millones de registros.
def histograma_log(serie, nombre):
    """Histograma en escala logarítmica: barras de igual ancho en log10."""
    lg = np.log10(serie)
    if lg.max() - lg.min() < 0.3:                      # casi no hay dispersión: escala común
        conteo, b = np.histogram(serie, bins=30)
        fig = go.Figure(go.Bar(x=(b[:-1] + b[1:]) / 2, y=conteo, width=np.diff(b) * .94, marker_color=AZUL,
                               customdata=np.column_stack([b[:-1], b[1:]]),
                               hovertemplate="%{customdata[0]:,.2f} a %{customdata[1]:,.2f}<br>%{y:,} registros<extra></extra>"))
        fig.update_xaxes(title=nombre)
    else:
        conteo, b = np.histogram(lg, bins=40)
        fig = go.Figure(go.Bar(x=(b[:-1] + b[1:]) / 2, y=conteo, width=np.diff(b) * .94, marker_color=AZUL,
                               customdata=np.column_stack([10 ** b[:-1], 10 ** b[1:]]),
                               hovertemplate="%{customdata[0]:,.0f} a %{customdata[1]:,.0f}<br>%{y:,} registros<extra></extra>"))
        pot = np.arange(np.floor(lg.min()), np.ceil(lg.max()) + 1)
        fig.update_xaxes(title=f"{nombre} (escala logarítmica)", tickvals=pot,
                         ticktext=[fmt(10 ** p) if p >= 0 else f"{10 ** p:g}".replace(".", ",") for p in pot])
    fig.update_yaxes(title="Registros")
    return fig

estr = A.estr
if HAY_MONTO:
    agregar(f"Distribución de {MONTO}",
            "Cada barra es un rango de importes (solo importes mayores a 0). Pasá el mouse para ver el rango y la "
            "cantidad. Una cola larga hacia la derecha indica pocos importes muy grandes.",
            histograma_log(positivos, MONTO))

    # Pareto + estratos (un solo eje, en %)
    orden = positivos.sort_values(ascending=False).to_numpy()
    acum = orden.cumsum() / orden.sum() * 100
    pos = np.unique(np.linspace(0, len(acum) - 1, 500).astype(int))
    fig = go.Figure()
    if estr is not None:
        e = estr.sort_values("maximo", ascending=False)
        inicio = e["%_registros"].cumsum() - e["%_registros"]
        fig.add_trace(go.Bar(
            x=inicio + e["%_registros"] / 2, y=e["%_importe"], width=e["%_registros"], name="Estratos",
            marker=dict(color=AZUL, opacity=.28, line=dict(color="white", width=2)),
            text=[f"<b>{k}</b> · {v:.0f}%" for k, v in e["%_importe"].items()],
            textposition="outside", textangle=0, cliponaxis=False,
            customdata=np.column_stack([e.index, e["%_registros"], e["%_importe"], e["minimo"], e["maximo"], e["cantidad"]]),
            hovertemplate=("<b>%{customdata[0]}</b><br>%{customdata[3]:,.0f} a %{customdata[4]:,.0f}"
                           "<br>%{customdata[5]:,} registros (%{customdata[1]:.1f}%)"
                           "<br>%{customdata[2]:.1f}% del importe<extra></extra>")))
    fig.add_trace(go.Scatter(x=(pos + 1) / len(acum) * 100, y=acum[pos], mode="lines", name="Importe acumulado",
                             line=dict(color="#0d366b", width=2),
                             hovertemplate="El %{x:.1f}% de los registros<br>explica el %{y:.1f}% del importe<extra></extra>"))
    fig.add_hline(y=80, line_dash="dash", line_color=GRIS, line_width=1)
    fig.update_xaxes(title="% de registros (de mayor a menor importe)", range=[0, 100])
    fig.update_yaxes(title="% del importe", range=[0, 104])
    fig.update_layout(legend=dict(orientation="h", y=1.1), bargap=0)
    agregar("Concentración del importe (Pareto) y estratos",
            "La línea es el importe acumulado, de mayor a menor. Cada barra es un estrato: su ancho es el % de "
            "registros que tiene y su altura el % del importe que aporta. Un estrato angosto y alto justifica "
            "revisarlo completo; uno ancho y bajo se cubre con una muestra chica.", fig)

    if estr is not None:
        fig = go.Figure([
            go.Bar(x=estr.index.astype(str), y=estr["%_registros"], name="% de registros", marker_color=AZUL,
                   hovertemplate="%{x}: %{y:.1f}% de los registros<extra></extra>"),
            go.Bar(x=estr.index.astype(str), y=estr["%_importe"], name="% del importe", marker_color=NARANJA,
                   hovertemplate="%{x}: %{y:.1f}% del importe<extra></extra>")])
        fig.update_layout(barmode="group", legend=dict(orientation="h", y=1.1))
        fig.update_yaxes(title="%")
        fig.update_xaxes(title="Estrato (de menor a mayor importe)")
        agregar("Estratos sugeridos para el muestreo",
                "Cortes en la mediana, el percentil 90 y el menor entre el percentil 99 y el límite de atípicos. "
                "El último estrato suele revisarse al 100%. El reparto sugerido de la muestra entre los demás "
                "(asignación de Neyman) está en la tabla de estratos, al final.", fig)

# ---------------------------------------------------------------- 2. CATEGORÍAS
def caja_con_atipicos(campo, valores):
    """Cajas por categoría. Los atípicos se calculan sobre el logaritmo del importe: con importes tan
    asimétricos, la regla 1,5*IQR en escala lineal marcaría como atípicos a miles de importes altos normales."""
    fig, total = go.Figure(), 0
    datos = df[df[MONTO] > 0]
    for valor in valores:
        grupo = datos[datos[campo].isna()] if pd.isna(valor) else datos[datos[campo] == valor]
        if len(grupo) < 5:
            continue
        v, lv = grupo[MONTO], np.log10(grupo[MONTO])
        q1, med, q3 = lv.quantile([.25, .5, .75])
        fuera = (lv < q1 - 1.5 * (q3 - q1)) | (lv > q3 + 1.5 * (q3 - q1))
        total += int(fuera.sum())
        atip = grupo[fuera]
        if len(atip) > MAX_ATIPICOS:                       # quedarse con los más extremos
            atip = atip.loc[(np.log10(atip[MONTO]) - med).abs().nlargest(MAX_ATIPICOS).index]
        nombre = corto(valor)
        fig.add_trace(go.Box(x=[nombre], q1=[10 ** q1], median=[10 ** med], q3=[10 ** q3],
                             lowerfence=[v[~fuera].min()], upperfence=[v[~fuera].max()],
                             name=nombre, marker_color=AZUL, line_width=1.5, showlegend=False))
        if len(atip):
            fig.add_trace(go.Scatter(
                x=[nombre] * len(atip), y=atip[MONTO], mode="markers", showlegend=False,
                marker=dict(color=AZUL, size=7, symbol="circle-open"),
                customdata=crudo.loc[atip.index, CLAVES].astype(str).to_numpy(),
                hovertemplate=f"<b>Atípico</b><br>{hover_claves}<br>{esc(MONTO)}: %{{y:,.2f}}<extra></extra>"))
    fig.update_yaxes(type="log", dtick=1, tickformat=",.0f", title=f"{MONTO} (escala log)")
    return fig, total

for campo in A.categorias[:MAX_CAMPOS]:
    t = A.tablas[f"CAT_{campo}"].head(15)
    etiquetas = [corto(x) for x in t.index]
    fig = make_subplots(rows=1, cols=2 if MONTO else 1, horizontal_spacing=.18,
                        subplot_titles=("Cantidad de registros", f"Total de {MONTO}") if MONTO else ("Cantidad de registros",))
    fig.add_trace(go.Bar(y=etiquetas, x=t["cantidad"], orientation="h", marker_color=AZUL,
                         customdata=t["%_cantidad"],
                         hovertemplate="%{y}<br>%{x:,} registros (%{customdata:.1f}%)<extra></extra>"), 1, 1)
    if MONTO:
        fig.add_trace(go.Bar(y=etiquetas, x=t["total"], orientation="h", marker_color=NARANJA,
                             customdata=t["%_importe"],
                             hovertemplate="%{y}<br>%{x:,.0f} (%{customdata:.1f}% del importe)<extra></extra>"), 1, 2)
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(showlegend=False)
    nota = f" Se muestran los 15 valores principales de {len(A.tablas[f'CAT_{campo}'])}." if len(t) == 15 else ""
    agregar(f"Registros e importe por {campo}" if MONTO else f"Registros por {campo}",
            ("Si un valor pesa mucho más en importe que en cantidad, concentra registros caros: es candidato a "
             "estrato propio." if MONTO else "Cantidad de registros de cada valor.") + nota,
            fig, alto=max(330, 26 * len(t) + 130))

    if HAY_MONTO:
        fig, n_at = caja_con_atipicos(campo, list(t.index[:10]))
        if fig.data:
            agregar(f"{MONTO} por {campo}, con atípicos",
                    f"Caja = 50% central de los importes; línea = mediana. Cada círculo es un registro atípico "
                    f"({fmt(n_at)} en total): pasá el mouse para identificarlo. Si las cajas están a alturas muy "
                    f"distintas, {campo} sirve para estratificar. Hasta 10 valores y {MAX_ATIPICOS} atípicos por caja.", fig)

# ---------------------------------------------------------------- 3. FECHAS Y AÑOS
def muestra_dispersion(base):
    """Devuelve como mucho MAX_PUNTOS registros: siempre los 1.000 importes más altos + una muestra del resto."""
    if len(base) <= MAX_PUNTOS:
        return base, ""
    altos = base.nlargest(1000, MONTO)
    resto = base.drop(altos.index).sample(MAX_PUNTOS - 1000, random_state=1)
    return pd.concat([altos, resto]), (f" Se muestran {fmt(MAX_PUNTOS)} puntos: los 1.000 importes más altos y una "
                                       f"muestra al azar del resto.")

for campo in (A.fechas + A.anios)[:MAX_CAMPOS]:
    es_fecha = campo in A.fechas
    # fechas futuras o anteriores a 1900 y años imposibles ya están en los controles: acá estirarían el eje
    valido = (df[campo].between(pd.Timestamp(1900, 1, 1), A.hoy) if es_fecha
              else df[campo].between(1900, A.hoy.year + 1))
    anio = (df.loc[valido, campo].dt.year if es_fecha else df.loc[valido, campo]).astype(int)
    if anio.empty:
        continue
    g = pd.DataFrame({"anio": anio, "monto": df.loc[valido, MONTO] if MONTO else 0}).groupby("anio")["monto"]
    rango = range(int(anio.min()), int(anio.max()) + 1)
    cuenta, total = g.size().reindex(rango, fill_value=0), g.sum().reindex(rango, fill_value=0)
    fig = go.Figure(go.Bar(x=list(rango), y=cuenta.values, marker_color=AZUL, customdata=total.values,
                           hovertemplate="%{x}<br>%{y:,} registros" + (f"<br>{esc(MONTO)}: %{{customdata:,.0f}}" if MONTO else "")
                                         + "<extra></extra>"))
    fig.update_xaxes(title=f"Año ({campo})")
    fig.update_yaxes(title="Registros")
    agregar(f"Registros por año de {campo}",
            "Un pico aislado suele indicar una carga masiva, una migración de sistema o un cambio de criterio; "
            "conviene averiguar qué pasó ese año antes de muestrear.", fig)

    if HAY_MONTO:
        base, nota = muestra_dispersion(df[valido & (df[MONTO] > 0)])
        # los años son enteros: se desplazan un poco al azar para que los puntos no se apilen en una línea
        x = base[campo] if es_fecha else base[campo] + np.random.default_rng(1).uniform(-.35, .35, len(base))
        fx = "" if es_fecha else ":.0f"
        fig = go.Figure(go.Scattergl(
            x=x, y=base[MONTO], mode="markers", marker=dict(size=5, opacity=.4, color=AZUL),
            customdata=crudo.loc[base.index, CLAVES].astype(str).to_numpy(),
            hovertemplate=f"{hover_claves}<br>{esc(campo)}: %{{x{fx}}}<br>{esc(MONTO)}: %{{y:,.2f}}<extra></extra>"))
        fig.update_yaxes(type="log", dtick=1, tickformat=",.0f", title=f"{MONTO} (escala log)")
        fig.update_xaxes(title=campo)
        agregar(f"{MONTO} según {campo}",
                "Pasá el mouse sobre un punto raro para identificarlo. Podés hacer zoom arrastrando." + nota, fig)

# ---------------------------------------------------------------- 4. OTRAS MEDIDAS Y BENFORD
for campo in [c for c in A.medidas if c != MONTO][:3]:
    v = df.loc[df[campo] > 0, campo]
    if len(v) >= 20:
        agregar(f"Distribución de {campo}", "Otro campo numérico del archivo (solo valores mayores a 0).",
                histograma_log(v, campo))

if HAY_MONTO and A.benford is not None:
    b = A.benford
    ordenes = np.log10(positivos.max() / positivos.min())
    fig = go.Figure([go.Bar(x=list(b.index), y=b["observado_%"], name="Observado", marker_color=AZUL,
                            hovertemplate="Dígito %{x}: %{y:.2f}%<extra></extra>"),
                     go.Scatter(x=list(b.index), y=b["esperado_%"], name="Esperado (Benford)", mode="lines+markers",
                                line=dict(color=NARANJA, width=2), marker=dict(size=8),
                                hovertemplate="Dígito %{x}: %{y:.2f}%<extra></extra>")])
    fig.update_xaxes(title="Primer dígito", dtick=1)
    fig.update_yaxes(title="% de los registros")
    fig.update_layout(legend=dict(orientation="h", y=1.1))
    agregar("Ley de Benford (primer dígito)",
            "Un desvío es una señal a revisar, no prueba de error: los importes calculados por tabla o con topes "
            "no siguen Benford. " + ("Acá los importes abarcan menos de 2 órdenes de magnitud, así que la "
                                     "comparación es poco confiable." if ordenes < 2 else ""), fig)

# ---------------------------------------------------------------- 5. CALIDAD DE LOS DATOS (va al principio)
def celda(casos, pct):
    if casos == 0:
        return "<td class='c ok'>✓</td>"
    nivel = "alto" if pct >= UMBRAL_ALTO else "medio"
    return f"<td class='c {nivel}'><b>{fmt(casos)}</b><small>{pct_txt(pct)}</small></td>"

def chip(casos, pct):
    if casos == 0:
        return "<span class='chip ok'>✓ Sin casos</span>"
    return ("<span class='chip alto'>▲ Atención</span>" if pct >= UMBRAL_ALTO
            else "<span class='chip medio'>● Revisar</span>")

DESCRIPCIONES = {
    "negativo": "Valores menores a cero en un campo de cantidad o importe.",
    "duplicado": "El campo parece un identificador único pero tiene valores repetidos: posible doble registración.",
    "año < 1900 o futuro": "Años imposibles (incluye el año 0). Suelen ser datos mal cargados.",
    "fecha futura": "Fechas posteriores a hoy.",
    "fecha anterior a 1900": "Fechas demasiado antiguas; suelen ser valores por defecto del sistema.",
    "fila completa duplicada": "Registros idénticos en todos los campos: el archivo puede tener filas repetidas.",
}
COLUMNAS_COMP = [("vacío", "Vacío"), ("texto de relleno", "Texto de relleno"),
                 ("igual a 0", "Igual a 0"), ("no convertible", "No convertible")]
ctrl = A.tabla_ctrl
es_comp = ctrl["control"].str.startswith(tuple(k for k, _ in COLUMNAS_COMP))

filas = ""
for campo in crudo.columns:
    filas += f"<tr><th>{esc(campo)}<small>{rol[campo]}</small></th>"
    for clave, _ in COLUMNAS_COMP:
        r = ctrl[(ctrl["campo"] == campo) & ctrl["control"].str.startswith(clave)]
        filas += celda(int(r["casos"].iloc[0]), float(r["%"].iloc[0])) if len(r) else "<td class='c na'>—</td>"
    filas += "</tr>"
matriz = ("<table class='matriz'><tr><th>Campo</th>" + "".join(f"<th>{n}</th>" for _, n in COLUMNAS_COMP)
          + f"</tr>{filas}</table>")

filas = ""
for _, r in ctrl[~es_comp].iterrows():
    filas += (f"<tr><td>{chip(r['casos'], r['%'])}</td><td><b>{esc(r['campo'])}: {esc(r['control'])}</b><br>"
              f"<small>{DESCRIPCIONES.get(r['control'], '')}</small></td>"
              f"<td class='num'>{fmt(r['casos'])}</td><td class='num'>{pct_txt(r['%'])}</td></tr>")
lista = ("<table class='lista'><tr><th>Estado</th><th>Control</th><th class='num'>Casos</th>"
         f"<th class='num'>% de registros</th></tr>{filas}</table>") if filas else "<p>No hay controles de consistencia aplicables.</p>"

excluidos = ""
if MONTO:
    mm = df[MONTO]
    excluidos = (f"<h3>Qué queda fuera de los gráficos</h3><ul class='excluidos'>"
                 f"<li>Gráficos de <b>importe</b> (distribución, Pareto, estratos, cajas, dispersión, Benford): se excluyen "
                 f"{fmt(mm.isna().sum())} registros con {esc(MONTO)} vacío o inválido, {fmt((mm == 0).sum())} con valor 0 y "
                 f"{fmt((mm < 0).sum())} negativos. Esos registros no pueden salir en un muestreo por importe: "
                 f"hay que tratarlos como una población aparte.</li>"
                 f"<li>Gráficos por <b>año o fecha</b>: se excluyen los registros sin fecha válida, con fecha futura o con año imposible.</li>"
                 f"<li>Los gráficos de <b>cantidad por categoría</b> cuentan todos los registros.</li></ul>")

con_casos = int((ctrl["casos"] > 0).sum())
avisos_html = "".join(f"<p class='aviso'>{esc(a)}</p>" for a in A.avisos)
calidad_html = f"""<section><h2>Calidad de los datos</h2>
<p>Se hicieron <b>{len(ctrl)}</b> controles: <b>{len(ctrl) - con_casos}</b> sin casos y <b>{con_casos}</b> con casos a revisar.
<b>{fmt(len(A.casos))}</b> registros ({pct_txt(len(A.casos) / N * 100)}) tienen al menos un problema; el detalle está en
casos_con_problemas.txt. Un mismo registro puede aparecer en más de un control.</p>
<p class='leyenda'><span class='chip ok'>✓ Sin casos</span>
<span class='chip medio'>● Revisar: menos del {pct_txt(UMBRAL_ALTO)}</span>
<span class='chip alto'>▲ Atención: {pct_txt(UMBRAL_ALTO)} o más</span>
<span class='chip na'>— No aplica</span></p>
<h3>Completitud por campo</h3><div class='scroll'>{matriz}</div>
<h3>Consistencia</h3><div class='scroll'>{lista}</div>{excluidos}</section>"""

# ---------------------------------------------------------------- 6. TABLAS
def tabla_html(t, titulo, nota=""):
    if t is None or len(t) == 0:
        return ""
    return (f"<h3>{titulo}</h3>" + (f"<p>{nota}</p>" if nota else "")
            + "<div class='scroll'>" + t.to_html(classes="tabla", border=0, na_rep="", float_format=fmt2) + "</div>")

tablas_html = tabla_html(A.campos, "Campos detectados",
                         "El análisis se adapta a esta clasificación automática: conviene verificarla. "
                         "medida = cantidad o importe; categoría = pocos valores; código = muchos valores repetidos; "
                         "identificador = casi único.")
if MONTO:
    tablas_html += tabla_html(A.tablas.get("MARCO_MUESTREO"), "Marco de muestreo")
    tablas_html += tabla_html(estr, "Estratos sugeridos",
                              "%_muestra_neyman: cómo repartir la muestra entre los estratos que no se revisan al 100%.")
tablas_html += tabla_html(A.tablas.get("ESTADISTICA_NUMERICOS"), "Estadística de los campos numéricos")

# ---------------------------------------------------------------- 7. ARMADO DEL HTML
kpis = [("Registros", fmt(N)), ("Campos", str(len(crudo.columns)))]
if MONTO:
    kpis.append((f"Total de {MONTO}", fmt(df[MONTO].sum())))
if HAY_MONTO:
    n80 = int((acum <= 80).sum()) + 1
    kpis += [("Mediana (importe > 0)", fmt(positivos.median())),
             ("Registros que explican el 80% del importe", pct_txt(n80 / len(acum) * 100))]
kpis.append(("Registros con algún problema", pct_txt(len(A.casos) / N * 100)))
kpis_html = "".join(f"<div class='kpi'><span>{esc(k)}</span><b>{v}</b></div>" for k, v in kpis)

secciones = ""
for i, (titulo, explicacion, fig) in enumerate(graficos):
    secciones += (f"<section><h2>{esc(titulo)}</h2><p>{explicacion}</p>"
                  + fig.to_html(full_html=False, include_plotlyjs=False, div_id=f"g{i}",
                                config={"displaylogo": False, "responsive": True})
                  + "</section>")

pagina = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Informe de {esc(os.path.basename(A.args.archivo))}</title>
<script>{get_plotlyjs()}</script>
<style>
 body {{font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f4f5f7; color: #0b0b0b;}}
 header {{background: #0d366b; color: white; padding: 24px 32px;}}
 header h1 {{margin: 0; font-size: 24px;}}
 header p {{margin: 6px 0 0; opacity: .85;}}
 main {{max-width: 1150px; margin: auto; padding: 20px 16px;}}
 .kpis {{display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 20px;}}
 .kpi {{flex: 1 1 170px; background: white; border-radius: 8px; padding: 14px 16px; box-shadow: 0 1px 3px rgba(0,0,0,.1);}}
 .kpi span {{display: block; font-size: 13px; color: #52514e;}}
 .kpi b {{font-size: 22px;}}
 section {{background: white; border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1);}}
 section h2 {{margin: 0 0 4px; font-size: 19px;}}
 section h3 {{font-size: 16px; margin: 18px 0 8px;}}
 section p {{margin: 0 0 8px; color: #52514e; font-size: 14px; line-height: 1.45;}}
 .tabla {{border-collapse: collapse; font-size: 13px; margin-bottom: 16px;}}
 .tabla th, .tabla td {{padding: 5px 10px; border-bottom: 1px solid #e1e0d9; text-align: right; white-space: nowrap;}}
 .tabla th:first-child, .tabla td:first-child {{text-align: left;}}
 .aviso {{background: #fff1cc; color: #5c4100; padding: 10px 12px; border-radius: 6px;}}
 .scroll {{overflow-x: auto;}}
 .leyenda {{display: flex; flex-wrap: wrap; gap: 8px;}}
 .chip {{display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; white-space: nowrap;}}
 .ok {{background: #e3f4e6; color: #1e7b34;}}
 .medio {{background: #fff1cc; color: #8a6100;}}
 .alto {{background: #fde0de; color: #b42318;}}
 .na {{background: #f1f1f1; color: #898781;}}
 .matriz, .lista {{border-collapse: separate; border-spacing: 4px; font-size: 14px; width: 100%;}}
 .matriz th {{text-align: left; font-weight: 600; padding: 4px 8px; color: #52514e;}}
 .matriz th small {{display: block; font-weight: 400; font-size: 11px; color: #898781;}}
 .matriz td.c {{text-align: center; border-radius: 6px; padding: 6px 8px; min-width: 110px;}}
 .matriz td small {{display: block; font-size: 11px;}}
 .lista {{border-spacing: 0;}}
 .lista th {{text-align: left; font-size: 12px; color: #52514e; padding: 6px 8px; border-bottom: 2px solid #e1e0d9;}}
 .lista td {{padding: 8px; border-bottom: 1px solid #eee; vertical-align: top;}}
 .lista td small {{color: #52514e;}}
 .lista .num {{text-align: right; white-space: nowrap;}}
 .excluidos {{font-size: 14px; color: #52514e; line-height: 1.5;}}
</style></head>
<body>
<header><h1>Análisis exploratorio para muestreo</h1>
<p>Archivo: {esc(os.path.basename(A.args.archivo))} · {fmt(N)} registros · Campo de importe: {esc(MONTO or 'no se detectó ninguno')}
 · Generado el {pd.Timestamp.today():%d/%m/%Y %H:%M}</p></header>
<main>
<div class="kpis">{kpis_html}</div>
{avisos_html}
{calidad_html}
{secciones}
<section><h2>Tablas</h2>{tablas_html}</section>
</main></body></html>"""

with open(SALIDA, "w", encoding="utf-8") as f:
    f.write(pagina)
print(f"Informe generado: {SALIDA} ({os.path.getsize(SALIDA) / 1e6:.1f} MB)")
