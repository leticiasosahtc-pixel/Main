"""
Análisis exploratorio genérico de un archivo de texto delimitado, orientado a preparar un muestreo.
No necesita conocer las columnas: detecta separador, tipo y rol de cada campo y adapta el análisis.

Uso:   python analisis_generico.py archivo.txt [--monto COLUMNA] [--sep "|"]
         --monto  campo de importe sobre el que se va a muestrear (si se omite, se elige uno y se avisa)
         --sep    separador (si se omite, se detecta entre | ; tabulación y coma)
Requiere: pip install pandas openpyxl      (matplotlib es opcional: sin él no se generan los gráficos)
Salida: carpeta 'salida_analisis' con resumen.txt, tablas_analisis.xlsx, casos_con_problemas.txt y gráficos
"""
import argparse
import math
import os
import re
import numpy as np
import pandas as pd

ap = argparse.ArgumentParser()
ap.add_argument("archivo")
ap.add_argument("--monto")
ap.add_argument("--sep")
args = ap.parse_args()

SALIDA = "salida_analisis"
os.makedirs(SALIDA, exist_ok=True)
RELLENO = {"NULL", "NAN", "NONE", "N/A", "NA", "S/D", "SD", "SIN DATO", "-", "--", "."}
UMBRAL = 0.90            # proporción mínima de valores convertibles para considerar un campo número o fecha
MAX_CATEGORIAS = 50      # hasta cuántos valores distintos un campo se trata como categoría
PISTA_MONTO = re.compile(r"MONTO|IMPORTE|IMPUESTO|TOTAL|VALOR|SALDO|PRECIO|FACTUR|DEUDA|PAGO|NETO|BRUTO")
PISTA_CODIGO = re.compile(r"^(ID|COD|NRO|NUM|DNI|CUIT|CUIL|TEL|CP)|_ID$|CODIGO|NUMERO")
hoy = pd.Timestamp.today()

resumen = []
def log(texto=""):
    print(texto)
    resumen.append(str(texto))

def ruta_salida(nombre):
    """Ruta dentro de la carpeta de salida. Si el archivo está abierto en otro programa (Excel lo bloquea),
    usa un nombre alternativo en vez de cortar el análisis al final."""
    ruta = os.path.join(SALIDA, nombre)
    try:
        open(ruta, "a").close()
    except PermissionError:
        base, ext = os.path.splitext(nombre)
        ruta = os.path.join(SALIDA, f"{base}_{pd.Timestamp.today():%Y%m%d_%H%M%S}{ext}")
        log(f"\nATENCIÓN: {nombre} está abierto en otro programa y no se puede reemplazar. "
            f"Se guarda como {os.path.basename(ruta)}. Cerrar el archivo antes de volver a correr.")
    return ruta

def tabla(t):
    return t.to_string() if len(t) else "(sin datos)"

# ---------------------------------------------------------------- 1. CARGA
def leer(ruta, sep):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, encoding=enc) as f:
                cabecera = f.readline()
            sep = sep or max("|;\t,", key=cabecera.count)
            return pd.read_csv(ruta, sep=sep, dtype=str, encoding=enc, keep_default_na=False), sep
        except UnicodeDecodeError:
            continue

df, SEP = leer(args.archivo, args.sep)
df.columns = df.columns.str.strip().str.upper()
df = df.apply(lambda c: c.str.strip()).replace("", pd.NA)
crudo = df.copy()                                   # valores originales como texto
N = len(df)
es_relleno = crudo.apply(lambda c: c.str.upper().isin(RELLENO))
con_dato = crudo.notna() & ~es_relleno

log("=" * 70)
log(f"ARCHIVO: {args.archivo}   separador: {SEP!r}")
log(f"Registros: {N:,}   Columnas: {len(df.columns)}")
log("=" * 70)
if N == 0:
    raise SystemExit("El archivo no tiene registros.")

# ---------------------------------------------------------------- 2. TIPO Y ROL DE CADA CAMPO
avisos = []

def a_numero(s, campo=None):
    """Acepta 1234.56, 1.234,56 y 1,234.56 (decide por la posición de la coma respecto del punto)."""
    s = s.str.replace(r"[\s$]", "", regex=True)
    coma = s.str.contains(",", na=False, regex=False)
    if coma.any():
        europeo = (s[coma].str.rfind(",") > s[coma].str.rfind(".")).mean() > 0.5
        s = (s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False) if europeo
             else s.str.replace(",", "", regex=False))
    elif campo and s.str.fullmatch(r"-?\d{1,3}(\.\d{3})+", na=False).mean() > 0.5:
        avisos.append(f"{campo}: los puntos parecen separadores de miles (ej. 1.234) pero se leyeron como "
                      f"decimales. Verificar el formato antes de usar los totales.")
    return pd.to_numeric(s, errors="coerce")

def a_fecha(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True, format="mixed")

tipo, rol, numerico = {}, {}, {}
for c in crudo.columns:
    v = crudo.loc[con_dato[c], c]
    muestra = v.sample(min(len(v), 5000), random_state=1)
    if len(v) == 0:
        tipo[c], rol[c] = "vacío", "vacío"
        continue
    distintos, unico = v.nunique(), v.nunique() / len(v)
    if a_numero(muestra).notna().mean() >= UMBRAL:
        tipo[c] = "número"
        num = a_numero(v, c)
        numerico[c] = a_numero(crudo[c])
        entero = bool((num.dropna() % 1 == 0).all())
        ceros_izq = v.str.fullmatch(r"0\d+").mean() > 0.05                 # 00123 es un código, no una cantidad
        largo_fijo = entero and v.str.len().nunique() == 1 and v.str.len().iloc[0] >= 6
        if entero and unico >= 0.99 and distintos > 1:
            rol[c] = "identificador"
        elif entero and num.between(1900, hoy.year + 1).mean() >= 0.95:
            rol[c] = "año"
        elif entero and distintos <= 20:
            rol[c] = "categoría"
        elif entero and (ceros_izq or largo_fijo or PISTA_CODIGO.search(c)):
            rol[c] = "código"
        else:
            rol[c] = "medida"
    elif a_fecha(muestra).notna().mean() >= UMBRAL:
        tipo[c], rol[c] = "fecha", "fecha"
        df[c] = a_fecha(crudo[c])
    else:
        tipo[c] = "texto"
        rol[c] = ("identificador" if unico >= 0.99 and distintos > 1
                  else "categoría" if distintos <= min(MAX_CATEGORIAS, len(v) / 2) else "código")

medidas = [c for c in df.columns if rol[c] == "medida"]
categorias = [c for c in df.columns if rol[c] == "categoría"]
fechas = [c for c in df.columns if rol[c] == "fecha"]
anios = [c for c in df.columns if rol[c] == "año"]
ids = [c for c in df.columns if rol[c] == "identificador"]
codigos = [c for c in df.columns if rol[c] == "código"]

# Campo de importe para el muestreo: el indicado, o el que más parece un importe
MONTO = args.monto.strip().upper() if args.monto else None
if MONTO and MONTO not in df.columns:
    raise SystemExit(f"--monto {MONTO}: no existe. Columnas: {', '.join(df.columns)}")
if MONTO and tipo[MONTO] != "número":
    raise SystemExit(f"--monto {MONTO}: el campo no es numérico (se detectó como {tipo[MONTO]}).")
if MONTO and MONTO not in medidas:
    rol[MONTO] = "medida"
    medidas.append(MONTO)
elif not MONTO and medidas:
    def puntaje(c):
        v = numerico[c].dropna()
        return (bool(PISTA_MONTO.search(c)), bool((v % 1 != 0).any()), v.abs().sum())
    MONTO = max(medidas, key=puntaje)
    avisos.append(f"Campo de importe elegido automáticamente: {MONTO}. Si no es el correcto, "
                  f"volver a correr con --monto COLUMNA.")

# Solo las medidas y los años se usan como número; categorías, códigos e identificadores quedan como texto
# (así "01" no se convierte en 1).
for c in medidas + anios:
    df[c] = numerico[c]
pd.options.display.float_format = "{:,.2f}".format

campos = pd.DataFrame({
    "tipo": tipo, "rol": rol,
    "con_dato": con_dato.sum(), "%_con_dato": (con_dato.mean() * 100).round(2),
    "distintos": {c: crudo.loc[con_dato[c], c].nunique() for c in crudo.columns},
    "ejemplo": {c: (crudo.loc[con_dato[c], c].iloc[0][:30] if con_dato[c].any() else "") for c in crudo.columns}})
log("\n--- CAMPOS DETECTADOS ---")
log(tabla(campos))
log("(rol: medida = cantidad/importe; categoría = pocos valores; código = muchos valores repetidos; "
    "identificador = casi único)")

# ---------------------------------------------------------------- 3. CALIDAD DE DATOS
# Cada control se lista aunque dé 0 casos, para dejar constancia de que se hizo.
controles = []      # (campo, control, máscara)
for c in crudo.columns:
    controles.append((c, "vacío", crudo[c].isna()))
    controles.append((c, "texto de relleno (NULL, S/D, -, etc.)", es_relleno[c]))
    if tipo[c] == "número":
        controles.append((c, "no convertible a número", con_dato[c] & numerico[c].isna()))
        if rol[c] != "categoría":
            controles.append((c, "igual a 0", numerico[c] == 0))
    if tipo[c] == "fecha":
        controles.append((c, "no convertible a fecha", con_dato[c] & df[c].isna()))
    if rol[c] == "medida":
        controles.append((c, "negativo", df[c] < 0))
    if rol[c] == "año":
        controles.append((c, "año < 1900 o futuro", (df[c] < 1900) | (df[c] > hoy.year + 1)))
    if rol[c] == "fecha":
        controles.append((c, "fecha futura", df[c] > hoy))
        controles.append((c, "fecha anterior a 1900", df[c] < pd.Timestamp(1900, 1, 1)))
    if rol[c] == "identificador":
        controles.append((c, "duplicado", con_dato[c] & crudo[c].duplicated(keep=False)))
controles.append(("(registro)", "fila completa duplicada", crudo.duplicated(keep=False)))
controles = [(c, n, mk.fillna(False).astype(bool)) for c, n, mk in controles]

tabla_ctrl = pd.DataFrame([{"campo": c, "control": n, "casos": int(mk.sum()),
                            "%": round(mk.mean() * 100, 2)} for c, n, mk in controles])
log("\n--- CALIDAD DE DATOS ---")
con_casos = tabla_ctrl[tabla_ctrl["casos"] > 0]
log(con_casos.to_string(index=False) if len(con_casos) else "(ningún control con casos)")
log(f"({len(tabla_ctrl)} controles realizados; se muestran los que tienen casos. Detalle completo en el Excel.)")

problema = pd.Series("", index=df.index)
for c, n, mk in controles:
    problema.loc[mk] = problema.loc[mk] + f"{c}: {n}; "
casos = crudo[problema != ""].copy()
casos.insert(0, "PROBLEMA", problema[problema != ""].str.rstrip("; "))
casos.to_csv(ruta_salida("casos_con_problemas.txt"), sep="|", index=False, encoding="utf-8")
log(f"Registros con al menos un problema: {len(casos):,} de {N:,} ({len(casos) / N * 100:.1f}%)")

# ---------------------------------------------------------------- 4. ESTADÍSTICA DE LOS CAMPOS NUMÉRICOS
tablas = {}
cuantitativos = medidas + anios
if cuantitativos:
    est = df[cuantitativos].describe(percentiles=[.01, .05, .25, .5, .75, .9, .95, .99]).T
    est["suma"] = df[cuantitativos].sum()
    est["coef_variacion"] = est["std"] / est["mean"]
    est["asimetria"] = df[cuantitativos].skew()
    tablas["ESTADISTICA_NUMERICOS"] = est
    log("\n--- CAMPOS NUMÉRICOS ---")
    log(est.T.to_string())
for c in fechas:
    log(f"\n{c}: de {df[c].min():%d/%m/%Y} a {df[c].max():%d/%m/%Y}")
    tablas[f"POR_AÑO_{c}"] = df[c].dt.year.value_counts().sort_index().rename_axis("año").to_frame("cantidad")

# ---------------------------------------------------------------- 5. ANÁLISIS DEL IMPORTE (base del muestreo)
m = acum = benford = estr = None
if MONTO:
    m = df[MONTO].dropna()
    total = m.sum()
    log(f"\n--- IMPORTE PARA EL MUESTREO: {MONTO} ---")
    log(f"Total: {total:,.2f}   Promedio: {m.mean():,.2f}   Mediana: {m.median():,.2f}   Desvío: {m.std():,.2f}")
    log(f"Coef. de variación: {m.std() / m.mean():.2f}   Asimetría: {m.skew():.2f}   "
        f"(>1 = pocos importes grandes: conviene estratificar)")
    q1, q3 = m.quantile([.25, .75])
    lim_sup = q3 + 3 * (q3 - q1)
    atipicos = df[df[MONTO] > lim_sup]
    tablas["ATIPICOS"] = crudo.loc[atipicos.sort_values(MONTO, ascending=False).index].head(100_000)
    log(f"Atípicos extremos (> Q3 + 3*IQR = {lim_sup:,.2f}): {len(atipicos):,} registros, "
        f"{atipicos[MONTO].sum() / total * 100:.1f}% del importe")

    pos = m[m > 0]
    if len(pos):
        orden = pos.sort_values(ascending=False).reset_index(drop=True)
        acum = orden.cumsum() / orden.sum()
        n80 = int((acum <= 0.8).sum()) + 1
        log(f"El 80% del importe lo explican {n80:,} registros ({n80 / len(orden) * 100:.1f}% de los positivos)")

        # Benford: solo es informativo si los importes abarcan varios órdenes de magnitud
        d1 = np.floor(pos / 10 ** np.floor(np.log10(pos))).astype(int).clip(1, 9)
        benford = pd.DataFrame({"observado_%": d1.value_counts(normalize=True).reindex(range(1, 10), fill_value=0) * 100})
        benford.index.name = "dígito"
        benford["esperado_%"] = [100 * math.log10(1 + 1 / d) for d in benford.index]
        benford["diferencia"] = benford["observado_%"] - benford["esperado_%"]
        tablas["BENFORD"] = benford.round(2)
        ordenes = math.log10(pos.max() / pos.min())
        log(f"\nBenford (primer dígito). Rango de {ordenes:.1f} órdenes de magnitud"
            + (" -> poco confiable (se necesitan al menos 2-3)" if ordenes < 2 else ""))
        log(benford.round(2).to_string())

    # ------------------------------------------------------------ 6. MARCO Y ESTRATOS PARA EL MUESTREO
    log("\n--- MARCO DE MUESTREO ---")
    marco = pd.DataFrame({"registros": [N, m.gt(0).sum(), m.eq(0).sum(), m.lt(0).sum(), N - len(m)],
                          "importe": [total, pos.sum(), 0, m[m < 0].sum(), None]},
                         index=["Población total", "Importe > 0 (muestreable por importe)", "Importe = 0",
                                "Importe negativo", "Importe vacío o inválido"])
    marco["%_registros"] = (marco["registros"] / N * 100).round(2)
    tablas["MARCO_MUESTREO"] = marco
    log(marco.to_string())
    log("Los ceros, negativos y vacíos nunca salen en un muestreo por importe: tratarlos como población aparte.")

    if len(pos) >= 20:
        validos = df[df[MONTO] > 0].copy()
        cortes = sorted({0, pos.quantile(.50), pos.quantile(.90), min(lim_sup, pos.quantile(.99)), float("inf")})
        etiquetas = [f"E{i + 1}" for i in range(len(cortes) - 1)]
        validos["ESTRATO"] = pd.cut(validos[MONTO], bins=cortes, labels=etiquetas, include_lowest=True)
        estr = validos.groupby("ESTRATO", observed=True)[MONTO].agg(
            cantidad="count", total="sum", minimo="min", maximo="max", promedio="mean", desvio="std")
        estr["%_registros"] = (estr["cantidad"] / estr["cantidad"].sum() * 100).round(2)
        estr["%_importe"] = (estr["total"] / estr["total"].sum() * 100).round(2)
        # Asignación de Neyman: reparte la muestra según cantidad x desvío (sin el último estrato, que va al 100%)
        peso = (estr["cantidad"] * estr["desvio"].fillna(0)).iloc[:-1]
        estr["%_muestra_neyman"] = (peso / peso.sum() * 100).round(1) if peso.sum() else None
        tablas["ESTRATOS"] = estr
        log("\n--- ESTRATOS SUGERIDOS (por importe; cortes en p50, p90 y min(p99, atípicos)) ---")
        log(estr.round(2).to_string())
        log("(El último estrato suele revisarse al 100%. %_muestra_neyman indica cómo repartir la muestra "
            "entre los demás.)")

# ---------------------------------------------------------------- 7. CATEGORÍAS (posibles variables de estratificación)
if categorias:
    log("\n--- DISTRIBUCIÓN POR CATEGORÍAS ---")
for c in categorias:
    g = df.groupby(c, dropna=False)
    t = g.size().to_frame("cantidad")
    t["%_cantidad"] = (t["cantidad"] / N * 100).round(2)
    if MONTO:
        t = t.join(g[MONTO].agg(total="sum", promedio="mean", mediana="median", desvio="std",
                                minimo="min", maximo="max"))
        t["%_importe"] = (t["total"] / t["total"].sum() * 100).round(2)
    t = t.sort_values("total" if MONTO else "cantidad", ascending=False)
    tablas[f"CAT_{c}"] = t
    log(f"\nPor {c} ({len(t)} valores):")
    log(t.round(2).head(30).to_string())
    chicas = int((t["cantidad"] < 30).sum())
    if chicas:
        log(f"  Atención: {chicas} categoría(s) con menos de 30 registros; agruparlas si se estratifica por {c}.")

# Códigos (sujetos, proveedores, etc.): concentración por valor
for c in codigos:
    g = df.groupby(c)
    t = g.size().to_frame("registros")
    if MONTO:
        t["total"] = g[MONTO].sum()
    t = t.sort_values("total" if MONTO else "registros", ascending=False)
    tablas[f"TOP_{c}"] = t.head(50)
    txt = f"\n{c}: {len(t):,} valores distintos; registros por valor: promedio {t['registros'].mean():.2f}, " \
          f"máximo {t['registros'].max():,}"
    if MONTO and t["total"].sum():
        txt += f"; los 10 mayores concentran el {t['total'].head(10).sum() / t['total'].sum() * 100:.1f}% del importe"
    log(txt)

if avisos:
    log("\n--- AVISOS ---")
    for a in avisos:
        log("* " + a)

# ---------------------------------------------------------------- 8. GRÁFICOS (opcionales)
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except ImportError:
    plt = None
    log("\n(matplotlib no está instalado: no se generan gráficos. pip install matplotlib)")

def guardar(nombre):
    plt.tight_layout()
    plt.savefig(os.path.join(SALIDA, re.sub(r"[^\w.-]", "_", nombre)), dpi=120)
    plt.close("all")

if plt:
    if MONTO and acum is not None:
        plt.figure(figsize=(8, 5))
        plt.hist(np.log10(m[m > 0]), bins=60)
        plt.title(f"Distribución de {MONTO} (escala log10)"); plt.xlabel(f"log10({MONTO})"); plt.ylabel("Registros")
        guardar("01_histograma_importe_log.png")

        plt.figure(figsize=(7, 5))
        plt.plot(np.arange(1, len(acum) + 1) / len(acum) * 100, acum * 100)
        plt.axhline(80, ls="--", c="gray")
        plt.title("Curva de concentración (Pareto)"); plt.xlabel("% de registros"); plt.ylabel("% del importe")
        guardar("02_pareto.png")

        benford[["observado_%", "esperado_%"]].plot.bar(figsize=(7, 4), title="Ley de Benford - primer dígito")
        guardar("03_benford.png")
    for c in categorias:
        t = tablas[f"CAT_{c}"].head(20)
        t.index = t.index.map(lambda x: "(vacío)" if pd.isna(x) else str(x)[:30])
        fig, ax = plt.subplots(1, 2 if MONTO else 1, figsize=(12 if MONTO else 7, 5), squeeze=False)
        t["cantidad"].plot.barh(ax=ax[0][0], title=f"Cantidad por {c}")
        ax[0][0].invert_yaxis()
        if MONTO:
            t["total"].plot.barh(ax=ax[0][1], title=f"{MONTO} por {c}", color="orange")
            ax[0][1].invert_yaxis()
        guardar(f"04_barras_{c}.png")
    for c in fechas + anios:
        serie = df[c].dt.year if c in fechas else df[c].where(df[c].between(1900, hoy.year + 1))
        cuenta = serie.dropna().astype(int).value_counts().sort_index()
        if len(cuenta):
            cuenta.plot.bar(figsize=(max(7, min(18, len(cuenta) * 0.3)), 4), title=f"Registros por año ({c})")
            guardar(f"05_por_anio_{c}.png")

# ---------------------------------------------------------------- 9. EXPORTAR
usadas = set()
def hoja(nombre):
    base = re.sub(r"[\[\]:*?/\\]", "_", nombre)[:31]
    final, i = base, 2
    while final.upper() in usadas:
        final = f"{base[:28]}_{i}"; i += 1
    usadas.add(final.upper())
    return final

with pd.ExcelWriter(ruta_salida("tablas_analisis.xlsx"), engine="openpyxl") as xl:
    campos.to_excel(xl, sheet_name=hoja("CAMPOS"))
    tabla_ctrl.to_excel(xl, sheet_name=hoja("CONTROLES"), index=False)
    for nombre, t in tablas.items():
        t.to_excel(xl, sheet_name=hoja(nombre), index=nombre != "ATIPICOS")
    for c in crudo.columns:      # valores más frecuentes de cada campo
        vc = crudo[c].fillna("(vacío)").value_counts().head(30).rename_axis("valor").reset_index(name="cantidad")
        vc["%"] = (vc["cantidad"] / N * 100).round(2)
        vc.to_excel(xl, sheet_name=hoja("FREC_" + c), index=False)
    casos.head(1_000_000).to_excel(xl, sheet_name=hoja("CASOS_CON_PROBLEMAS"), index=False)

with open(ruta_salida("resumen.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(resumen))
print(f"\nListo. Resultados en la carpeta '{SALIDA}'")
