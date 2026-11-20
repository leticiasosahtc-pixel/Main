"""
Análisis exploratorio genérico de un archivo de texto delimitado, orientado a preparar un muestreo.
Hace el mismo análisis que analisis_bd.py, pero sin conocer las columnas de antemano: detecta separador,
tipo y rol de cada campo y adapta los controles, tablas y gráficos.

Uso:   python analisis_generico.py archivo.txt [opciones]
  --monto COLUMNA        campo de importe sobre el que se va a muestrear
  --sujeto COLUMNA       campo que identifica al sujeto / contribuyente / titular
  --categorias A,B       categorías principales: A es el grupo (ej. CLASE) y B la subcategoría (ej. CATEGORIA)
  --anio COLUMNA         año de origen (ej. AÑO_FABRICACION) ...
  --fecha COLUMNA        ... y fecha de alta (ej. FECHA_ALTA): habilita los controles y gráficos de antigüedad
  --duplicados A,B       campos que no pueden repetirse (ej. PADRON): se controlan duplicados aunque el campo no
                         se detecte como identificador
  --ignorar A,B          campos que se quitan del análisis (no se analizan ni aparecen en ninguna salida)
  --excluir A,B          campos cuyos problemas no se cuentan en el segundo listado de casos (por defecto, el sujeto)
  --sep "|"              separador (si se omite, se detecta entre | ; tabulación y coma)
Lo que no se indica se detecta automáticamente y se avisa en el resumen.

Requiere: pip install pandas openpyxl matplotlib      (sin matplotlib no se generan los .png)
Salida: carpeta 'salida_analisis' con resumen.txt, tablas_analisis.xlsx, casos_con_problemas*.txt y gráficos
"""
import argparse
import math
import os
import re
import numpy as np
import pandas as pd

ap = argparse.ArgumentParser()
ap.add_argument("archivo")
for opcion in ("--monto", "--sujeto", "--categorias", "--anio", "--fecha", "--excluir", "--ignorar", "--duplicados", "--sep"):
    ap.add_argument(opcion)
args = ap.parse_args()

SALIDA = "salida_analisis"
os.makedirs(SALIDA, exist_ok=True)
RELLENO = {"NULL", "NAN", "NONE", "N/A", "NA", "S/D", "SD", "SIN DATO", "-", "--", "."}
UMBRAL = 0.90            # proporción mínima de valores convertibles para considerar un campo número o fecha
MAX_CATEGORIAS = 50      # hasta cuántos valores distintos un campo se trata como categoría
EXCEL_MAX = 1_000_000    # Excel admite como máximo 1.048.576 filas por hoja
PISTA_MONTO = re.compile(r"MONTO|IMPORTE|IMPUESTO|TOTAL|VALOR|SALDO|PRECIO|FACTUR|DEUDA|PAGO|NETO|BRUTO")
PISTA_CODIGO = re.compile(r"^(ID|COD|NRO|NUM|DNI|CUIT|CUIL|TEL|CP)|_ID$|CODIGO|NUMERO")
PISTA_SUJETO = re.compile(r"SUJETO|CUIT|CUIL|CONTRIBUY|TITULAR|CLIENTE|PROVEED|INSCRIP|DNI")
PISTA_ORIGEN = re.compile(r"FABRIC|MODELO|CONSTRU|NACIM|ORIGEN|CONSTITU")    # años que marcan un "inicio"
PISTA_ALTA = re.compile(r"ALTA|INSCRIP|INICIO|REGISTR|RADIC")
hoy = pd.Timestamp.today()

resumen = []
def log(texto=""):
    print(texto)
    resumen.append(str(texto))

def ruta_salida(nombre):
    """Ruta dentro de la carpeta de salida. Si el archivo está abierto en otro programa (Excel lo bloquea),
    usa un nombre alternativo en vez de cortar el análisis al final."""
    nombre = re.sub(r"[^\w.-]", "_", nombre)            # un "/" en el nombre de un campo (ej. CUIT/DNI) crearía carpetas
    ruta = os.path.join(SALIDA, nombre)
    try:
        open(ruta, "a").close()
    except PermissionError:
        base, ext = os.path.splitext(nombre)
        ruta = os.path.join(SALIDA, f"{base}_{pd.Timestamp.today():%Y%m%d_%H%M%S}{ext}")
        log(f"\nATENCIÓN: {nombre} está abierto en otro programa y no se puede reemplazar. "
            f"Se guarda como {os.path.basename(ruta)}. Cerrar el archivo antes de volver a correr.")
    return ruta

def excel(t):
    """Excel no admite caracteres de control (salvo tabulación y salto de línea): en la copia que va al Excel se
    reemplazan por '\ufffd'. Los .txt conservan el valor original."""
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    limpiar = lambda v: ILLEGAL_CHARACTERS_RE.sub("\ufffd", v) if isinstance(v, str) else v
    texto = lambda s: not (pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s))
    t = t.copy()
    for c in t.columns:
        if texto(t[c]):
            t[c] = t[c].map(limpiar)
    if texto(t.index):
        t.index = t.index.map(limpiar)
    t.columns = t.columns.map(limpiar)
    return t

# ---------------------------------------------------------------- 1. CARGA
def tiene_nulos(ruta):
    """True si el archivo tiene caracteres nulos (byte 0)."""
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 24), b""):
            if b"\x00" in bloque:
                return True
    return False

def leer(ruta, sep):
    nulos_archivo = tiene_nulos(ruta)
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, encoding=enc) as f:
                cabecera = f.readline()
            sep = sep or max("|;\t,", key=cabecera.count)
            fuente = ruta
            if nulos_archivo:
                # pandas corta el texto en el primer carácter nulo y pierde el resto del valor sin avisar:
                # se reemplazan por '\ufffd' antes de leer, y el control de caracteres de control los cuenta
                import io
                with open(ruta, encoding=enc) as f:
                    texto = f.read()
                avisos_carga.append(f"El archivo tiene {texto.count(chr(0)):,} caracteres nulos (byte 0); se reemplazaron "
                                    "por '\ufffd' para no perder el resto del valor. Ver el control 'caracteres de control'.")
                fuente = io.StringIO(texto.replace("\x00", "\ufffd"))
            return pd.read_csv(fuente, sep=sep, dtype=str, keep_default_na=False,
                               **({} if nulos_archivo else {"encoding": enc})), sep
        except UnicodeDecodeError:
            continue

avisos_carga = []

df, SEP = leer(args.archivo, args.sep)
df.columns = df.columns.str.strip().str.upper()
IGNORADOS = [x.strip().upper() for x in (args.ignorar or "").split(",") if x.strip()]
faltan = [c for c in IGNORADOS if c not in df.columns]
if faltan:
    raise SystemExit(f"--ignorar {', '.join(faltan)}: no existe. Columnas: {', '.join(df.columns)}")
df = df.drop(columns=IGNORADOS)
if df.columns.empty:
    raise SystemExit("--ignorar: no queda ninguna columna para analizar.")
df = df.apply(lambda c: c.str.strip()).replace("", pd.NA)
crudo = df.copy()                                   # valores originales como texto (para los controles)
N = len(df)
es_relleno = crudo.apply(lambda c: c.str.upper().isin(RELLENO))
con_dato = crudo.notna() & ~es_relleno

log("=" * 70)
log(f"ARCHIVO: {args.archivo}   separador: {SEP!r}")
log(f"Registros: {N:,}   Columnas: {len(df.columns)}"
    + (f"   (se ignoran: {', '.join(IGNORADOS)})" if IGNORADOS else ""))
log("=" * 70)
if N == 0:
    raise SystemExit("El archivo no tiene registros.")

# ---------------------------------------------------------------- 2. TIPO Y ROL DE CADA CAMPO
avisos = list(avisos_carga)

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

def columna(valor, opcion, tipo_requerido=None):
    """Valida una columna indicada por parámetro."""
    if not valor:
        return None
    c = valor.strip().upper()
    if c not in df.columns:
        raise SystemExit(f"{opcion} {c}: no existe. Columnas: {', '.join(df.columns)}")
    if tipo_requerido and tipo[c] != tipo_requerido:
        raise SystemExit(f"{opcion} {c}: tiene que ser de tipo {tipo_requerido} (se detectó como {tipo[c]}).")
    return c

# Lo indicado por parámetro manda sobre la detección automática
MONTO = columna(args.monto, "--monto", "número")
ANIO_ORIGEN = columna(args.anio, "--anio", "número")
FECHA_ALTA = columna(args.fecha, "--fecha", "fecha")
if bool(ANIO_ORIGEN) != bool(FECHA_ALTA):
    raise SystemExit("--anio y --fecha se indican juntos (año de origen y fecha de alta).")
DUPLICADOS = [columna(x, "--duplicados") for x in (args.duplicados or "").split(",") if x.strip()]
ORDEN_CAT = [columna(x, "--categorias") for x in (args.categorias or "").split(",") if x.strip()]
if MONTO:
    rol[MONTO] = "medida"
if ANIO_ORIGEN:
    rol[ANIO_ORIGEN] = "año"
for c in ORDEN_CAT:
    rol[c] = "categoría"

medidas = [c for c in df.columns if rol[c] == "medida"]
categorias = ORDEN_CAT + [c for c in df.columns if rol[c] == "categoría" and c not in ORDEN_CAT]
fechas = [c for c in df.columns if rol[c] == "fecha"]
anios = [c for c in df.columns if rol[c] == "año"]
ids = [c for c in df.columns if rol[c] == "identificador"]
codigos = [c for c in df.columns if rol[c] == "código"]

# Campo de importe: el indicado, o el que más parece un importe
if not MONTO and medidas:
    def puntaje(c):
        v = numerico[c].dropna()
        return (bool(PISTA_MONTO.search(c)), bool((v % 1 != 0).any()), v.abs().sum())
    MONTO = max(medidas, key=puntaje)
    avisos.append(f"Campo de importe elegido automáticamente: {MONTO}. Si no es el correcto: --monto COLUMNA.")

# Sujeto (equivale a ID_SUJETO en analisis_bd.py)
SUJETO = columna(args.sujeto, "--sujeto")
if not SUJETO:
    SUJETO = next((c for c in codigos + ids if PISTA_SUJETO.search(c)), None)
    if SUJETO:
        avisos.append(f"Campo de sujeto elegido automáticamente: {SUJETO}. Si no es el correcto: --sujeto COLUMNA.")

# Grupo y subcategoría (equivalen a CLASE y CATEGORIA)
GRUPO = categorias[0] if categorias else None
SUBCAT = categorias[1] if len(categorias) > 1 else None
if not ORDEN_CAT and GRUPO:
    avisos.append(f"Categorías principales tomadas por orden de columna: grupo = {GRUPO}"
                  + (f", subcategoría = {SUBCAT}" if SUBCAT else "") + ". Para cambiarlas: --categorias A,B.")

# Año de origen y fecha de alta (equivalen a AÑO_FABRICACION y FECHA_ALTA). Solo se emparejan solos si el nombre
# del año indica un origen: comparar una fecha de alta con, por ejemplo, un año de período daría falsos errores.
if not ANIO_ORIGEN:
    cand_a = [c for c in anios if PISTA_ORIGEN.search(c)]
    cand_f = [c for c in fechas if PISTA_ALTA.search(c)] or (fechas if len(fechas) == 1 else [])
    if len(cand_a) == 1 and len(cand_f) == 1:
        ANIO_ORIGEN, FECHA_ALTA = cand_a[0], cand_f[0]
        avisos.append(f"Se toma {ANIO_ORIGEN} como año de origen y {FECHA_ALTA} como fecha de alta (controles y "
                      f"gráficos de antigüedad). Si no corresponde: --anio COLUMNA --fecha COLUMNA.")
PAR = bool(ANIO_ORIGEN and FECHA_ALTA)

EXCLUIDOS = ([columna(x, "--excluir") for x in args.excluir.split(",") if x.strip()] if args.excluir is not None
             else [SUJETO] if SUJETO else [])

# Solo las medidas y los años se usan como número; categorías, códigos e identificadores quedan como texto
# (así "01" no se convierte en 1).
for c in medidas + anios:
    df[c] = numerico[c]
pd.options.display.float_format = "{:,.2f}".format
anio_ok = {c: df[c].where(df[c].between(1900, hoy.year + 1)) for c in anios}   # años imposibles = vacíos

campos = pd.DataFrame({
    "tipo": tipo, "rol": rol,
    "con_dato": con_dato.sum(), "%_con_dato": (con_dato.mean() * 100).round(2),
    "distintos": {c: crudo.loc[con_dato[c], c].nunique() for c in crudo.columns},
    "ejemplo": {c: (crudo.loc[con_dato[c], c].iloc[0][:30] if con_dato[c].any() else "") for c in crudo.columns}})
papel = {MONTO: "importe", SUJETO: "sujeto", GRUPO: "grupo", SUBCAT: "subcategoría",
         ANIO_ORIGEN: "año de origen", FECHA_ALTA: "fecha de alta"}
campos["uso"] = [papel.get(c, "") for c in campos.index]
log("\n--- CAMPOS DETECTADOS ---")
log(campos.to_string())
log("(rol: medida = cantidad/importe; categoría = pocos valores; código = muchos valores repetidos; "
    "identificador = casi único)")

# ---------------------------------------------------------------- 3. CALIDAD DE DATOS
no_convertible = pd.DataFrame({c: con_dato[c] & (numerico[c].isna() if tipo[c] == "número" else
                                                 df[c].isna() if tipo[c] == "fecha" else False)
                               for c in crudo.columns}, index=crudo.index)
sin_dato = ~con_dato | no_convertible
nulos = pd.DataFrame({"vacíos": crudo.isna().sum(), "rellenos": es_relleno.sum(),
                      "no_convertibles": no_convertible.sum(), "total_sin_dato": sin_dato.sum(),
                      "%": (sin_dato.mean() * 100).round(2)})
log("\n--- CALIDAD DE DATOS ---")
log("Valores vacíos, de relleno o no convertibles por campo:")
log(nulos.to_string())

# Completitud: para CADA campo; consistencia: según el rol. Se listan aunque den 0 casos, para dejar
# constancia de que el control se hizo.
controles = []      # (tipo, campo, control, máscara)
def control(tipo_c, campo, nombre, mascara):
    controles.append((tipo_c, campo, nombre, mascara.fillna(False).astype(bool)))

for c in crudo.columns:
    control("Completitud", c, "vacío", crudo[c].isna())
    control("Completitud", c, "texto de relleno (NULL, S/D, -, etc.)", es_relleno[c])
    if tipo[c] == "número" and rol[c] != "categoría":
        control("Completitud", c, "igual a 0", numerico[c] == 0)
    if tipo[c] in ("número", "fecha"):
        control("Completitud", c, f"no convertible a {tipo[c]}", no_convertible[c])
for c in crudo.columns:
    if rol[c] == "identificador" or c in DUPLICADOS:
        control("Consistencia", c, "duplicado", con_dato[c] & crudo[c].duplicated(keep=False))
    if tipo[c] == "texto":     # Excel no los admite y suelen venir de cargas o exportaciones defectuosas
        control("Consistencia", c, "caracteres de control (no imprimibles)",
                crudo[c].str.contains(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffd]", regex=True, na=False))
    if rol[c] == "medida":
        control("Consistencia", c, "negativo", df[c] < 0)
    if rol[c] == "año":
        control("Consistencia", c, "año < 1900 o futuro", (df[c] < 1900) | (df[c] > hoy.year + 1))
    if rol[c] == "fecha":
        control("Consistencia", c, "fecha futura", df[c] > hoy)
        control("Consistencia", c, "fecha anterior a 1900", df[c] < pd.Timestamp(1900, 1, 1))
if PAR:     # un 0 km fabricado a fin de año puede darse de alta al año siguiente: se tolera 1 año
    control("Consistencia", FECHA_ALTA, f"alta anterior al año de {ANIO_ORIGEN}",
            df[FECHA_ALTA].dt.year < anio_ok[ANIO_ORIGEN] - 1)
control("Consistencia", "(registro)", "fila completa duplicada", crudo.duplicated(keep=False))

tabla_ctrl = pd.DataFrame([{"tipo": t, "campo": c, "control": n, "casos": int(mk.sum()),
                            "%": round(mk.mean() * 100, 2)} for t, c, n, mk in controles])
log("\nControles de completitud y consistencia:")
log(tabla_ctrl.to_string(index=False))

# ---------------------------------------------------------------- 4. CASOS CON PROBLEMAS (entrada para muestreo)
# Una fila por registro con TODOS sus problemas en la columna PROBLEMA (separados por "; "),
# con el nombre completo del control (ej. "CATEGORIA: vacío") para saber de qué campo es cada problema.
def exportar_casos(nombre_archivo, campos_excluidos=()):
    problema = pd.Series("", index=df.index)
    for _, c, n, mk in controles:
        if c not in campos_excluidos:
            problema.loc[mk] = problema.loc[mk] + f"{c}: {n}; "
    con_problema = problema != ""
    resultado = crudo[con_problema].copy()
    resultado.insert(0, "PROBLEMA", problema[con_problema].str.rstrip("; "))
    ruta = ruta_salida(nombre_archivo)
    resultado.to_csv(ruta, sep="|", index=False, encoding="utf-8")
    sin = f" (sin contar problemas de {', '.join(campos_excluidos)})" if campos_excluidos else ""
    log(f"Registros con al menos un problema{sin}: {len(resultado):,} de {N:,} "
        f"({len(resultado) / N * 100:.1f}%) -> {ruta}")
    return resultado

log("")
casos = exportar_casos("casos_con_problemas.txt")
if EXCLUIDOS:   # segundo listado sin los problemas de esos campos
    exportar_casos("casos_con_problemas_sin_" + "_".join(EXCLUIDOS) + ".txt", EXCLUIDOS)

# ---------------------------------------------------------------- 5. ESTADÍSTICA
tablas = {}
cuantitativos = medidas + anios
if cuantitativos:
    est = df[cuantitativos].describe(percentiles=[.01, .05, .25, .5, .75, .9, .95, .99]).T
    est["suma"] = df[cuantitativos].sum()
    est["coef_variacion"] = est["std"] / est["mean"]
    est["asimetria"] = df[cuantitativos].skew()
    tablas["ESTADISTICA_NUMERICOS"] = est

m = acum = benford = estr = atipicos = None
if MONTO:
    m = df[MONTO].dropna()
    total = m.sum()
    est_monto = m.describe(percentiles=[.01, .05, .25, .5, .75, .9, .95, .99])
    tablas["ESTADISTICA_MONTO"] = est_monto.to_frame(MONTO)
    log(f"\n--- ESTADÍSTICA DE {MONTO} ---")
    log(est_monto.to_string())
    log(f"Total: {total:,.2f}")
    log(f"Coef. de variación: {m.std() / m.mean():.2f}")
    log(f"Asimetría (skew): {m.skew():.2f}   (>1 = muy asimétrica: pocos importes grandes, conviene estratificar)")
    q1, q3 = m.quantile([.25, .75])
    lim_sup = q3 + 3 * (q3 - q1)
    atipicos = df[df[MONTO] > lim_sup]
    log(f"Atípicos extremos (> Q3 + 3*IQR = {lim_sup:,.2f}): {len(atipicos):,} registros, "
        f"{atipicos[MONTO].sum() / total * 100:.1f}% del importe")

    pos = m[m > 0]
    if len(pos):
        # Concentración: qué % de registros explica el 80% del importe
        orden = pos.sort_values(ascending=False).reset_index(drop=True)
        acum = orden.cumsum() / orden.sum()
        n80 = int((acum <= 0.8).sum()) + 1
        log(f"El 80% del importe lo explican {n80:,} registros ({n80 / len(orden) * 100:.1f}% de los positivos)")

# ---------------------------------------------------------------- 6. CATEGORÍAS
if categorias:
    log("\n--- DISTRIBUCIÓN POR CATEGORÍAS ---")
for c in categorias:
    g = df.groupby(c, dropna=False)
    t = g.size().to_frame("cantidad")                  # size: cuenta todos los registros
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

if GRUPO and SUBCAT:
    tablas[f"{GRUPO}_x_{SUBCAT}"] = pd.crosstab(df[GRUPO].fillna("(vacío)"), df[SUBCAT].fillna("(vacío)"),
                                                margins=True, margins_name="TOTAL")
    log(f"\nCruce {GRUPO} x {SUBCAT}: en el Excel, hoja {GRUPO}_x_{SUBCAT}")

for c in codigos:       # códigos (descripciones, dominios, etc.): valores más frecuentes
    if c == SUJETO:
        continue
    g = df.groupby(c)
    t = g.size().to_frame("registros")
    if MONTO:
        t["total"] = g[MONTO].sum()
    t = t.sort_values("registros", ascending=False)
    tablas[f"TOP_{c}"] = t.head(30)
    log(f"\n{c}: {len(t):,} valores distintos; registros por valor: promedio {t['registros'].mean():.2f}, "
        f"máximo {t['registros'].max():,} (top 30 en el Excel)")

# ---------------------------------------------------------------- 7. SUJETOS
suj = None
if SUJETO:
    log(f"\n--- SUJETOS ({SUJETO}) ---")
    g = df.groupby(SUJETO)
    suj = g.size().to_frame("registros")
    if MONTO:
        suj["total"] = g[MONTO].sum()
    suj = suj.sort_values("total" if MONTO else "registros", ascending=False)
    log(f"Sujetos distintos: {len(suj):,}")
    log(f"Registros por sujeto: promedio {suj['registros'].mean():.2f}, máximo {suj['registros'].max():,}")
    if MONTO and suj["total"].sum():
        log(f"Top 10 sujetos concentran el {suj['total'].head(10).sum() / suj['total'].sum() * 100:.1f}% del importe")
    tablas["TOP_SUJETOS"] = suj.head(50)

# ---------------------------------------------------------------- 8. FECHAS
if fechas or anios:
    log("\n--- FECHAS ---")
for c in anios:
    if anio_ok[c].notna().any():
        log(f"{c} (válidos): {anio_ok[c].min():.0f} a {anio_ok[c].max():.0f}")
for c in fechas:
    if df[c].notna().any():
        log(f"{c}: {df[c].min():%d/%m/%Y} a {df[c].max():%d/%m/%Y}")
        tablas[f"POR_AÑO_{c}"] = df[c].dt.year.value_counts().sort_index().rename_axis("año").to_frame("cantidad")

# ---------------------------------------------------------------- 9. BENFORD (primer dígito)
if acum is not None:
    d1 = np.floor(pos / 10 ** np.floor(np.log10(pos))).astype(int).clip(1, 9)
    benford = pd.DataFrame({"observado_%": d1.value_counts(normalize=True).reindex(range(1, 10), fill_value=0) * 100})
    benford.index.name = "dígito"
    benford["esperado_%"] = [100 * math.log10(1 + 1 / d) for d in benford.index]
    benford["diferencia"] = benford["observado_%"] - benford["esperado_%"]
    tablas["BENFORD"] = benford.round(2)
    ordenes = math.log10(pos.max() / pos.min())
    log(f"\n--- BENFORD (primer dígito de {MONTO}) --- rango de {ordenes:.1f} órdenes de magnitud"
        + (" -> poco confiable (se necesitan al menos 2-3)" if ordenes < 2 else ""))
    log(benford.round(2).to_string())

# ---------------------------------------------------------------- 10. MARCO Y ESTRATOS PARA EL MUESTREO
if MONTO:
    log("\n--- MARCO DE MUESTREO ---")
    marco = pd.DataFrame({"registros": [N, m.gt(0).sum(), m.eq(0).sum(), m.lt(0).sum(), N - len(m)],
                          "importe": [total, m[m > 0].sum(), 0, m[m < 0].sum(), None]},
                         index=["Población total", "Importe > 0 (muestreable por importe)", "Importe = 0",
                                "Importe negativo", "Importe vacío o inválido"])
    marco["%_registros"] = (marco["registros"] / N * 100).round(2)
    tablas["MARCO_MUESTREO"] = marco
    log(marco.to_string())
    log("Los ceros, negativos y vacíos nunca salen en un muestreo por importe: tratarlos como población aparte.")

    if acum is not None and len(pos) >= 20:
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

if avisos:
    log("\n--- AVISOS ---")
    for a in avisos:
        log("* " + a)

# ---------------------------------------------------------------- 11. GRÁFICOS (.png)
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except ImportError:
    plt = None
    log("\n(matplotlib no está instalado: no se generan los gráficos .png. pip install matplotlib)")

def guardar(nombre):
    plt.tight_layout()
    plt.savefig(os.path.join(SALIDA, re.sub(r"[^\w.-]", "_", nombre)), dpi=120)
    plt.close("all")

if plt:
    if acum is not None:
        plt.figure(figsize=(8, 5))
        plt.hist(pos, bins=60)
        plt.title(f"Distribución de {MONTO}"); plt.xlabel(MONTO); plt.ylabel("Registros")
        guardar("01_histograma_importe.png")

        plt.figure(figsize=(8, 5))
        plt.hist(np.log10(pos), bins=60)
        plt.title(f"Distribución de {MONTO} (escala log10)"); plt.xlabel(f"log10({MONTO})"); plt.ylabel("Registros")
        guardar("02_histograma_importe_log.png")

    for c in categorias:
        t = tablas[f"CAT_{c}"].head(20)
        t.index = t.index.map(lambda x: "(vacío)" if pd.isna(x) else str(x)[:30])
        fig, ax = plt.subplots(1, 2 if MONTO else 1, figsize=(12 if MONTO else 7, 5), squeeze=False)
        t["cantidad"].plot.barh(ax=ax[0][0], title=f"Cantidad por {c}")
        ax[0][0].invert_yaxis()
        if MONTO:
            t["total"].plot.barh(ax=ax[0][1], title=f"{MONTO} por {c}", color="orange")
            ax[0][1].invert_yaxis()
        guardar(f"03_barras_{c.lower()}.png")

    if GRUPO and acum is not None:
        top_grupo = tablas[f"CAT_{GRUPO}"].dropna(subset=["total"]).head(10).index
        datos = df[df[GRUPO].isin(top_grupo) & (df[MONTO] > 0)]
        if len(datos):
            datos.boxplot(column=MONTO, by=GRUPO, rot=45, figsize=(10, 5))
            plt.yscale("log"); plt.title(f"{MONTO} por {GRUPO} (log)"); plt.suptitle("")
            guardar(f"04_boxplot_{GRUPO.lower()}.png")

    for c in anios:
        cuenta = anio_ok[c].dropna().astype(int).value_counts().sort_index()
        if len(cuenta):
            cuenta.plot(figsize=(9, 4), title=f"Registros por {c}")
            guardar(f"05_{c.lower()}.png")
    for c in fechas:
        cuenta = df[c].dt.year.dropna().astype(int).value_counts().sort_index()
        if len(cuenta):
            cuenta.plot.bar(figsize=(max(7, min(18, len(cuenta) * 0.3)), 4), title=f"Registros por año de {c}")
            guardar(f"06_por_anio_{c.lower()}.png")

    if acum is not None:
        plt.figure(figsize=(7, 5))
        plt.plot(np.arange(1, len(acum) + 1) / len(acum) * 100, acum * 100)
        plt.axhline(80, ls="--", c="gray")
        plt.title("Curva de concentración (Pareto)"); plt.xlabel("% de registros"); plt.ylabel("% del importe")
        guardar("07_pareto.png")

        benford[["observado_%", "esperado_%"]].plot.bar(figsize=(7, 4), title="Ley de Benford - primer dígito")
        guardar("08_benford.png")

        for c in anios:
            plt.figure(figsize=(7, 5))
            plt.scatter(anio_ok[c], df[MONTO].where(df[MONTO] > 0), s=3, alpha=.3)
            plt.yscale("log"); plt.title(f"{MONTO} vs {c}")
            guardar(f"09_importe_vs_{c.lower()}.png")

# ---------------------------------------------------------------- 12. ANÁLISIS POR CAMPO (una pestaña por campo)
def patron(serie):
    """Convierte cada valor en su formato: AB123CD -> AA999AA (A = letra, 9 = número)."""
    return (serie.str.upper().str.replace(r"[A-ZÑÁÉÍÓÚÜ]", "A", regex=True)
                 .str.replace(r"\d", "9", regex=True))

def frecuencias(serie, top):
    """Los 'top' valores más frecuentes; el resto se agrupa en una fila."""
    vc = serie.fillna("(vacío)").value_counts()
    t = vc.head(top).rename_axis("valor").reset_index(name="cantidad")
    if len(vc) > top:
        t.loc[len(t)] = [f"(otros {len(vc) - top:,} valores)".replace(",", "."), vc.iloc[top:].sum()]
    t["%"] = (t["cantidad"] / len(serie) * 100).round(2)
    return t

def analisis_campo(campo):
    txt = crudo[campo]                                   # valor original, tal como vino en el archivo
    validos = txt[con_dato[campo]]
    largo = validos.str.len()
    res = [("Tipo / rol detectado", f"{tipo[campo]} / {rol[campo]}"),
           ("Registros totales", len(txt)),
           ("Con dato (sin vacíos ni rellenos)", len(validos)),
           ("% con dato", round(len(validos) / len(txt) * 100, 2)),
           ("Valores distintos", validos.nunique()),
           ("Largo mínimo (caracteres)", largo.min() if len(largo) else None),
           ("Largo máximo (caracteres)", largo.max() if len(largo) else None)]
    if rol[campo] in ("medida", "año"):
        v = df[campo].dropna()
        if len(v):
            res += [("Mínimo", v.min()), ("Máximo", v.max()), ("Promedio", round(v.mean(), 2)), ("Mediana", v.median())]
    if rol[campo] == "fecha":
        v = df[campo].dropna()
        if len(v):
            res += [("Fecha mínima", f"{v.min():%d/%m/%Y}"), ("Fecha máxima", f"{v.max():%d/%m/%Y}")]
    resumen_c = pd.DataFrame(res, columns=["indicador", "valor"])

    propios = [(n, mk) for _, c, n, mk in controles if c == campo]
    controles_c = pd.DataFrame([(n, int(mk.sum()), round(mk.mean() * 100, 2)) for n, mk in propios],
                               columns=["control", "casos", "%"])
    # Lista COMPLETA de registros con algún problema en este campo, con el motivo
    problema = pd.Series("", index=df.index)
    for n, mk in propios:
        problema.loc[mk] = problema.loc[mk] + n + "; "
    registros = crudo[problema != ""].copy()
    registros.insert(0, "PROBLEMA", problema[problema != ""].str.rstrip("; "))
    registros = registros.sort_values(["PROBLEMA", campo], na_position="first")   # agrupa duplicados
    return resumen_c, controles_c, frecuencias(txt, 20), frecuencias(patron(txt), 15), registros

usadas = set()
def hoja(nombre):
    """Nombre de hoja válido para Excel (máximo 31 caracteres, sin repetir)."""
    base = re.sub(r"[\[\]:*?/\\]", "_", nombre)[:31]
    final, i = base, 2
    while final.upper() in usadas:
        final = f"{base[:28]}_{i}"; i += 1
    usadas.add(final.upper())
    return final

def escribir_hoja_campo(xl, campo):
    from openpyxl.styles import Font
    nombre = hoja("CAMPO_" + campo)
    resumen_c, controles_c, frec, patr, registros = analisis_campo(campo)
    bloques = [("RESUMEN", resumen_c, 0), ("CONTROLES", controles_c, 3),
               ("VALORES MÁS FRECUENTES", frec, 7), ("FORMATOS (A = letra, 9 = número)", patr, 11)]
    for _, t, col in bloques:
        excel(t).to_excel(xl, sheet_name=nombre, startrow=2, startcol=col, index=False)
    ws = xl.sheets[nombre]
    ws.cell(1, 1, f"Análisis del campo {campo}").font = Font(bold=True, size=14)
    for titulo, _, col in bloques:
        ws.cell(2, col + 1, titulo).font = Font(bold=True)

    fila = max(len(t) for _, t, _ in bloques) + 6
    n = len(registros)
    titulo = f"REGISTROS CON PROBLEMAS EN ESTE CAMPO: {n:,}".replace(",", ".")
    if n == 0:
        titulo += " (ninguno)"
    elif n > EXCEL_MAX:      # no entran en Excel: se guarda el listado completo aparte
        txt = ruta_salida(f"problemas_{campo}.txt")
        registros.to_csv(txt, sep="|", index=False, encoding="utf-8-sig")
        titulo += (f" (se muestran los primeros {EXCEL_MAX:,}; el listado completo está en "
                   f"{os.path.basename(txt)})").replace(",", ".")
        registros = registros.head(EXCEL_MAX)
    ws.cell(fila, 1, titulo).font = Font(bold=True, color="B42318")
    if n:
        excel(registros).to_excel(xl, sheet_name=nombre, startrow=fila, index=False)
    for col, ancho in zip("ABCDEFGHIJKLMN", [36, 16, 3, 40, 10, 8, 3, 28, 10, 8, 3, 22, 10, 8]):
        ws.column_dimensions[col].width = ancho

# ---------------------------------------------------------------- 13. EXPORTAR
log("\nEscribiendo el Excel (con muchos registros con problemas puede tardar varios minutos)...")
with pd.ExcelWriter(ruta_salida("tablas_analisis.xlsx"), engine="openpyxl") as xl:
    excel(campos).to_excel(xl, sheet_name=hoja("CAMPOS"))
    excel(nulos).to_excel(xl, sheet_name=hoja("NULOS"))
    excel(tabla_ctrl).to_excel(xl, sheet_name=hoja("CONTROLES"), index=False)
    for c in crudo.columns:
        escribir_hoja_campo(xl, c)
    for nombre, t in tablas.items():
        excel(t).to_excel(xl, sheet_name=hoja(nombre))
    if atipicos is not None:
        excel(crudo.loc[atipicos.sort_values(MONTO, ascending=False).index].head(EXCEL_MAX)).to_excel(
            xl, sheet_name=hoja("ATIPICOS"), index=False)

with open(ruta_salida("resumen.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(resumen))
print(f"\nListo. Resultados en la carpeta '{SALIDA}'")
