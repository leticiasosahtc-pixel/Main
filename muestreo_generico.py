"""
Muestreo aleatorio para auditoría sobre una lista de registros (por ejemplo, los casos con problemas que
deja analisis_generico.py). Hace lo mismo que muestreo.py, pero sin nombres de columnas fijos: detecta el
separador y la columna que identifica a cada registro.

Uso:
    python muestreo_generico.py                                       (usa el listado de casos de salida_analisis)
    python muestreo_generico.py archivo.txt                           (texto delimitado)
    python muestreo_generico.py salida_analisis/tablas_analisis.xlsx CAMPO_CATEGORIA   (una pestaña del análisis)
Opciones (todas opcionales; si no se indican valen los PARÁMETROS de abajo):
    --clave COLUMNA        junta en un solo registro todas las filas con el mismo valor de esa columna
                           (sin --clave solo se juntan filas idénticas, salvo por PROBLEMA)
    --completo duplicado   los registros cuyo PROBLEMA contiene ese texto entran TODOS en la muestra
                           (se pueden poner varios separados por coma)
    --estrato COLUMNA      columna para estratificar; "ninguno" = aleatorio simple sobre toda la base
    --confianza 0.95  --error 0.09  --tasa 0.5  --tamanio 60  --minimo 2  --reserva 0.10
    --semilla 123456       ("azar" = semilla al azar, que queda registrada en el Excel)
    --por-estrato          calcula una muestra suficiente para concluir sobre CADA estrato
    --sep "|"              separador (si se omite, se detecta entre | ; tabulación y coma)
Requiere: pip install pandas openpyxl
Salida: muestra_AAAAMMDD_HHMMSS.xlsx con la muestra, la reserva y los parámetros usados.
"""
import argparse
import glob
import hashlib
import math
import os
import re
import secrets
import sys
from statistics import NormalDist
import numpy as np
import pandas as pd

# ---------------------------------------------------------------- PARÁMETROS (valores por defecto)
CONFIANZA = 0.95          # nivel de confianza (0.90, 0.95, 0.99)
ERROR_TOLERABLE = 0.09    # margen de error aceptado (0.06 = ±6 puntos porcentuales)
TASA_ESPERADA = 0.50      # % de errores que se espera encontrar; 0.50 es lo más conservador
TAMANIO_FIJO = None       # si querés un tamaño fijo (ej. 60), ponelo acá y se ignora la fórmula
ESTRATO = "PROBLEMA"      # columna para estratificar; None = aleatorio simple sobre toda la base
CONCLUIR_POR_ESTRATO = False   # True = calcula una muestra suficiente para concluir sobre CADA estrato
MINIMO_POR_ESTRATO = 2    # ningún estrato queda con menos casos (si tiene menos, se toman todos)
RESERVA = 0.10            # casos extra (10%) para reemplazar los que no se puedan verificar
SEMILLA = 123456          # semilla fija: con el mismo archivo y parámetros sale siempre la misma muestra
CLAVE = None              # columna para juntar filas del mismo registro; None = solo se juntan filas idénticas
COMPLETO = []             # textos de PROBLEMA cuyos registros se revisan al 100% (ej. ["duplicado"])
COL_PROBLEMA = "PROBLEMA" # columna con la descripción del problema
ESTRATO_VARIOS = "Más de un problema"   # estrato para los registros con 2 o más problemas
PISTA_ID = re.compile(r"^(ID|COD|NRO|NUM|CLAVE)|_ID$|CODIGO|NUMERO|DOMINIO|PATENTE|PARTIDA|CUENTA|INSCRIP|CUIT")
CARPETA_ANALISIS = "salida_analisis"

ap = argparse.ArgumentParser()
ap.add_argument("archivo", nargs="?")
ap.add_argument("hoja", nargs="?")
ap.add_argument("--clave")
ap.add_argument("--completo")
ap.add_argument("--estrato")
ap.add_argument("--confianza", type=float)
ap.add_argument("--error", type=float)
ap.add_argument("--tasa", type=float)
ap.add_argument("--tamanio", type=int)
ap.add_argument("--minimo", type=int)
ap.add_argument("--reserva", type=float)
ap.add_argument("--semilla")
ap.add_argument("--por-estrato", action="store_true")
ap.add_argument("--sep")
args = ap.parse_args()

CONFIANZA = args.confianza or CONFIANZA
ERROR_TOLERABLE = args.error or ERROR_TOLERABLE
TASA_ESPERADA = args.tasa or TASA_ESPERADA
TAMANIO_FIJO = args.tamanio or TAMANIO_FIJO
MINIMO_POR_ESTRATO = MINIMO_POR_ESTRATO if args.minimo is None else args.minimo
RESERVA = RESERVA if args.reserva is None else args.reserva
CONCLUIR_POR_ESTRATO = args.por_estrato or CONCLUIR_POR_ESTRATO
CLAVE = args.clave or CLAVE
if args.completo is not None:
    COMPLETO = [t.strip() for t in args.completo.split(",") if t.strip()]
if args.estrato is not None:
    ESTRATO = None if args.estrato.strip().lower() in ("", "ninguno", "none") else args.estrato
if args.semilla is not None:
    SEMILLA = None if args.semilla.strip().lower() == "azar" else int(args.semilla)
for nombre, valor in (("--confianza", CONFIANZA), ("--error", ERROR_TOLERABLE), ("--tasa", TASA_ESPERADA)):
    if not 0 < valor < 1:
        sys.exit(f"{nombre} tiene que estar entre 0 y 1 (ej. 0.95), no {valor}.")

# Sin archivo: el listado de casos que deja analisis_generico.py (sin los problemas del sujeto, si existe)
ARCHIVO = args.archivo
if not ARCHIVO:
    candidatos = (sorted(glob.glob(os.path.join(CARPETA_ANALISIS, "casos_con_problemas_sin_*.txt")))
                  or glob.glob(os.path.join(CARPETA_ANALISIS, "casos_con_problemas.txt")))
    if not candidatos:
        sys.exit("Indicá el archivo a muestrear, por ejemplo:\n   python muestreo_generico.py casos.txt")
    ARCHIVO = candidatos[0]
    print(f"Archivo de población: {ARCHIVO}")
if not os.path.exists(ARCHIVO):
    sys.exit(f"No existe el archivo {ARCHIVO}")
HOJA = args.hoja
SALIDA = f"muestra_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
avisos = []

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

# ---------------------------------------------------------------- 1. LECTURA
def leer(archivo, hoja):
    """Devuelve la tabla y el número de fila del archivo donde empieza el primer registro."""
    if archivo.lower().endswith((".xlsx", ".xlsm", ".xls")):
        hojas = pd.ExcelFile(archivo).sheet_names
        if hoja is None and len(hojas) > 1:
            sys.exit("El Excel tiene varias pestañas: indicá cuál usar, por ejemplo\n"
                     f"   python muestreo_generico.py {archivo} "
                     f"{next((h for h in hojas if h.startswith('CAMPO_')), hojas[-1])}\nPestañas: {', '.join(hojas)}")
        if hoja is not None and hoja not in hojas:
            sys.exit(f"No existe la pestaña '{hoja}'. Pestañas: {', '.join(hojas)}")
        raw = pd.read_excel(archivo, sheet_name=hoja or 0, header=None, dtype=str)
        # En las pestañas CAMPO_ del análisis, la lista empieza debajo de "REGISTROS CON PROBLEMAS..."
        marca = raw[0].astype(str).str.startswith("REGISTROS CON PROBLEMAS")
        fila = int(marca.idxmax()) + 1 if marca.any() else 0
        if fila >= len(raw):                  # pestaña CAMPO_ sin registros con problemas
            return pd.DataFrame(), fila + 2, None
        tabla = raw.iloc[fila + 1:].copy()
        tabla.columns = raw.iloc[fila].astype(str)
        tabla = tabla.dropna(axis=1, how="all")
        return tabla, fila + 2, None
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(archivo, encoding=enc) as f:
                cabecera = f.readline()
            sep = args.sep or max("|;\t,", key=cabecera.count)
            return pd.read_csv(archivo, sep=sep, dtype=str, keep_default_na=False, encoding=enc), 2, sep
        except UnicodeDecodeError:
            continue

pob, primera_fila, SEP = leer(ARCHIVO, HOJA)
if pob.empty:
    sys.exit("El archivo o la pestaña no tiene registros para muestrear.")
pob.columns = pob.columns.astype(str).str.strip().str.upper()
pob = pob.apply(lambda c: c.str.strip() if c.dtype == object else c)
pob = pob.reset_index(drop=True)
pob.insert(0, "FILA_EN_ARCHIVO", (pob.index + primera_fila).astype(str))   # para ubicar cada caso en el archivo
FILAS = len(pob)
COL_PROBLEMA = COL_PROBLEMA.upper()
tiene_problema = COL_PROBLEMA in pob.columns

# Huella digital del archivo: prueba de que la población no cambió después del muestreo
sha = hashlib.sha256()
with open(ARCHIVO, "rb") as f:
    for bloque in iter(lambda: f.read(1 << 20), b""):
        sha.update(bloque)
huella = sha.hexdigest()

# ---------------------------------------------------------------- 1 bis. CLAVE: QUÉ IDENTIFICA A UN REGISTRO
datos = [c for c in pob.columns if c not in ("FILA_EN_ARCHIVO", COL_PROBLEMA)]

distintas = pob.drop_duplicates(datos)     # un registro repetido en varias filas (una por problema) cuenta una vez
def unicidad(c):
    v = distintas[c].replace("", pd.NA).dropna()
    return v.nunique() / len(v) if len(v) else 0

if CLAVE:
    CLAVE = CLAVE.strip().upper()
    if CLAVE not in pob.columns:
        sys.exit(f"--clave {CLAVE}: no existe. Columnas: {', '.join(datos)}")
    clave, clave_txt = [CLAVE], CLAVE
    vacias = int(pob[CLAVE].replace("", pd.NA).isna().sum())
    if vacias:
        avisos.append(f"{vacias:,} filas tienen la clave {CLAVE} vacía: entre ellas se consideran el mismo registro.")
else:
    # Sin --clave solo se juntan filas idénticas (un mismo registro listado una vez por problema). Dos registros
    # con el mismo identificador pero datos distintos son registros distintos: justamente los duplicados que
    # detecta el análisis, y no se pueden juntar ni perder.
    clave = datos
    clave_txt = "todos los datos del registro (solo se juntan filas idénticas, salvo por PROBLEMA)"
    # Control para no muestrear por error una tabla de resumen: tiene que haber un identificador o la columna PROBLEMA
    ids = [c for c in datos if PISTA_ID.search(c) and unicidad(c) >= 0.9]
    if not ids and not tiene_problema:
        sys.exit(f"No encuentro una columna que identifique a cada registro ni la columna {COL_PROBLEMA}: puede que el "
                 f"archivo no sea una lista de registros.\nColumnas encontradas: {', '.join(datos)}\n"
                 "Si es una lista de registros, indicá cuál los identifica con --clave COLUMNA.")
    for c in ids:
        valores = distintas[c].replace("", pd.NA).dropna()
        n = int(valores.duplicated(keep=False).sum())
        if n:
            avisos.append(f"{n:,} registros comparten {c} con otro registro de datos distintos: se muestrean como "
                          f"registros separados. Si en realidad son el mismo registro: --clave {c}.")

# ---------------------------------------------------------------- 1 ter. UN REGISTRO = UNA FILA
# Si un registro aparece en varias filas (una por problema), se junta en una sola fila con todos
# sus problemas. Así la unidad de muestreo es el registro y nunca puede salir dos veces.
def unir_problemas(valores):
    """'A; B' + 'B; C' -> 'A; B; C' (sin repetir y en orden alfabético)."""
    partes = {p.strip() for v in valores if pd.notna(v) for p in str(v).split(";") if p.strip()}
    return "; ".join(sorted(partes))

repetidos = pob.duplicated(clave, keep=False)
datos_distintos = 0
if repetidos.any():
    rep = pob[repetidos]
    otras = [c for c in datos if c not in clave]
    if otras:   # el mismo registro con datos distintos en sus filas: se toma la primera y se avisa
        datos_distintos = int((rep.groupby(clave, dropna=False)[otras].nunique(dropna=False) > 1).any(axis=1).sum())
    reglas = {c: "first" for c in pob.columns if c not in clave}
    reglas["FILA_EN_ARCHIVO"] = lambda s: "; ".join(s)
    if tiene_problema:
        reglas[COL_PROBLEMA] = unir_problemas
    juntos = rep.groupby(clave, dropna=False, sort=False).agg(reglas).reset_index()
    pob = pd.concat([pob[~repetidos], juntos[pob.columns]], ignore_index=True)
    avisos.append(f"Se juntaron {repetidos.sum():,} filas repetidas en {len(juntos):,} registros únicos.")
    if datos_distintos:
        avisos.append(f"{datos_distintos:,} valores de {clave_txt} tienen datos distintos en sus filas: se juntaron y "
                      "se tomó la primera fila. Si son registros distintos (una clave duplicada), sacá --clave.")
N = len(pob)

if tiene_problema:
    pob[COL_PROBLEMA] = pob[COL_PROBLEMA].map(lambda v: unir_problemas([v]) or "(vacío)")
    pob.insert(pob.columns.get_loc(COL_PROBLEMA) + 1, "CANT_PROBLEMAS",
               pob[COL_PROBLEMA].str.count(";") + 1)
varios = int((pob["CANT_PROBLEMAS"] > 1).sum()) if tiene_problema else 0

# Estrato de cada registro: su problema, o "Más de un problema" si tiene varios
estrato = None
if ESTRATO and ESTRATO.upper() in pob.columns:
    estrato = "ESTRATO"
    if ESTRATO.upper() == COL_PROBLEMA:
        valores = pob[COL_PROBLEMA].where(pob["CANT_PROBLEMAS"] == 1, ESTRATO_VARIOS)
    else:
        valores = pob[ESTRATO.upper()].replace("", "(vacío)").fillna("(vacío)")
    pob.insert(1, estrato, valores)
elif ESTRATO:
    if args.estrato:
        sys.exit(f"--estrato {ESTRATO}: no existe. Columnas: {', '.join(datos)}")
    avisos.append(f"No existe la columna {ESTRATO}; se hace muestreo aleatorio simple.")
ESTRATO_COL = ESTRATO.upper() if estrato else None      # columna real de estratificación (None = ninguna)

# Revisión completa: los registros cuyo PROBLEMA contiene alguno de los textos de COMPLETO entran todos
censo = []
if COMPLETO:
    if not tiene_problema:
        sys.exit(f"--completo: el archivo no tiene la columna {COL_PROBLEMA}.")
    etiqueta = pd.Series(pd.NA, index=pob.index, dtype="object")
    for texto in COMPLETO:
        cumple = etiqueta.isna() & pob[COL_PROBLEMA].str.contains(texto, case=False, regex=False)
        etiqueta[cumple] = f"Revisión completa: {texto}"
    if etiqueta.notna().any():
        if not estrato:
            estrato = "ESTRATO"
            pob.insert(1, estrato, "(toda la base)")
        pob.loc[etiqueta.notna(), estrato] = etiqueta[etiqueta.notna()]
        censo = sorted(etiqueta.dropna().unique())
    sin = [t for t in COMPLETO if f"Revisión completa: {t}" not in censo]
    if sin:
        avisos.append(f"--completo: ningún registro tiene en PROBLEMA el texto {', '.join(sin)}.")

# ---------------------------------------------------------------- 2. TAMAÑO DE MUESTRA
z = NormalDist().inv_cdf((1 + CONFIANZA) / 2)

def tamanio(n_pob):
    """Muestreo de atributos con corrección por población finita."""
    if n_pob == 0:
        return 0
    if TAMANIO_FIJO:
        return min(TAMANIO_FIJO, n_pob)
    n0 = z ** 2 * TASA_ESPERADA * (1 - TASA_ESPERADA) / ERROR_TOLERABLE ** 2
    return min(n_pob, math.ceil(n0 / (1 + (n0 - 1) / n_pob)))

if estrato:
    tamanos = pob[estrato].value_counts().sort_index()
    resto = tamanos.drop(censo)                     # los de revisión completa no entran en el cálculo
    N_resto = int(resto.sum())
    n_objetivo = tamanio(N_resto)
    if CONCLUIR_POR_ESTRATO:
        asignado = {k: tamanio(v) for k, v in resto.items()}
    else:
        # Asignación proporcional al tamaño de cada estrato (método del mayor resto)
        exacto = {k: n_objetivo * v / N_resto for k, v in resto.items()}
        asignado = {k: min(v, max(MINIMO_POR_ESTRATO, math.floor(exacto[k]))) for k, v in resto.items()}
        falta = n_objetivo - sum(asignado.values())
        for k in sorted(exacto, key=lambda k: exacto[k] - math.floor(exacto[k]), reverse=True):
            if falta <= 0:
                break
            if asignado[k] < tamanos[k]:
                asignado[k] += 1
                falta -= 1
        if sum(asignado.values()) > n_objetivo:
            avisos.append(f"El mínimo de {MINIMO_POR_ESTRATO} casos por estrato lleva la muestra de {n_objetivo:,} "
                          f"a {sum(asignado.values()):,} casos ({len(resto):,} estratos).")
    asignado.update({k: int(tamanos[k]) for k in censo})
else:
    n_objetivo = tamanio(N)
    tamanos = pd.Series({"(toda la base)": N})
    asignado = {"(toda la base)": n_objetivo}

# ---------------------------------------------------------------- 3. SELECCIÓN ALEATORIA
if SEMILLA is None:
    SEMILLA = secrets.randbelow(10 ** 9)
azar = np.random.RandomState(SEMILLA)

muestra, reserva = [], []
grupos = pob.groupby(estrato, sort=True) if estrato else [("(toda la base)", pob)]
for k, grupo in grupos:
    mezclado = grupo.sample(frac=1, random_state=azar)      # orden aleatorio de todo el estrato
    n_h = asignado[k]
    r_h = min(len(grupo) - n_h, math.ceil(n_h * RESERVA))
    muestra.append(mezclado.iloc[:n_h])
    reserva.append(mezclado.iloc[n_h:n_h + r_h])            # los siguientes en el orden aleatorio

muestra = pd.concat(muestra)
reserva = pd.concat(reserva) if reserva else pob.iloc[0:0]
for t in (muestra, reserva):
    t.insert(0, "ORDEN", range(1, len(t) + 1))
    for col in ["RESULTADO", "OBSERVACIONES", "AUDITOR", "FECHA_REVISION"]:
        t[col] = ""

resumen = pd.DataFrame({
    "estrato": list(tamanos.index),
    "poblacion": list(tamanos.values),
    "muestra": [asignado[k] for k in tamanos.index],
    "reserva": [int((reserva[estrato] == k).sum()) if estrato else len(reserva) for k in tamanos.index],
})
resumen["% muestreado"] = (resumen["muestra"] / resumen["poblacion"] * 100).round(1)
resumen.loc[len(resumen)] = ["TOTAL", N, resumen["muestra"].sum(), resumen["reserva"].sum(),
                             round(resumen["muestra"].sum() / N * 100, 1)]

metodo = ("Estratificado por " + ESTRATO_COL
          + (f" (los registros con varios problemas van a '{ESTRATO_VARIOS}')" if ESTRATO_COL == COL_PROBLEMA else "")
          if ESTRATO_COL else "Aleatorio simple")
if censo:
    metodo += f" + revisión completa de los registros con: {', '.join(COMPLETO)}"
metodo += (" - tamaño suficiente para concluir sobre cada estrato" if estrato and CONCLUIR_POR_ESTRATO
           else " - asignación proporcional" if ESTRATO_COL else "")
parametros = pd.DataFrame([
    ("Archivo de población", os.path.abspath(ARCHIVO) + (f" (hoja {HOJA})" if HOJA else "")),
    ("Separador", repr(SEP) if SEP else "No aplica (Excel)"),
    ("Huella SHA-256 del archivo", huella),
    ("Fecha y hora del muestreo", f"{pd.Timestamp.now():%d/%m/%Y %H:%M:%S}"),
    ("Filas en el archivo", FILAS),
    ("Clave que identifica un registro", clave_txt),
    ("Registros únicos (población muestreada)", N),
    ("Registros con más de un problema", varios if tiene_problema else "No aplica (no hay columna PROBLEMA)"),
    ("Registros con datos distintos entre sus filas", datos_distintos),
    ("Método", metodo),
    ("Tamaño de muestra", f"Fijo ({TAMANIO_FIJO})" if TAMANIO_FIJO else f"Calculado (muestreo de atributos): {n_objetivo}"),
    ("Nivel de confianza", CONFIANZA),
    ("Error tolerable", ERROR_TOLERABLE),
    ("Tasa de error esperada", TASA_ESPERADA),
    ("Valor z", round(z, 4)),
    ("Revisión completa (100%)", f"{sum(int(tamanos[k]) for k in censo):,} registros con {', '.join(COMPLETO)}"
                                 if censo else "No"),
    ("Mínimo por estrato", MINIMO_POR_ESTRATO if estrato and not CONCLUIR_POR_ESTRATO else "No aplica"),
    ("Reserva", RESERVA),
    ("Semilla aleatoria", SEMILLA),
    ("Casos en la muestra", len(muestra)),
    ("Casos en la reserva", len(reserva)),
    ("Avisos", " | ".join(avisos) if avisos else "Ninguno"),
    ("Cómo repetir esta muestra",
     f"python muestreo_generico.py \"{ARCHIVO}\"" + (f" {HOJA}" if HOJA else "")
     + (f" --clave {CLAVE}" if CLAVE else "") + (f' --completo "{",".join(COMPLETO)}"' if COMPLETO else "") + f" --estrato {ESTRATO_COL or 'ninguno'}"
     + (f" --tamanio {TAMANIO_FIJO}" if TAMANIO_FIJO else f" --confianza {CONFIANZA} --error {ERROR_TOLERABLE} --tasa {TASA_ESPERADA}")
     + (" --por-estrato" if CONCLUIR_POR_ESTRATO else f" --minimo {MINIMO_POR_ESTRATO}")
     + f" --reserva {RESERVA} --semilla {SEMILLA}  (sobre el mismo archivo: comparar la huella)"),
], columns=["parámetro", "valor"])

# ---------------------------------------------------------------- 4. EXPORTAR
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

with pd.ExcelWriter(SALIDA, engine="openpyxl") as xl:
    excel(parametros).to_excel(xl, sheet_name="PARAMETROS", index=False)
    excel(resumen).to_excel(xl, sheet_name="RESUMEN_ESTRATOS", index=False)
    excel(muestra).to_excel(xl, sheet_name="MUESTRA", index=False)
    excel(reserva).to_excel(xl, sheet_name="RESERVA", index=False)
    for nombre, ws in xl.sheets.items():
        for celda in ws[1]:
            celda.font = Font(bold=True)
        ws.freeze_panes = "A2"
        for i, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), start=1):
            ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(max(ancho + 2, 10), 70)
    # Lista desplegable en RESULTADO para que el resultado se cargue siempre igual
    for nombre, tabla in (("MUESTRA", muestra), ("RESERVA", reserva)):
        if len(tabla):
            ws = xl.sheets[nombre]
            col = get_column_letter(list(tabla.columns).index("RESULTADO") + 1)
            dv = DataValidation(type="list", formula1='"CORRECTO,ERROR,NO SE PUDO VERIFICAR"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{col}2:{col}{len(tabla) + 1}")

for a in avisos:
    print("AVISO: " + a)
print(f"Filas en el archivo: {FILAS:,} | Registros únicos: {N:,} | Muestra: {len(muestra):,} | Reserva: {len(reserva):,}")
print(resumen.to_string(index=False))
print(f"Semilla: {SEMILLA}")
print(f"Archivo generado: {SALIDA}")
