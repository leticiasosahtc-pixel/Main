"""
Prepara un listado de movimientos de cuenta corriente para el análisis y el muestreo por cuenta.

  1. Descarta las líneas que no son registros (vacías, separadores "-----", TOTALES) y las guarda aparte.
  2. Separa las filas de saldo (SALDO INIC VTAS, TOTAL FACTURADO INICIAL, etc.) de los movimientos.
  3. Controla el saldo corrido: en cada cuenta, SALDO = SALDO anterior + DEBE - HABER, fila por fila.
  4. Compara los TOTALES del pie del archivo con la suma de los movimientos.
  5. Arma un resumen con una fila por cuenta.

Uso:   python movimientos_por_cuenta.py archivo.txt
Requiere: pip install pandas openpyxl
Salida: carpeta 'salida_cuentas' con movimientos.txt, saldos.txt, resumen_por_cuenta.txt,
        diferencias_saldo.txt, lineas_descartadas.txt y control_cuentas.xlsx

Después:
    python informe_generico.py salida_cuentas/resumen_por_cuenta.txt --monto TOTAL_DEBE --sujeto CUIT/DNI
    python muestreo_generico.py salida_cuentas/resumen_por_cuenta.txt --clave CUENTA --estrato ninguno
"""
import io
import os
import re
import sys
import pandas as pd

# ---------------------------------------------------------------- PARÁMETROS (modificar acá)
SEPARADOR = "|"
CUENTA, FECHA, DEBE, HABER, SALDO, TIPO = "CUENTA", "FEC_MOVIMIENTO", "IMP_DEBE", "IMP_HABER", "SALDO", "TIPO_MOVIM"
DATOS_CUENTA = ["TITULAR", "CUIT/DNI", "DOMICILIO", "GRUPO"]     # se toman del primer movimiento de cada cuenta
PERIODO = ["ANIO", "PERIODO"]
# Filas de saldo: TIPO_MOVIM empieza con alguno de estos textos (sin importar mayúsculas, espacios ni ":")
TIPOS_SALDO = ["SALDO INIC VTAS", "TOTAL FACTURADO INIC", "SALDO INI PAGOS", "TOTAL PAGOS INIC"]
TOLERANCIA = 0.01         # diferencia de redondeo aceptada en el control de saldo corrido
SALIDA = "salida_cuentas"

if len(sys.argv) < 2:
    sys.exit("Uso: python movimientos_por_cuenta.py archivo.txt")
ARCHIVO = sys.argv[1]
os.makedirs(SALIDA, exist_ok=True)

def normalizar(t):
    return re.sub(r"\s+", " ", str(t)).strip(" :").upper()

PATRON_SALDO = re.compile("^(" + "|".join(re.escape(normalizar(t)) for t in TIPOS_SALDO) + ")")

# ---------------------------------------------------------------- 1. LECTURA LÍNEA POR LÍNEA
for codificacion in ("utf-8-sig", "latin-1"):
    try:
        with open(ARCHIVO, encoding=codificacion, newline="") as f:
            lineas = f.read().split("\n")
        break
    except UnicodeDecodeError:
        continue

# La cabecera es la primera línea que nombra la cuenta y el debe
inicio = next((i for i, l in enumerate(lineas)
               if SEPARADOR in l and CUENTA in l.upper() and DEBE in l.upper().replace(" ", "")), None)
if inicio is None:
    sys.exit(f"No encuentro la línea de títulos (con {CUENTA} e {DEBE}).")
columnas = [c.strip().upper() for c in lineas[inicio].rstrip("\r").split(SEPARADOR)]
faltan = [c for c in (CUENTA, FECHA, DEBE, HABER, SALDO, TIPO) if c not in columnas]
if faltan:
    sys.exit(f"Faltan las columnas {', '.join(faltan)}. Columnas del archivo: {', '.join(columnas)}\n"
             "Si se llaman distinto, cambiá los nombres en PARÁMETROS.")

buenas, descartadas, totales_pie = [], [], None
for nro, linea in enumerate(lineas[inicio + 1:], start=inicio + 2):
    texto = linea.rstrip("\r")
    partes = texto.split(SEPARADOR)
    if not texto.strip():
        motivo = "línea vacía"
    elif set(texto.strip()) <= set("-=_ "):
        motivo = "línea separadora"
    elif partes[0].strip().upper().startswith("TOTAL") and len(partes) != len(columnas):
        motivo = "totales del pie"
        totales_pie = [p.strip() for p in partes]
    elif len(partes) != len(columnas):
        motivo = f"tiene {len(partes)} campos en vez de {len(columnas)}"
    else:
        buenas.append([str(nro)] + [p.strip() for p in partes])
        continue
    descartadas.append((nro, motivo, texto))

df = pd.DataFrame(buenas, columns=["NRO_LINEA"] + columnas)
if df.empty:
    sys.exit("No quedó ningún registro después de descartar vacías, separadores y totales.")

def a_numero(s):
    """1.234,56 / -1.234,56 / 1.234,56- -> número. Vacío = 0."""
    s = s.str.replace(r"\s", "", regex=True)
    s = s.str.replace(r"^(.*)-$", r"-\1", regex=True)                       # signo al final
    s = s.where(~s.str.contains(",", regex=False), s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s.replace("", "0"), errors="coerce")

for c in (DEBE, HABER, SALDO):
    df["_" + c] = a_numero(df[c])
df["_FECHA"] = pd.to_datetime(df[FECHA].replace("", pd.NA), errors="coerce", dayfirst=True)
no_numericos = {c: int(df["_" + c].isna().sum()) for c in (DEBE, HABER, SALDO)}
df["ES_SALDO"] = df[TIPO].map(normalizar).str.match(PATRON_SALDO)

# ---------------------------------------------------------------- 2. CONTROL DE SALDO CORRIDO
# En el orden del archivo, dentro de cada cuenta: SALDO = SALDO de la fila anterior + DEBE - HABER.
# Se controla en todas las filas (también las de saldo); la primera fila de cada cuenta no tiene anterior.
g = df.groupby(CUENTA, sort=False)
df["SALDO_ESPERADO"] = g["_" + SALDO].shift() + df["_" + DEBE] - df["_" + HABER]
df["DIF_SALDO"] = (df["_" + SALDO] - df["SALDO_ESPERADO"]).round(2)
df["SALDO_NO_CIERRA"] = df["DIF_SALDO"].abs() > TOLERANCIA
# Cuentas partidas: si una cuenta aparece en dos bloques separados del archivo, el saldo corrido se corta
bloques = (df[CUENTA] != df[CUENTA].shift()).cumsum()
partidas = bloques.groupby(df[CUENTA]).nunique()
partidas = partidas[partidas > 1]

movs, saldos = df[~df["ES_SALDO"]], df[df["ES_SALDO"]]
visibles = ["NRO_LINEA"] + columnas
movs[visibles].to_csv(os.path.join(SALIDA, "movimientos.txt"), sep="|", index=False, encoding="utf-8")
saldos[visibles].to_csv(os.path.join(SALIDA, "saldos.txt"), sep="|", index=False, encoding="utf-8")
difs = df[df["SALDO_NO_CIERRA"]]
difs[visibles + ["SALDO_ESPERADO", "DIF_SALDO"]].to_csv(os.path.join(SALIDA, "diferencias_saldo.txt"),
                                                        sep="|", index=False, encoding="utf-8")
pd.DataFrame(descartadas, columns=["NRO_LINEA", "MOTIVO", "TEXTO"]).to_csv(
    os.path.join(SALIDA, "lineas_descartadas.txt"), sep="|", index=False, encoding="utf-8")

# ---------------------------------------------------------------- 3. CÓMO VIENEN LAS FILAS DE SALDO
# Para saber en qué columna traen el importe (debe, haber o saldo) cada tipo de fila de saldo
tipos = df.assign(TIPO_FILA=df[TIPO].where(df["ES_SALDO"], "(movimientos)")).groupby("TIPO_FILA").agg(
    filas=(CUENTA, "size"), cuentas=(CUENTA, "nunique"),
    con_debe=("_" + DEBE, lambda s: int((s != 0).sum())), total_debe=("_" + DEBE, "sum"),
    con_haber=("_" + HABER, lambda s: int((s != 0).sum())), total_haber=("_" + HABER, "sum"),
    con_saldo=("_" + SALDO, lambda s: int((s != 0).sum())),
    saldo_no_cierra=("SALDO_NO_CIERRA", "sum"))

# ---------------------------------------------------------------- 4. TOTALES DEL PIE
control_totales = pd.DataFrame()
if totales_pie:
    numeros = [x for x in totales_pie[1:] if re.fullmatch(r"-?[\d.]*,?\d*-?", x.replace(" ", "")) and x]
    pie = a_numero(pd.Series(numeros, dtype=str)).tolist()
    calculado = [movs["_" + DEBE].sum(), movs["_" + HABER].sum(), movs["_" + DEBE].sum() - movs["_" + HABER].sum()]
    control_totales = pd.DataFrame({"concepto": ["Total debe", "Total haber", "Debe - haber"][:len(pie)],
                                    "según el pie": pie[:3], "suma de los movimientos": calculado[:len(pie[:3])]})
    control_totales["diferencia"] = control_totales["según el pie"] - control_totales["suma de los movimientos"]

# ---------------------------------------------------------------- 5. RESUMEN POR CUENTA
datos = [c for c in DATOS_CUENTA if c in df.columns]
gm = movs.groupby(CUENTA, sort=False)
resumen = pd.DataFrame({
    "MOVIMIENTOS": gm.size(),
    "TOTAL_DEBE": gm["_" + DEBE].sum(), "TOTAL_HABER": gm["_" + HABER].sum(),
    "FECHA_PRIMER_MOV": gm["_FECHA"].min(), "FECHA_ULTIMO_MOV": gm["_FECHA"].max()})
resumen["NETO"] = resumen["TOTAL_DEBE"] - resumen["TOTAL_HABER"]
for c in PERIODO:
    if c in df.columns:
        resumen[f"{c}_DISTINTOS"] = gm[c].nunique()
todas = df.groupby(CUENTA, sort=False)
resumen = resumen.reindex(todas.size().index)              # incluye cuentas que solo tienen filas de saldo
conteos = ["MOVIMIENTOS"] + [f"{c}_DISTINTOS" for c in PERIODO if c in df.columns]
resumen[conteos] = resumen[conteos].fillna(0).astype(int)
resumen[["TOTAL_DEBE", "TOTAL_HABER", "NETO"]] = resumen[["TOTAL_DEBE", "TOTAL_HABER", "NETO"]].fillna(0)
resumen["FILAS_SALDO"] = df["ES_SALDO"].groupby(df[CUENTA], sort=False).sum().astype(int)
resumen["SALDO_FINAL"] = todas["_" + SALDO].last()
resumen["FILAS_SALDO_NO_CIERRA"] = todas["SALDO_NO_CIERRA"].sum().astype(int)
resumen["CUENTA_PARTIDA_EN_EL_ARCHIVO"] = resumen.index.isin(partidas.index)
for c in datos:
    resumen[c] = todas[c].first()
    resumen[f"{c}_DISTINTOS"] = todas[c].nunique()       # más de uno: la cuenta cambia de titular o hay error
resumen = resumen.reset_index()
resumen_txt = resumen.copy()
for c in ("FECHA_PRIMER_MOV", "FECHA_ULTIMO_MOV"):
    resumen_txt[c] = resumen_txt[c].dt.strftime("%d/%m/%Y")
resumen_txt.to_csv(os.path.join(SALIDA, "resumen_por_cuenta.txt"), sep="|", index=False, encoding="utf-8",
                   float_format="%.2f")

# ---------------------------------------------------------------- 6. EXCEL Y PANTALLA
def excel(t):
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    return t.map(lambda v: ILLEGAL_CHARACTERS_RE.sub("�", v) if isinstance(v, str) else v)

lineas_resumen = [
    ("Archivo", os.path.abspath(ARCHIVO)),
    ("Líneas del archivo", len(lineas)),
    ("Registros leídos", len(df)),
    ("  Movimientos", len(movs)),
    ("  Filas de saldo", len(saldos)),
    ("Líneas descartadas", len(descartadas)),
    ("Cuentas", df[CUENTA].nunique()),
    ("Cuentas sin movimientos (solo saldo)", int((resumen["MOVIMIENTOS"] == 0).sum())),
    ("Filas donde el saldo corrido no cierra", int(df["SALDO_NO_CIERRA"].sum())),
    ("Cuentas con al menos una fila que no cierra", int((resumen["FILAS_SALDO_NO_CIERRA"] > 0).sum())),
    ("Cuentas partidas en dos o más bloques del archivo", len(partidas)),
] + [(f"Importes no numéricos en {c}", n) for c, n in no_numericos.items()] \
  + [(f"Cuentas con más de un {c}", int((resumen[f"{c}_DISTINTOS"] > 1).sum())) for c in datos]
general = pd.DataFrame(lineas_resumen, columns=["control", "valor"])

with pd.ExcelWriter(os.path.join(SALIDA, "control_cuentas.xlsx"), engine="openpyxl") as xl:
    excel(general).to_excel(xl, sheet_name="RESUMEN", index=False)
    excel(tipos).to_excel(xl, sheet_name="TIPOS_DE_FILA")
    if len(control_totales):
        control_totales.to_excel(xl, sheet_name="TOTALES_DEL_PIE", index=False)
    excel(resumen).to_excel(xl, sheet_name="RESUMEN_POR_CUENTA", index=False)
    excel(difs[visibles + ["SALDO_ESPERADO", "DIF_SALDO"]].head(1_000_000)).to_excel(
        xl, sheet_name="DIFERENCIAS_SALDO", index=False)
    excel(pd.DataFrame(descartadas, columns=["NRO_LINEA", "MOTIVO", "TEXTO"]).head(1_000_000)).to_excel(
        xl, sheet_name="LINEAS_DESCARTADAS", index=False)

pd.options.display.float_format = "{:,.2f}".format
print(general.to_string(index=False))
print("\nCómo vienen las filas de saldo (con_* = cuántas filas tienen importe en esa columna):")
print(tipos.to_string())
if len(control_totales):
    print("\nTotales del pie contra la suma de los movimientos:")
    print(control_totales.to_string(index=False))
motivos = pd.Series([m for _, m, _ in descartadas]).value_counts()
if len(motivos):
    print("\nLíneas descartadas:")
    print(motivos.to_string())
print(f"\nResultados en la carpeta '{SALIDA}'")
