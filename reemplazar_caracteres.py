"""
Reemplaza caracteres en un campo de un archivo de texto delimitado y guarda un archivo nuevo depurado.
El resto del archivo queda exactamente igual.

Uso:   python reemplazar_caracteres.py archivo.txt DESCRIPCION
Salida: archivo_depurado.txt (el original no se modifica)
"""
import re
import sys

# ---------------------------------------------------------------- QUÉ REEMPLAZAR (modificar acá)
REEMPLAZOS = {
    ";": ",",
    '"': "'",
    "\t": " ",
    "Ã‘": "Ñ",         # ejemplos de texto mal codificado (UTF-8 leído como Latin-1)
    "Ã±": "ñ",
}
QUITAR_CONTROL = True  # True = elimina los caracteres invisibles de control (nulos, etc.)
SEPARADOR = "|"

if len(sys.argv) < 3:
    sys.exit("Uso: python reemplazar_caracteres.py archivo.txt NOMBRE_DEL_CAMPO")
archivo, campo = sys.argv[1], sys.argv[2].strip().upper()
salida = re.sub(r"(\.\w+)?$", r"_depurado\1", archivo, count=1)
control = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

# Se lee y se escribe con la misma codificación que tiene el archivo (con o sin BOM)
with open(archivo, "rb") as f:
    bom = f.read(3) == b"\xef\xbb\xbf"
for codificacion in (("utf-8-sig" if bom else "utf-8"), "latin-1"):
    try:
        with open(archivo, encoding=codificacion, newline="") as f:
            # se corta solo en "\n": splitlines() también corta en caracteres de control y rompería registros
            lineas = [l + "\n" for l in f.read().split("\n")]
        lineas[-1] = lineas[-1][:-1]                       # la última no tenía "\n" agregado
        if lineas[-1] == "":
            lineas.pop()
        break
    except UnicodeDecodeError:
        continue

columnas = [c.strip().upper() for c in lineas[0].rstrip("\r\n").split(SEPARADOR)]
if campo not in columnas:
    sys.exit(f"No existe el campo {campo}. Campos: {', '.join(columnas)}")
pos = columnas.index(campo)

cuenta = {k: 0 for k in REEMPLAZOS}
cuenta["(caracteres de control)"] = 0
registros_cambiados, lineas_raras = 0, []
for i in range(1, len(lineas)):
    fin = lineas[i][len(lineas[i].rstrip("\r\n")):]          # conserva el salto de línea original
    partes = lineas[i][:len(lineas[i]) - len(fin)].split(SEPARADOR)
    if len(partes) != len(columnas):                          # línea rota: no se toca, se informa
        lineas_raras.append(i + 1)
        continue
    original = texto = partes[pos]
    if QUITAR_CONTROL:
        texto, n = control.subn("", texto)
        cuenta["(caracteres de control)"] += n
    for viejo, nuevo in REEMPLAZOS.items():
        cuenta[viejo] += texto.count(viejo)
        texto = texto.replace(viejo, nuevo)
    if texto != original:
        registros_cambiados += 1
        partes[pos] = texto
        lineas[i] = SEPARADOR.join(partes) + fin

with open(salida, "w", encoding=codificacion, newline="") as f:
    f.writelines(lineas)

print(f"Campo {campo}: {registros_cambiados:,} registros modificados de {len(lineas) - 1:,}")
for viejo, n in cuenta.items():
    if n:
        print(f"   {viejo!r}: {n:,} reemplazos")
if lineas_raras:
    print(f"ATENCIÓN: {len(lineas_raras):,} líneas no tienen {len(columnas)} campos y no se modificaron "
          f"(primeras: {', '.join(map(str, lineas_raras[:10]))}). Suele pasar cuando la descripción tiene "
          f"un salto de línea o el separador '{SEPARADOR}' adentro.")
print(f"Archivo generado: {salida}")
