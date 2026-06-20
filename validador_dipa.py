#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validador Fase 1 - Carga DIPA (Manipulacion de alimentos)
----------------------------------------------------------
Cruza la base de datos (hoja de Google exportada a .xlsx/.csv) contra la
carpeta de un aula y reporta, ANTES de cargar nada en DIPA:
  - Personas sin su archivo .pdf y/o .jpeg
  - Personas con datos obligatorios incompletos
  - Archivos sueltos en la carpeta que no corresponden a ninguna persona

No toca DIPA ni la hoja: solo lee y reporta. Riesgo cero.
"""

import re
import sys
import unicodedata
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Falta pandas. Instalalo con:  pip install pandas openpyxl")

try:
    import registro
except ImportError:
    registro = None   # el registro coloreado es opcional


# =========================== CONFIGURACION ===========================
# Editas estas lineas. Lo unico que cambia en cada tanda es AULA y
# RUTA_CARPETA_AULA; el resto queda fijo.

RUTA_HOJA = r"Aulas/Aula 1.xlsx"               # hoja exportada (.xlsx o .csv)
AULA = "1"                                     # valor de la columna "Aula"
RUTA_CARPETA_AULA = r"Aulas/Aula 1"            # carpeta con los .pdf y .jpeg

# Reconocimiento de columnas TOLERANTE: no importa mayusculas/minusculas,
# acentos ni espacios sobrantes en el encabezado. Cada campo logico acepta
# varios sinonimos (la hoja real usa "distrito"; la de prueba, "localidad").
ALIAS = {
    "aula":      ["aula"],
    "cuil":      ["cuil", "cuit", "cuit cuil"],
    "dni":       ["dni", "documento", "nro documento", "nro de documento",
                  "numero documento", "numero de documento", "nro doc", "num documento"],
    "nombre":    ["nombre", "nombres"],
    "apellido":  ["apellido", "apellidos"],
    "email":     ["correo", "email", "mail", "e mail", "correo electronico"],
    "celular":   ["telefono", "celular", "tel", "cel", "telefono celular"],
    "partido":   ["partido", "partido de residencia", "partido residencia"],
    "localidad": ["localidad", "distrito", "ciudad", "distrito municipio",
                  "municipio", "localidad distrito"],
    "domicilio": ["domicilio", "direccion"],
}

# Campos que deben estar completos en la hoja. Profesion/Oficio NO va aca:
# en la carga se resuelve con un guion "-" cuando haga falta.
OBLIGATORIOS = ["cuil", "dni", "nombre", "apellido", "email",
                "celular", "partido", "localidad", "domicilio"]
# =====================================================================


def _norm(texto):
    """Normaliza un encabezado: sin acentos, minusculas, y tratando cualquier
    signo de puntuacion (/, -, etc.) como espacio. Asi 'Distrito/municipio',
    'Número de documento' o 'CUIT/CUIL' se comparan de forma uniforme."""
    s = unicodedata.normalize("NFKD", str(texto).lower())
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return s.strip()


def resolver_columnas(df):
    """Mapea cada campo logico al nombre REAL de la columna en la hoja.
    Devuelve (mapping, faltantes)."""
    norm_a_real = {}
    for col in df.columns:
        norm_a_real.setdefault(_norm(col), col)
    mapping, faltantes = {}, []
    for logico, nombres in ALIAS.items():
        real = next((norm_a_real[n] for n in nombres if n in norm_a_real), None)
        if real:
            mapping[logico] = real
        else:
            faltantes.append(logico)
    return mapping, faltantes


def solo_digitos(valor):
    """Deja solo numeros (para comparar DNI con los nombres de archivo)."""
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    s = re.sub(r"\.0+$", "", s)          # corrige casos tipo "12345678.0"
    return re.sub(r"\D", "", s)


def cuil_normalizado(cuil_raw, dni_digits):
    """Devuelve el CUIL como 11 digitos, contemplando los formatos posibles:
        '12-12345678-9'  '12/12345678/9'  '12123456789'  -> 11 digitos directos
        '12-9'           -> prefijo y verificador; el DNI va en el medio,
                            asi que se reconstruye: prefijo + DNI(8) + verificador
    Devuelve '' si no se puede reconstruir."""
    if pd.isna(cuil_raw):
        return ""
    tokens = re.findall(r"\d+", str(cuil_raw))
    digitos = "".join(tokens)
    if len(digitos) == 11:
        return digitos
    if len(tokens) >= 2 and dni_digits:
        return tokens[0] + dni_digits.zfill(8) + tokens[-1]
    return ""


def vacio(valor):
    return pd.isna(valor) or str(valor).strip() == ""


def leer_hoja(ruta):
    p = Path(ruta)
    if not p.exists():
        sys.exit(f"[ERROR] No encuentro la hoja en: {p.resolve()}")
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p, dtype=str, keep_default_na=False)
    return pd.read_excel(p, dtype=str)


def indexar_archivos(carpeta):
    """Devuelve dos dicts {dni: nombre_archivo} para los .pdf y las imagenes."""
    pdfs, imgs = {}, {}
    for arch in carpeta.iterdir():
        if not arch.is_file():
            continue
        ext = arch.suffix.lower()
        dni = solo_digitos(arch.stem)
        if not dni:
            continue
        if ext == ".pdf":
            pdfs[dni] = arch.name
        elif ext in (".jpg", ".jpeg"):
            imgs[dni] = arch.name
    return pdfs, imgs


def validar(ruta_hoja, aula, ruta_carpeta, obligatorios=OBLIGATORIOS):
    df = leer_hoja(ruta_hoja)

    columnas, faltan = resolver_columnas(df)
    if faltan:
        print("[ERROR] No reconozco columna(s) para:", ", ".join(faltan))
        print("Columnas que tiene la hoja:", ", ".join(map(str, df.columns)))
        sys.exit(1)

    carpeta = Path(ruta_carpeta)
    if not carpeta.exists():
        sys.exit(f"[ERROR] No encuentro la carpeta del aula: {carpeta.resolve()}")

    col_aula = columnas["aula"]
    mask = (df[col_aula].fillna("").astype(str).str.strip().str.lower()
            == str(aula).strip().lower())
    df_aula = df[mask]

    if df_aula.empty:
        sys.exit(f"[ERROR] No hay filas con Aula = '{aula}'. "
                 f"Revisa el valor o el nombre de la columna '{col_aula}'.")

    pdfs, imgs = indexar_archivos(carpeta)
    dnis_hoja = set()
    filas_reporte = []

    for _, fila in df_aula.iterrows():
        dni = solo_digitos(fila[columnas["dni"]])
        dnis_hoja.add(dni)
        cuil = cuil_normalizado(fila[columnas["cuil"]], dni)

        datos_faltantes = [columnas[k] for k in obligatorios
                           if k != "cuil" and vacio(fila[columnas[k]])]
        tiene_pdf = dni in pdfs
        tiene_img = dni in imgs

        problemas = []
        if not dni:
            problemas.append("DNI vacio en la hoja")
        if len(cuil) != 11:
            problemas.append("CUIL no reconstruible")
        if not tiene_pdf:
            problemas.append("falta .pdf")
        if not tiene_img:
            problemas.append("falta .jpeg")
        if datos_faltantes:
            problemas.append("datos incompletos: " + ", ".join(datos_faltantes))

        filas_reporte.append({
            "DNI": dni,
            "Apellido": fila[columnas["apellido"]],
            "Nombre": fila[columnas["nombre"]],
            "CUIL": cuil,
            "PDF": pdfs.get(dni, "-"),
            "JPEG": imgs.get(dni, "-"),
            "Estado": "OK" if not problemas else "REVISAR",
            "Problemas": " | ".join(problemas),
        })

    huerfanos = []
    for d, n in pdfs.items():
        if d not in dnis_hoja:
            huerfanos.append(n)
    for d, n in imgs.items():
        if d not in dnis_hoja:
            huerfanos.append(n)
    huerfanos = sorted(set(huerfanos))

    reporte = pd.DataFrame(filas_reporte)
    total = len(reporte)
    ok = int((reporte["Estado"] == "OK").sum())
    revisar = total - ok

    print("=" * 60)
    print(f"  VALIDACION AULA '{aula}'")
    print("=" * 60)
    print(f"  Columnas reconocidas: " + ", ".join(f"{k}->{v}" for k, v in columnas.items()))
    print("-" * 60)
    print(f"  Personas en el aula : {total}")
    print(f"  Listas para cargar  : {ok}")
    print(f"  A revisar           : {revisar}")
    if huerfanos:
        print(f"  Archivos huerfanos  : {len(huerfanos)}")
    print("=" * 60)

    if revisar:
        print("\nPERSONAS A REVISAR:")
        for _, r in reporte[reporte["Estado"] == "REVISAR"].iterrows():
            etiqueta = f"{r['Apellido']} {r['Nombre']}".strip()
            print(f"  - DNI {r['DNI'] or '(vacio)'}  {etiqueta}: {r['Problemas']}")

    if huerfanos:
        print(f"\nARCHIVOS EN LA CARPETA SIN PERSONA EN EL AULA ({len(huerfanos)}):")
        for h in huerfanos[:20]:
            print(f"  - {h}")
        if len(huerfanos) > 20:
            print(f"  ... y {len(huerfanos) - 20} mas")

    salida = Path(f"reporte_validacion_aula_{str(aula).strip()}.csv")
    reporte.to_csv(salida, index=False, encoding="utf-8-sig")
    print(f"\nReporte detallado guardado en: {salida.resolve()}")

    # Registro coloreado (verde=listo, rojo=falta/incompleto): es la base
    # reanudable que luego el cargador va pintando de verde al cargar.
    if registro is not None:
        filas_reg = [{
            "dni": r["DNI"], "apellido": r["Apellido"], "nombre": r["Nombre"],
            "cuil": r["CUIL"],
            "estado": "LISTO" if r["Estado"] == "OK" else "FALTA",
            "detalle": r["Problemas"],
        } for r in filas_reporte]
        ruta_reg = Path(f"registro_aula_{str(aula).strip()}.xlsx")
        registro.escribir(ruta_reg, filas_reg)
        print(f"Registro coloreado (verde/rojo) guardado en: {ruta_reg.resolve()}")

    return reporte


if __name__ == "__main__":
    validar(RUTA_HOJA, AULA, RUTA_CARPETA_AULA)
