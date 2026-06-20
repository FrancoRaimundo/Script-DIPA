# -*- coding: utf-8 -*-
"""
Verifica que el entorno este listo en esta PC, SIN tocar MAIBA.
Genera datos de ejemplo (si no existen), corre el validador contra ellos y
chequea que todas las dependencias importen y que el cargador compile.

Uso:
  python probar_instalacion.py

Si termina con "TODO OK", la PC esta lista para usar el cargador (ver README.md).
NO carga nada en MAIBA: solo lee/escribe archivos locales de ejemplo.
"""
import sys
from pathlib import Path


def crear_datos_ejemplo():
    """Crea Aulas/Aula ejemplo (hoja + archivos dummy) si no existen. Asi el
    repo no necesita subir ninguna carpeta de datos: la prueba se autogenera."""
    import base64
    import pandas as pd
    base = Path("Aulas")
    hoja = base / "Aula ejemplo.xlsx"
    carp = base / "Aula ejemplo"
    if hoja.exists() and carp.exists():
        return
    carp.mkdir(parents=True, exist_ok=True)
    filas = [
        {"AULA":"ejemplo","CUIL":"20-11111111-2","Número de documento":"11111111","Nombre":"Juan","Apellido":"Perez","Correo":"juan@ejemplo.com","Teléfono":"1111111111","Distrito/municipio":"La Plata","Partido de residencia":"La Plata","Domicilio":"Calle 1 nro 100"},
        {"AULA":"ejemplo","CUIL":"27-22222222-3","Número de documento":"22.222.222","Nombre":"Maria","Apellido":"Gomez","Correo":"maria@ejemplo.com","Teléfono":"2222222222","Distrito/municipio":"Berisso","Partido de residencia":"Berisso","Domicilio":"Calle 2 nro 200"},
        {"AULA":"ejemplo","CUIL":"20-33333333-4","Número de documento":"33333333","Nombre":"Carlos","Apellido":"Lopez","Correo":"carlos@ejemplo.com","Teléfono":"3333333333","Distrito/municipio":"Ensenada","Partido de residencia":"Ensenada","Domicilio":"Calle 3 nro 300"},
    ]
    pd.DataFrame(filas).to_excel(hoja, index=False)
    pdf = b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF"
    jpg = base64.b64decode(
        "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAA//EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AfwD/2Q==")
    for dni in ("11111111", "22222222", "33333333"):
        (carp / f"{dni}.pdf").write_bytes(pdf)
    for dni in ("11111111", "22222222"):   # Carlos (33333333) queda sin .jpg a proposito
        (carp / f"{dni}.jpg").write_bytes(jpg)
    print("   (datos de ejemplo generados)")


print("== 1) Dependencias ==")
faltan = []
for mod in ("pandas", "openpyxl", "playwright"):
    try:
        __import__(mod)
        print(f"   [OK] {mod}")
    except Exception as e:
        print(f"   [FALTA] {mod}: {e}")
        faltan.append(mod)
if faltan:
    print("\n[ERROR] Faltan dependencias. Corre:  python -m pip install -r requirements.txt")
    sys.exit(1)

print("\n== 2) Chrome instalado ==")
import os
rutas = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
         r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
chrome = next((r for r in rutas if os.path.exists(r)), None)
if chrome:
    print(f"   [OK] {chrome}")
else:
    print("   [AVISO] No encontre chrome.exe en las rutas habituales.")
    print("           Editá la ruta en abrir_chrome_dipa.bat si tu Chrome esta en otro lado.")

print("\n== 3) Validador contra datos de ejemplo ==")
try:
    crear_datos_ejemplo()
    import validador_dipa as val
    # Aula 'ejemplo': Juan y Maria deberian dar LISTO; Carlos, a revisar (falta .jpg)
    val.validar(r"Aulas/Aula ejemplo.xlsx", "ejemplo", r"Aulas/Aula ejemplo")
except SystemExit as e:
    print("[ERROR] El validador fallo:", e)
    sys.exit(1)
except Exception as e:
    print("[ERROR] El validador fallo:", type(e).__name__, e)
    sys.exit(1)

print("\n== 4) El cargador compila ==")
import py_compile
try:
    py_compile.compile("cargador_dipa.py", doraise=True)
    print("   [OK] cargador_dipa.py compila")
except Exception as e:
    print("   [ERROR]", e)
    sys.exit(1)

print("\n=====================================================")
print(" TODO OK. La PC esta lista. Siguiente: leer README.md")
print(" (abrir_chrome_dipa.bat -> login -> aula -> cargar).")
print("=====================================================")
