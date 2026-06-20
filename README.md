# DIPA — Carga automatizada de participantes (MAIBA)

Automatiza la carga de participantes del curso de Manipulación de Alimentos en la
plataforma **MAIBA** (`plataforma.maa.gba.gov.ar`), que **solo permite carga manual**
vía formulario. El script completa el formulario por vos, persona por persona, y
frena para que subas los 2 archivos (foto carnet + constancia) a mano.

> **Estado:** funcionando end-to-end. El formulario se completa solo; los archivos
> se suben a mano en una pausa (la subida automática quedó como mejora futura).

---

## 📁 Qué hay en la carpeta

| Archivo | Para qué sirve |
|---|---|
| `validador_dipa.py` | **Fase 1.** Cruza la hoja contra la carpeta del aula y reporta faltantes (riesgo cero, no toca MAIBA). |
| `cargador_dipa.py` | **Fase 2.** Carga los participantes en MAIBA imitando al operador. |
| `registro.py` | Módulo del tracker coloreado (`registro_aula_X.xlsx`): verde=cargado, amarillo=listo, rojo=falta. |
| `abrir_chrome_dipa.bat` | Abre **Chrome con depuración remota** (puerto 9222) para que el cargador se conecte. |
| `probar_instalacion.py` | Verifica el entorno (deps, Chrome, validador) **sin tocar MAIBA**. |
| `requirements.txt` | Dependencias de Python. |
| `Aulas/` | Datos de entrada: una subcarpeta por aula con los `.pdf`/`.jpg`, y la hoja `.xlsx` del aula. Incluye `Aula ejemplo` (datos ficticios de prueba). |
| `Manual-DIPA.md` | Documentación técnica/del proyecto (cómo funciona todo por dentro). |

---

## ✅ Requisitos (en CUALQUIER PC)

1. **Windows** (el `.bat` y las rutas están pensados para Windows).
2. **Python 3** (probado en 3.14). Verificar: `python --version`.
3. **Google Chrome** instalado.
4. Las dependencias de Python: `pandas`, `openpyxl`, `playwright`.

---

## 🚀 Instalación en una PC nueva (paso a paso)

> Si usás **Claude Code**, mostrale el archivo `README.md` y pedile que ejecute esta
> sección. Las instrucciones para él están en *"Notas para Claude Code"* al final.

### 1. Copiar la carpeta del proyecto
Copiá **toda la carpeta DIPA** a la PC nueva (idealmente en `Documentos\DIPA`).
Incluí: `cargador_dipa.py`, `validador_dipa.py`, `registro.py`,
`abrir_chrome_dipa.bat`, `requirements.txt`, y la carpeta `Aulas/` con los datos.

> **NO hace falta copiar** `.perfil_dipa_cdp/`, `.perfil_dipa/`, `__pycache__/`,
> ni los `registro_aula_X.xlsx`/`reporte_*.csv` (se regeneran). Si querés mantener
> el progreso de cargas, sí copiá los `registro_aula_X.xlsx`.

### 2. Instalar Python (si no está)
Descargar de https://www.python.org/downloads/ y **tildar "Add Python to PATH"**
durante la instalación. Verificar en una terminal nueva:
```
python --version
```

### 3. Instalar las dependencias
En la terminal, parado en la carpeta DIPA:
```
python -m pip install -r requirements.txt
```

### 4. (Opcional) Navegador de Playwright
El cargador usa **tu Google Chrome** (no hace falta el navegador de Playwright).
Solo si alguna vez usaras `CONEXION = "propio"`, correr una vez:
```
python -m playwright install chromium
```

### 5. Verificar la ruta de Chrome en el `.bat`
Abrí `abrir_chrome_dipa.bat` con el Bloc de notas y confirmá que la ruta de Chrome
exista. Por defecto busca:
```
C:\Program Files\Google\Chrome\Application\chrome.exe
C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
```
Si tu Chrome está en otro lado, corregí la línea `set "CHROME=..."`.

### 6. Probar que todo quedó bien (sin tocar MAIBA)
```
python probar_instalacion.py
```
Chequea dependencias, Chrome, corre el validador contra los **datos de ejemplo**
(`Aulas/Aula ejemplo`) y verifica que el cargador compile. Si termina con **"TODO OK"**,
la PC está lista. **No carga nada en MAIBA** (solo lee archivos locales).

> Los datos de ejemplo (`Aulas/Aula ejemplo.xlsx` + carpeta `Aula ejemplo/`) son
> ficticios y solo sirven para esta prueba. **No los cargues en MAIBA.**

---

## ⚙️ Configuración por aula (cada vez que cambiás de curso)

Abrí `cargador_dipa.py` y editá el bloque **CONFIGURACION** (arriba del todo):

```python
RUTA_HOJA = r"Aulas/Aula 1.xlsx"      # la hoja .xlsx del aula
AULA = "1"                            # valor de la columna "Aula" en la hoja
RUTA_CARPETA_AULA = r"Aulas/Aula 1"   # carpeta con los .pdf y .jpg
URL_AULA = r"https://plataforma.maa.gba.gov.ar/ma/cursos/editar/27698"  # la del lápiz verde
```

Otros ajustes útiles (no hace falta tocarlos normalmente):
- `LOTE_DEFECTO = 5` → cuántos cargás por tanda.
- `ESPERA_AUTOCOMPLETADO_S = 7` → cuánto espera el autocompletado por CUIL.
- `SUBIR_ARCHIVOS_AUTO = False` → archivos a mano (recomendado por ahora).
- `MODO = "revision"` → frena y pide confirmación antes de guardar cada lote.

---

## ▶️ Uso (rutina de trabajo)

### Paso 0 — Validar (opcional pero recomendado)
```
python validador_dipa.py
```
Te dice quiénes están listos (tienen pdf+jpg+datos) y quiénes faltan. Genera
`reporte_validacion_aula_X.csv` y `registro_aula_X.xlsx`.

### Paso 1 — Abrir Chrome conectable
**Doble clic en `abrir_chrome_dipa.bat`**. Va a:
- Avisarte que cierra todo Chrome (guardá tu trabajo) → apretá una tecla.
- Abrir **una sola** ventana de Chrome en el login de DIPA.

> ⚠️ **Importante:** usá SOLO esa ventana. Si abrís otra ventana de Chrome aparte,
> el script no la va a ver.

### Paso 2 — Login y navegación (a mano, en ESA ventana)
1. Logueate (toda la cadena SSO: DIPA → sso.gba → sso-gdeba).
2. Navegá hasta el **aula** correcta → pestaña **Participantes** (que se vea el botón
   "Cargar Participante").
3. La sesión queda guardada en `.perfil_dipa_cdp/`, así que la próxima vez ya entrás
   logueado.

### Paso 3 — Correr el cargador
En la terminal, en la carpeta DIPA:
```
python cargador_dipa.py cargar
```
- `cargar` → lote por defecto (5).
- `cargar 8` → lote de 8.
- `cargar todas` → todas las pendientes.

### Qué hace por cada participante
1. Click "Cargar Participante" → completa el formulario (CUIL, datos, Partido/Localidad,
   Profesión).
2. Detecta si **autocompletó** (persona existente):
   - **Sí (Caso A):** no pisa los datos; marca diferencias contra la hoja; completa
     Localidad si vino vacía; quita imágenes viejas (X roja).
   - **No (Caso B):** completa todos los campos obligatorios.
   - Si a los 7s no está seguro, **te pregunta** `¿Se cargó el form? (s/n)`.
3. **FRENA** con el formulario abierto → vos subís **Foto carnet** y **Constancia** a mano.
4. Apretás **Enter** → "Guardar participante" (lo agrega a la tabla) → siguiente.
5. Tras los X del lote → **Actualizar** + **Confirmar** (una sola vez) → persiste todo.

### Reanudar
El tracker `registro_aula_X.xlsx` marca en **verde** a los cargados. Si cortás y
volvés a correr, **saltea los verdes** automáticamente (no duplica).

---

## 🆘 Problemas comunes

| Síntoma | Solución |
|---|---|
| `No me pude conectar a tu Chrome por CDP` | No corriste `abrir_chrome_dipa.bat`, o cerraste esa ventana. Abrila de nuevo. |
| `No encuentro ninguna pestaña en el aula` (lista solo `logindpsit`) | Estás navegando en otra ventana de Chrome. Usá **solo** la del `.bat`. Cerrá las demás. |
| `No encontré chrome.exe` (en el `.bat`) | Corregí la ruta de Chrome en `abrir_chrome_dipa.bat`. |
| `Falta pandas/playwright` | `python -m pip install -r requirements.txt`. |
| El autocompletado tarda y se lee mal | Subí `ESPERA_AUTOCOMPLETADO_S` a 10-12 en `cargador_dipa.py`. |
| El modal "Confirmar" no se aprieta solo | Confirmalo a mano; la carga igual se guarda. |

---

## 🤖 Notas para Claude Code (setup en PC nueva)

Si te piden preparar este proyecto en una PC nueva, hacé en orden:

1. **Verificá Python:** `python --version` (debe ser 3.x). Si no está, avisá que hay
   que instalarlo desde python.org con "Add to PATH".
2. **Instalá dependencias:** `python -m pip install -r requirements.txt`.
3. **Verificá Chrome:** confirmá que exista `chrome.exe` en
   `C:\Program Files\Google\Chrome\Application\` o `C:\Program Files (x86)\...`.
   Si está en otra ruta, editá `abrir_chrome_dipa.bat` (línea `set "CHROME=..."`).
4. **Corré la verificación completa:** `python probar_instalacion.py`. Debe terminar
   con **"TODO OK"** (chequea deps, Chrome, valida los datos de ejemplo y compila el
   cargador). Si falla, seguí lo que indique el error.
5. **Ajustá la config del aula** en `cargador_dipa.py`: `RUTA_HOJA`, `AULA`,
   `RUTA_CARPETA_AULA`, `URL_AULA` (pedíselos al usuario si no los sabés).
6. **Recordale al usuario el flujo de uso:** doble clic en `abrir_chrome_dipa.bat` →
   login + navegar al aula > Participantes en ESA ventana → `python cargador_dipa.py cargar`.

**Cosas que NO debés hacer:**
- No manejes usuario/contraseña: el login lo hace siempre el operador a mano.
- No cambies a `CONEXION = "propio"` salvo que el `.bat`/CDP fallen (es menos confiable
  en esta plataforma).
- El script depende de la **conexión CDP** a un Chrome abierto con el `.bat`: sin eso,
  no puede ver la pestaña del aula.

**Detalles técnicos clave** (por si tenés que ajustar selectores): ver
`Manual-DIPA.md`, sección 7 (reconocimiento del formulario MAIBA).
