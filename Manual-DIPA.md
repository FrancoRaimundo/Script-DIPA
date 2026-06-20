# Proyecto: Automatización de carga de participantes en DIPA

**Curso:** Manipulación de alimentos
**Plataforma destino:** DIPA (carga manual vía formulario web)
**Última actualización:** Junio 2026

---

## 1. Objetivos del proyecto

1. **Eliminar el trabajo manual repetitivo** de cargar participantes en DIPA, dado el alto volumen de información a revisar y cargar.
2. **Garantizar la integridad de los datos antes de cargar**: que cada persona tenga sus dos archivos y todos los datos obligatorios completos, detectando inconsistencias por adelantado.
3. **Procesar por aula/grupo**, indicando en cada tanda qué aula se trabaja.
4. **Minimizar errores en documentación oficial** (los carnets), incorporando un paso de revisión humana antes de cada guardado.

---

## 2. Cómo se reciben los datos (origen)

### 2.1 Estructura de carpetas (confirmada)

- Una **carpeta madre `Aulas/`** que contiene:
  - **Una subcarpeta por aula** (ej. `Aula 1`, `Aula 2`, …) con los archivos.
  - **El archivo de datos** del aula (`.csv` o `.xlsx`) al lado de las subcarpetas.
- Cada subcarpeta contiene, **por participante**, dos archivos nombrados con el **DNI**:
  - `{DNI}.pdf` → imágenes del DNI físico → campo **Constancia documento**.
  - `{DNI}.jpeg` / `.jpg` / `.JPG` → foto tipo carnet → campo **Foto carnet actualizada**.
- Ejemplo: `16143471.pdf` y `16143471.JPG` pertenecen a la misma persona. La extensión puede venir en mayúscula; los scripts la normalizan.

### 2.2 Base de datos (hoja de Google exportada)

Contiene los datos personales. Columnas relevantes:

`Aula`, `Cuil`, `DNI`, `Nombre`, `Apellido`, `Correo`, `Telefono` (celular), `Partido`, `distrito`/`localidad`, `Domicilio`.

> **Reconocimiento tolerante de columnas:** los encabezados se reconocen sin importar **mayúsculas, acentos ni espacios sobrantes**, y aceptan **sinónimos** (`Correo`/`correo`/`email`, `Telefono`/`celular`, `distrito`/`localidad`, etc.). Así funciona tanto con la hoja real como con hojas de prueba con encabezados distintos.

> **Formato recomendado: `.xlsx`.** Conviene sobre `.csv` porque (a) no rompe los **ceros a la izquierda** del DNI/CUIL, (b) maneja bien **acentos** sin líos de codificación, (c) evita el problema de comas dentro del domicilio, y sobre todo (d) permite el **registro coloreado** (ver §5.2.1): el seguimiento verde/rojo que hoy se hace a mano **solo es posible en xlsx**. Los scripts igual leen ambos formatos.

### 2.3 Clave de relación

El **DNI** es el identificador único que conecta las tres fuentes: la fila en la hoja, los dos archivos en la carpeta y el registro a cargar en DIPA. Esto permite automatizar sin ambigüedad.

---

## 3. Cómo se cargan los datos (destino: DIPA)

### 3.1 Restricciones de la plataforma

La plataforma es **MAIBA** (`plataforma.maa.gba.gov.ar`), de la Provincia de Buenos Aires.

- **No tiene API ni importación masiva (CSV/Excel): todo es manual** vía formulario.
- **Login:** usuario y contraseña.
- **Flujo real (confirmado con capturas):**
  1. `/ma/entidadmanipuladora/334` → **lista de cursos** (cada curso funciona como un "aula").
  2. El **lápiz verde** de un curso abre `/ma/cursos/editar/{ID}` → **wizard de 4 pasos** (Curso · Docentes · Participantes · Confirmación).
  3. En el paso **3 "Participantes"** está la tabla de inscriptos y el botón azul **"Cargar Participante"**, que abre el formulario.
  4. Al **"Guardar participante"** vuelve a la tabla; se repite para el siguiente.

> Nota: el "aula" de la hoja/carpetas se corresponde con un **curso** de MAIBA. El operador identifica la URL del curso (`.../cursos/editar/{ID}`) y la pone en `URL_CURSO`.

### 3.2 Mapeo de columnas → campos del formulario

| Columna en la hoja | Campo en DIPA |
|---|---|
| Cuil | CUIT/CUIL |
| DNI | Número documento |
| Nombre | Nombre |
| Apellido | Apellido |
| Correo | Email |
| Telefono | **Celular** |
| Partido | Partido |
| distrito | Localidad |
| Domicilio | Domicilio |
| *(fijo)* | Tipo documento = **DNI** |
| `{DNI}.jpeg` | Foto carnet actualizada |
| `{DNI}.pdf` | Constancia documento |
| *(ver casos especiales)* | Profesión / Oficio |
| Aula | *No va al formulario; sirve para filtrar filas y ubicar la carpeta* |

> Nota: la columna **`Telefono`** de la hoja corresponde al campo **Celular** de DIPA. El campo **Teléfono** de DIPA (opcional) se ignora.

### 3.3 Reglas de comportamiento del formulario

- **Campos bloqueados hasta ingresar el CUIL:** recién al cargar el CUIL se habilitan los demás campos. Por eso el CUIL siempre va primero.
- **Tipo documento:** fijo en **DNI**, no se modifica.
- **Localidad bloqueada hasta elegir Partido:** al seleccionar el Partido se habilita Localidad con las opciones correspondientes (desplegable dependiente que carga sus opciones después de elegir el partido).
- **Solo se completan campos obligatorios** (los del asterisco rojo). Los opcionales (Teléfono, Lugar de trabajo) **no se completan**.
- **Dos cargas de archivo:** Foto carnet y Constancia documento.

---

## 4. Casos especiales a tener en cuenta

1. **Persona ya cargada en el sistema (autocompletado por CUIL) — MUY probable.**
   Al ingresar el CUIL, la plataforma puede autocompletar los datos. En ese caso:
   - **No se recarga** la persona; se **revisa que los datos coincidan** con la hoja.
   - Se **reemplazan los archivos**: si ya aparecen los anteriores (con la **X roja**), se borran y se suben los nuevos de la carpeta del curso.
   - Puede pasar que **los archivos no aparezcan** aunque los datos sí se hayan autocompletado; en ese caso simplemente se suben.

2. **Profesión / Oficio.**
   - Si la persona se autocompleta y el campo **trae un valor**, se deja tal cual.
   - Si **no se autocompleta** o el campo **está vacío**, se carga un guion `-` para poder guardar. El dato no es relevante por ahora.

3. **Localidad faltante en la hoja.**
   A veces el dato no está completo. Como regla de *fallback*, **casi siempre existe una localidad que coincide con el nombre del partido**, y esa es la opción por defecto.

4. **Diferencias entre datos autocompletados y la hoja.**
   Cuando un dato autocompletado no coincide con la hoja, se **marca para revisión pero no se pisa automáticamente** (editar un registro existente es delicado). La decisión queda en manos del operador, en el paso de revisión.

5. **Formato del DNI.**
   En la hoja el DNI puede venir con puntos (`12.345.678`) mientras que el archivo es plano (`12345678.pdf`). Se **normaliza a solo dígitos** para que matcheen.

8. **Formatos del CUIL (varios).** La columna Cuil puede venir de distintas formas; el script las **normaliza todas a 11 dígitos** antes de tipear:
   - `12-12345678-9`, `12/12345678/9`, `12123456789` → se toman los 11 dígitos directos.
   - `12-9` (solo **prefijo y verificador**): el DNI va **en el medio**, así que se reconstruye como `prefijo + DNI(8 dígitos) + verificador`. Ej.: Cuil `12-3` + DNI `12396854` → `12123968543`.
   - Si no se llega a 11 dígitos, se marca **"CUIL no reconstruible"** y la persona queda a revisar.

6. **Extensión de la imagen.**
   Se aceptan `.jpeg` y `.jpg` (y sus variantes en mayúscula).

7. **Faltantes y sobrantes.**
   - Persona sin su `.pdf` y/o `.jpeg`.
   - Persona con datos obligatorios incompletos.
   - Archivos sueltos en la carpeta que no corresponden a ninguna persona del aula (huérfanos).

---

## 5. Plan por fases

### Fase 0 — Reconocimiento y diseño ✅ HECHO
Definición del flujo completo, confirmación de que DIPA no tiene API ni carga masiva, identificación del tipo de login, del comportamiento de los desplegables dependientes (Partido → Localidad) y de la secuencia de desbloqueo de campos (CUIL primero). Mapeo de columnas a campos. Definición de casos especiales.

### Fase 1 — Validador de datos ✅ HECHO
Script local (`validador_dipa.py`) que cruza la hoja contra la carpeta de un aula **antes** de cargar nada en DIPA. No toca DIPA ni la hoja: solo lee y reporta. **Riesgo cero.**

**Qué valida:**
- Que cada persona del aula tenga su `.pdf` y su `.jpeg`.
- Que los datos obligatorios estén completos en la hoja.
- Que no haya archivos huérfanos en la carpeta.

**Entradas:** hoja exportada (`.xlsx`/`.csv`), aula a procesar, carpeta del aula.
**Salidas:** resumen en pantalla (listas para cargar / a revisar / huérfanos) + `reporte_validacion_aula_X.csv` (detalle) + `registro_aula_X.xlsx` (tracker coloreado, ver abajo).

#### 5.2.1 Registro coloreado (tracker verde/rojo) — `registro.py`
Automatiza lo que el operador hacía **a mano** (pintar la casilla del DNI). Es un `.xlsx` con una fila por participante y la **casilla del DNI pintada** según el estado, y funciona como **fuente de verdad reanudable**:

| Color | Estado | Significado |
|---|---|---|
| 🟡 Amarillo | `LISTO` | Validado y a la espera de cargar |
| 🟢 Verde | `CARGADO` | Cargado correctamente en MAIBA |
| 🔴 Rojo | `FALTA` / `INCOMPLETO` / `ERROR` | Falta pdf/jpeg, datos incompletos, o falló la carga |

- El **validador** lo crea con cada persona en 🟡 LISTO o 🔴 FALTA.
- El **cargador** lo lee para **saltear** a los que ya están en 🟢 (reanudar) y lo va pintando: 🟢 al guardar, 🔴 si hubo error.
- (El amarillo es un agregado útil sobre el esquema manual verde/rojo: distingue "pendiente" de "ya cargado".)

### Fase 2 — Cargador automático ✅ FUNCIONANDO (probado en producción)
Automatización del navegador (Python + **Playwright**) que imita al operador en el
formulario de MAIBA, persona por persona, en lotes. Implementado en `cargador_dipa.py`.
Se probó cargando participantes reales del Aula 1 de a 1, 3 y 5. **El formulario se
completa solo; los archivos (foto+constancia) se suben a mano** en una pausa (ver §5.2.3).

> **Cómo se usa, en concreto, está en `README.md`.** Acá se documenta el *cómo funciona*.

#### 5.2.2 Conexión al navegador: CDP (clave para que funcione)
La pieza que costó resolver. En esta plataforma, hacer que Playwright **lance** su
propia ventana NO sirve: el login es una cadena SSO que abre ventanas que Playwright
no controla, y si hay otro Chrome abierto, la navegación del operador cae en una
ventana que el script no ve (se quedaba "ciego" en `logindpsit`).

**Solución (la que funciona):** el operador abre Chrome con **`abrir_chrome_dipa.bat`**
(cierra todo Chrome y abre **uno solo** con `--remote-debugging-port=9222` y perfil
`.perfil_dipa_cdp`). El cargador se **conecta** a ESE Chrome con
`connect_over_cdp("http://localhost:9222")` → ve **todas** las pestañas reales del
operador, sin importar la cadena de login. El operador navega a mano hasta el aula >
Participantes; el script **identifica la pestaña por la URL exacta del aula**
(`editar/<id>`). Config: `CONEXION = "cdp"`.

#### 5.2.3 Subida de archivos (foto carnet + constancia) — manual por ahora
El widget es **jasny `fileinput` + bootstrap `filestyle`**, con subida **AJAX** al
seleccionar (la imagen va a `/uploads/...` y el preview la muestra). `set_input_files`
no dispara esa subida de forma confiable (evento no "trusted"), así que:
- **`SUBIR_ARCHIVOS_AUTO = False` (default):** el script completa todo el resto, en
  Caso A **quita las imágenes viejas** (X roja = `[data-dismiss="fileinput"]`), y
  **frena** con el formulario abierto para que el operador suba los 2 archivos a mano;
  Enter → "Guardar participante".
- `SUBIR_ARCHIVOS_AUTO = True` (experimental): intenta subirlos con el file chooser
  de Playwright y espera a que el preview apunte a `/uploads/`. Poco confiable aún.

#### 5.2.4 Selectores reales del formulario MAIBA
- Subform de participante: `...curso[participantes][N][datoPersona][campo]`. Se usa el
  **sufijo del id/name** y siempre el **último** (`.last`): los participantes ya
  guardados quedan como inputs ocultos (índices 0..N-1) y el subform abierto es el último.
- Campos: `cuit`, `nombre`, `apellido`, `numeroDocumento`, `email`, `celular`,
  `domicilio`; **Profesión = `<textarea name$="[ocupacion]">` `readonly`** (se setea por
  JS quitando el readonly; el dato es `-`).
- Selects **Select2**: `tipoDocumento` (opción "DNI"), `partido`, `localidad`. Se
  operan por JS (setear value + `change` jQuery), con match **sin acentos**.
- Archivos: `input[type=file][name$="[fotoCarnet][archivo]"]` y `[constanciaDni][archivo]`.
  (Filtrar `type=file`: tras autocompletar aparece un `hidden` con el mismo name.)
- Botones: "Cargar Participante" y "Guardar participante" son **`<a>`**; "Actualizar"
  es `<button type=submit>`.

#### 5.2.5 Guardado en TRES pasos (la plataforma lo exige)
1. **"Guardar participante"** → agrega la persona a la tabla del wizard (en memoria).
2. **"Actualizar"** (verde) → envía el curso.
3. **Modal "¿Desea guardar el curso?" → "Confirmar"** → persiste de verdad.
Se hace **una sola vez por lote** (no por persona). Tras confirmar, MAIBA redirige a
la lista de cursos; el operador vuelve al aula para el próximo lote.

#### 5.2.6 Detección de autocompletado (Caso A vs B)
Al tipear el CUIL, MAIBA puede autocompletar (persona existente). El timing es
variable. El script espera hasta `ESPERA_AUTOCOMPLETADO_S` (7s): si aparece el
Nombre → **Caso A**; si a los 7s sigue vacío, **pregunta al operador**
`¿Se cargó el form? (s/n)` (que está mirando la pantalla) y espera su respuesta.
- **Caso A:** no pisa los datos de texto; marca diferencias contra la hoja; completa
  Partido/Localidad si vinieron vacíos; Localidad **no** es dependiente (trae las 5351).
- **Caso B:** completa todos los obligatorios (Tipo doc=DNI, datos, Partido→Localidad,
  Profesión=`-`).

**Otras decisiones de diseño:**
- Reutiliza la Fase 1 (`validador_dipa`): normalización de DNI/CUIL, lectura de hoja,
  reconocimiento tolerante de columnas. Solo carga personas que pasan la validación.
- CUIL enmascarado → se tipea dígito a dígito.
- **El login lo hace siempre el operador**; el script nunca maneja usuario ni contraseña.
- **Reanudable** por el tracker `registro_aula_X.xlsx` (saltea verdes).
- **Lotes** configurables (`LOTE_DEFECTO`, o `cargar N` / `cargar todas`).

### Fase 3 — Robustez y operación 🔵 PARCIAL
- **Conexión CDP + login del operador:** ✅ resuelto (perfil `.perfil_dipa_cdp`).
- **Proceso reanudable:** ✅ tracker `registro_aula_X.xlsx` (saltea los verdes).
- **Subida automática de archivos:** ⏳ pendiente (hoy manual). Mejora futura:
  capturar el endpoint AJAX real de subida y replicarlo.
- **Manejo de errores y reintentos / modo automático:** ⏳ pendiente.

---

## 6. Estado actual

| Fase | Estado |
|---|---|
| Fase 0 — Reconocimiento y diseño | ✅ Hecho |
| Fase 1 — Validador de datos | ✅ Hecho y probado (`validador_dipa.py`) |
| Fase 2 — Cargador automático | ✅ **Funcionando en producción** (`cargador_dipa.py`) — probado de a 1/3/5 |
| Fase 3 — Robustez y operación | 🔵 Parcial (CDP + login operador + reanudable OK; falta auto-subida de archivos y modo automático) |

> **Uso operativo y portabilidad a otra PC:** ver `README.md`.

---

## 7. Aprendizajes clave (lo que costó resolver)

Resumen de los problemas reales que se encontraron al llevarlo a producción, para no
repetirlos (el detalle técnico está en §5.2):

1. **No se podía "ver" la navegación del operador.** Lanzar Chrome desde Playwright
   no captura la cadena de login SSO ni convive con otro Chrome abierto. → Solución:
   **conexión CDP** a un Chrome que abre el operador con `abrir_chrome_dipa.bat`
   (`--remote-debugging-port=9222`), y detección de la pestaña por la **URL exacta del aula**.
2. **Identificar la pestaña correcta.** El "Participantes" del menú superior y la lista
   de cursos (`/entidadmanipuladora/334`) tienen botones parecidos → se identifica por
   `editar/<id>`, no por texto.
3. **Profesión/Oficio** es un `<textarea readonly>` "search-dependent" → se setea por JS.
4. **Selects Select2** → setear value + `change` jQuery; match sin acentos.
5. **Tras autocompletar aparece un `input hidden`** con el mismo name que el file →
   filtrar `type=file` y usar `.last`.
6. **Guardado en 3 pasos:** Guardar participante → Actualizar → modal **Confirmar**.
7. **Subida de archivos AJAX (jasny/filestyle)** no se dispara confiable por código →
   **se suben a mano** en una pausa (mejora futura: replicar el endpoint AJAX).
8. **Timing del autocompletado** variable → espera 7s + pregunta `s/n` al operador si
   sigue en duda.

### Mejoras futuras
- **Auto-subida de archivos:** capturar (con la pestaña de red del navegador) el
  endpoint AJAX real al subir un archivo a mano, y replicar ese POST desde el script.
- **Modo automático** (`MODO = "automatico"`): saltear las confirmaciones manuales.
- **Reintentos** ante errores transitorios.

---

## 8. Stack técnico

- **Lenguaje:** Python 3 (probado en 3.14).
- **Automatización del navegador:** Playwright, **conectado por CDP** al Google Chrome
  instalado (no al Chromium de Playwright).
- **Lectura de datos:** pandas + openpyxl (hoja exportada a `.xlsx`/`.csv`).
- **Tracker coloreado:** openpyxl (`registro.py`, celdas verde/amarillo/rojo).
- **Criterio de seguridad:** el login lo hace siempre el operador; el script no maneja
  usuario ni contraseña.

### Archivos del proyecto
- `README.md` — **instalación, uso y portabilidad a otra PC** (empezar por acá).
- `validador_dipa.py` — Fase 1: valida y genera el reporte + el tracker coloreado.
- `registro.py` — módulo del tracker `registro_aula_X.xlsx` (verde/amarillo/rojo, reanudable).
- `cargador_dipa.py` — Fase 2: carga en MAIBA (conexión CDP, lotes, modo revisión).
- `abrir_chrome_dipa.bat` — abre Chrome con depuración remota para el cargador.
- `requirements.txt` — dependencias de Python.
- `Aulas/` — datos de entrada: subcarpeta por aula + hoja (`.csv`/`.xlsx`).
