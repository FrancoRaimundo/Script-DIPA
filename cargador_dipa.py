#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cargador Fase 2 - Carga DIPA / MAIBA (Manipulacion de alimentos)
----------------------------------------------------------------
Automatiza el formulario de participantes de la plataforma MAIBA
(plataforma.maa.gba.gov.ar) imitando al operador, persona por persona, para
el curso (aula) indicado. Arranca en MODO REVISION: completa todo y FRENA
antes de "Guardar participante" para que un humano confirme.

FLUJO REAL DE LA PLATAFORMA (confirmado con capturas):
  1. /ma/entidadmanipuladora/334  -> lista de cursos. El lapiz verde de cada
     curso lleva a /ma/cursos/editar/{ID}.
  2. /ma/cursos/editar/{ID}  -> wizard de 4 pasos. Interesa el paso 3,
     "Participantes", donde esta el boton azul "Cargar Participante".
  3. "Cargar Participante" abre el formulario (CUIT/CUIL primero, etc.).
     Al guardar, vuelve a la tabla de participantes y se repite.

PRINCIPIOS DE SEGURIDAD (Manual, seccion 8):
  - El login lo hace SIEMPRE el operador a mano. El script NO maneja usuario
    ni contrasena. La sesion queda en un perfil de navegador persistente.
  - Modo revision por defecto: nada se guarda sin confirmacion humana.
  - Solo carga personas que ya pasaron el validador (Fase 1) sin problemas.
  - Proceso reanudable: registro CSV de cargados para no repetir.

EL SCRIPT NO NAVEGA: el operador deja la ventana de Chrome ya posicionada en
el AULA correcta, pestana 'Participantes' (con el boton 'Cargar Participante'
a la vista). El script se concentra en cargar y confirmar.

DOS MODOS DE USO:
  python cargador_dipa.py capturar
      Vuelca todos los campos del formulario a 'formulario_dipa.txt' para
      verificar selectores si algo cambiara en la plataforma.

  python cargador_dipa.py cargar [X]
      Carga hasta X participantes (lote): por cada uno hace "Cargar
      Participante" -> llena -> "Guardar participante" (a la tabla); y al
      final del lote aprieta "Actualizar" UNA vez para persistir todo.
      Sin X, carga todas las pendientes. Ej: cargar 1  (prueba de a una).

Requiere: pip install playwright pandas openpyxl  &&  playwright install chromium
"""

import sys
from pathlib import Path

# Reutilizamos la logica ya probada del validador (Fase 1): normalizacion de
# DNI, lectura de la hoja e indexado de archivos. No se duplica nada.
try:
    import validador_dipa as val
    import registro
except ImportError:
    sys.exit("[ERROR] Faltan validador_dipa.py / registro.py en esta carpeta. "
             "El cargador reutiliza sus funciones.")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    sys.exit("Falta Playwright. Instalalo con:\n"
             "  pip install playwright\n"
             "  playwright install chromium")


# =========================== CONFIGURACION ===========================
# Lo unico que cambia en cada tanda es AULA y RUTA_CARPETA_AULA.
# El script abre SU PROPIA ventana de Chrome (perfil persistente) y la deja en
# la pagina de inicio de DIPA. EL OPERADOR hace TODO el login (cadena SSO) y la
# navegacion interna hasta el aula > pestana Participantes, EN ESA VENTANA. El
# script espera y detecta cuando llegaste. La sesion queda guardada para la
# proxima (no habra que re-loguear cada vez).

RUTA_HOJA = r"Aulas/Aula 1.xlsx"               # misma hoja que usa el validador
AULA = "1"                                     # valor de la columna "Aula"
RUTA_CARPETA_AULA = r"Aulas/Aula 1"            # carpeta con los .pdf y .jpeg

# Pagina de inicio de DIPA (entrada correcta al login). El script solo abre
# ESTA url para anclar la ventana; el resto lo navegas vos a mano.
URL_INICIO = r"https://plataforma.maa.gba.gov.ar/logindpsit"

# URL del AULA a cargar (la del lapiz verde, .../ma/cursos/editar/ID). El script
# la usa para identificar EXACTAMENTE la pestana correcta cuando apretas Enter.
# Cambiala por aula en cada tirada.
URL_AULA = r"https://plataforma.maa.gba.gov.ar/ma/cursos/editar/27698"

# Como se conecta al navegador:
#   "cdp"    -> el operador abre Chrome con 'abrir_chrome_dipa.bat' (depuracion
#               remota) y el script se conecta a ESE Chrome. RECOMENDADO: ve
#               todas tus pestanas reales, sirve aunque el login sea complejo.
#   "propio" -> el script lanza su propia ventana (fallaba si hay otro Chrome).
CONEXION = "cdp"
CDP_URL = "http://localhost:9222"

PERFIL_NAVEGADOR = r".perfil_dipa"       # carpeta donde se guarda la sesion
MODO = "revision"                        # "revision" (confirma el lote antes de Actualizar) | "automatico"

# Subida de archivos: False = vos los subis a mano en la pausa (el script igual
# completa todo el resto y, en Caso A, quita las imagenes viejas). True = el
# script intenta subirlos solo (file chooser) -- por ahora poco confiable.
SUBIR_ARCHIVOS_AUTO = False

# Cuantos participantes por tanda (lote). Se puede pisar por linea de comandos:
# 'python cargador_dipa.py cargar 8'. Si no se pasa numero, usa este.
LOTE_DEFECTO = 5

# Segundos a esperar el autocompletado por CUIL. Si se llena Nombre antes, sigue
# de una (Caso A). Si a los X segundos sigue vacio, asume persona nueva (Caso B).
# Subilo si ves autocompletados muy lentos que se leen mal como nuevos.
ESPERA_AUTOCOMPLETADO_S = 7

# El reconocimiento de columnas y los obligatorios se reutilizan del validador.
OBLIGATORIOS = val.OBLIGATORIOS

# ----------------------------- SELECTORES ----------------------------
# Selectores REALES del formulario MAIBA (confirmados por inspeccion del DOM
# vivo). Se usa el sufijo del id/name del subform de participante
# (...curso[participantes][N][datoPersona][campo]) que es estable y NO se
# confunde con los campos del curso. Los desplegables son Select2 (ver
# _seleccionar mas abajo); los uploads son input[type=file] reales.
SEL = {
    # Navegacion: abrir el subform inline (es un <a>, no un <button>)
    "abrir_form":   'text="Cargar Participante"',
    # Campos de texto (por sufijo de id, robusto ante el indice [N])
    "cuil":         'input[id$="datoPersona_cuit"]',          # tiene mascara mask_cuit
    "nombre":       'input[id$="datoPersona_nombre"]',
    "apellido":     'input[id$="datoPersona_apellido"]',
    "dni":          'input[id$="datoPersona_numeroDocumento"]',
    "email":        'input[id$="datoPersona_email"]',
    "celular":      'input[id$="datoPersona_celular"]',
    "domicilio":    'input[id$="datoPersona_domicilio"]',
    "profesion":    'textarea[name$="[ocupacion]"]',          # Profesion / Oficio (obligatorio)
    # Selects Select2 (se operan por JS, ver _seleccionar)
    "tipo_doc":     'select[id$="datoPersona_tipoDocumento"]',
    "partido":      'select[id$="datoPersona_partido"]',
    "localidad":    'select[id$="datoPersona_localidad"]',    # trae todas las localidades (no dependiente)
    # Carga de archivos. OJO: tras el autocompletado aparece un input HIDDEN con
    # el mismo name; por eso filtramos por type=file para no agarrar el hidden.
    "file_foto":    'input[type="file"][name$="[fotoCarnet][archivo]"]',    # Foto carnet <- {DNI}.jpeg
    "file_doc":     'input[type="file"][name$="[constanciaDni][archivo]"]', # Constancia  <- {DNI}.pdf
    # Botones de guardado (dos pasos). "Guardar participante" es un <a>, no un
    # <button>; "Actualizar" si es <button type=submit>.
    "guardar":      'a:has-text("Guardar participante")',        # agrega a la tabla
    "actualizar":   'button:has-text("Actualizar")',             # persiste en la plataforma (verde)
}

VALOR_TIPO_DOC = "DNI"     # opcion fija del select Tipo documento
PROFESION_FALLBACK = "-"   # cuando Profesion viene vacia
REGISTRO = f"registro_aula_{str(AULA).strip()}.xlsx"   # tracker coloreado y reanudable
# =====================================================================


# --------------------------- DATOS DE ENTRADA ------------------------

def construir_personas(ruta_hoja, aula, ruta_carpeta):
    """Lee la hoja + la carpeta y devuelve, para el aula pedida, una lista de
    registros con todos los campos y sus dos archivos. Marca cada uno como
    'ok' (listo para cargar) o no, con la MISMA regla del validador."""
    df = val.leer_hoja(ruta_hoja)
    columnas, faltan = val.resolver_columnas(df)
    if faltan:
        sys.exit("[ERROR] No reconozco columna(s) para: " + ", ".join(faltan))

    col_aula = columnas["aula"]
    mask = (df[col_aula].fillna("").astype(str).str.strip().str.lower()
            == str(aula).strip().lower())
    df_aula = df[mask]
    if df_aula.empty:
        sys.exit(f"[ERROR] No hay filas con Aula = '{aula}'.")

    carpeta = Path(ruta_carpeta)
    if not carpeta.exists():
        sys.exit(f"[ERROR] No encuentro la carpeta del aula: {carpeta.resolve()}")
    pdfs, imgs = val.indexar_archivos(carpeta)

    personas = []
    for _, fila in df_aula.iterrows():
        dni = val.solo_digitos(fila[columnas["dni"]])
        cuil = val.cuil_normalizado(fila[columnas["cuil"]], dni)
        faltantes = [columnas[k] for k in OBLIGATORIOS
                     if k != "cuil" and val.vacio(fila[columnas[k]])]
        tiene_pdf = dni in pdfs
        tiene_img = dni in imgs
        ok = bool(dni) and len(cuil) == 11 and tiene_pdf and tiene_img and not faltantes

        datos = {k: ("" if val.vacio(fila[columnas[k]]) else str(fila[columnas[k]]).strip())
                 for k in columnas if k != "aula"}
        datos["cuil"] = cuil   # CUIL ya normalizado a 11 digitos para tipear
        personas.append({
            "dni": dni,
            "datos": datos,
            "pdf": str((carpeta / pdfs[dni]).resolve()) if tiene_pdf else None,
            "jpeg": str((carpeta / imgs[dni]).resolve()) if tiene_img else None,
            "ok": ok,
            "motivo": "" if ok else "no paso validacion (correr validador_dipa.py)",
        })
    return personas


# ------------------------------ REGISTRO -----------------------------
# El tracker es el .xlsx coloreado (modulo registro): verde = CARGADO,
# rojo = ERROR/FALTA. El cargador lee de ahi quien ya esta en verde para
# reanudar, y pinta cada fila al terminar.

def dnis_ya_cargados(ruta_registro):
    """DNIs ya en VERDE (CARGADO) en el registro xlsx, para reanudar."""
    return registro.leer_cargados(ruta_registro)


def registrar(ruta_registro, dni, etiqueta, resultado, detalle=""):
    registro.actualizar(ruta_registro, dni, resultado, detalle)


# ------------------------------ NAVEGADOR ----------------------------

def abrir_contexto(p):
    """Devuelve (ctx, es_cdp). El script nunca ve usuario ni contrasena.

    CONEXION='cdp': se CONECTA al Chrome que el operador abrio con
    'abrir_chrome_dipa.bat' (depuracion remota). Asi ve TODAS las pestanas
    reales del operador (lo unico que funciona si el login es complejo o si hay
    multiples ventanas de Chrome).

    CONEXION='propio': el script LANZA su propia ventana (channel=chrome)."""
    if CONEXION == "cdp":
        try:
            browser = p.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            sys.exit(
                f"[ERROR] No me pude conectar a tu Chrome por CDP ({CDP_URL}).\n"
                "Primero abri Chrome con depuracion remota: doble clic en\n"
                "   abrir_chrome_dipa.bat\n"
                "logueate, entra al AULA > pestana Participantes, y RECIEN ahi\n"
                "corre el cargador.\n"
                f"Detalle: {type(e).__name__}: {str(e).splitlines()[0]}")
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        es_cdp = True
    else:
        # chromium_sandbox=True evita el cartel "--no-sandbox"; ignore_default_args
        # quita la barra de "controlado por software automatizado".
        opciones = dict(headless=False, chromium_sandbox=True,
                        ignore_default_args=["--enable-automation"])
        try:
            ctx = p.chromium.launch_persistent_context(
                PERFIL_NAVEGADOR, channel="chrome", **opciones)
        except Exception:
            print("[aviso] No pude usar Google Chrome; pruebo el Chromium de Playwright.")
            ctx = p.chromium.launch_persistent_context(PERFIL_NAVEGADOR, **opciones)
        es_cdp = False
    ctx.set_default_timeout(15000)   # que un fallo aparezca en 15s, no que cuelgue 30
    # Aceptar solos los dialogos (confirm/alert) en cualquier pestaña, presente o futura.
    def _enganchar(pg):
        pg.on("dialog", lambda dlg: dlg.accept())
    ctx.on("page", _enganchar)
    for pg in ctx.pages:
        _enganchar(pg)
    return ctx, es_cdp


def _clave_aula(url_aula):
    """Extrae 'editar/<id>' de la URL del aula, para identificar la pestana."""
    import re
    m = re.search(r"editar/(\d+)", url_aula or "")
    return f"editar/{m.group(1)}" if m else None


def _pagina_aula(ctx, url_aula):
    """Devuelve la pestana cuya URL corresponde EXACTAMENTE al aula indicada
    (por su id 'editar/<id>'). None si no esta abierta."""
    clave = _clave_aula(url_aula)
    if not clave:
        return None
    for pg in ctx.pages:
        if pg.is_closed():
            continue
        try:
            if clave in pg.url:
                return pg
        except Exception:
            pass
    return None


def _listar_pestanas(ctx):
    urls = []
    for pg in ctx.pages:
        if pg.is_closed():
            continue
        try:
            urls.append(pg.url)
        except Exception:
            pass
    return urls


# --------------------------- MODO CAPTURAR ---------------------------

def modo_capturar():
    """Vuelca todos los campos del formulario a un archivo, para verificar
    los selectores si la plataforma cambiara algun texto/placeholder."""
    with sync_playwright() as p:
        ctx, page = abrir_contexto(p)
        print("\n>>> Logueate, entra al AULA > pestana 'Participantes' y abri el")
        print("    formulario con 'Cargar Participante'.")
        input(">>> Con el formulario a la vista, apreta Enter... ")

        campos = page.eval_on_selector_all(
            "input, select, textarea, button",
            """els => els.map(e => ({
                tag: e.tagName.toLowerCase(),
                type: e.type || '',
                id: e.id || '',
                name: e.getAttribute('name') || '',
                placeholder: e.getAttribute('placeholder') || '',
                label: (e.labels && e.labels[0] ? e.labels[0].innerText : ''),
                text: (e.innerText || '').trim().slice(0, 40)
            }))"""
        )
        salida = Path("formulario_dipa.txt")
        with salida.open("w", encoding="utf-8") as f:
            for c in campos:
                f.write(f"{c['tag']:8} type={c['type']:10} "
                        f"id={c['id']:25} name={c['name']:25} "
                        f"ph={c['placeholder'][:30]:30} "
                        f"label={c['label'][:25]:25} text={c['text']}\n")
        print(f"\nCampos volcados en: {salida.resolve()}")
        input(">>> Enter para cerrar el navegador... ")
        ctx.close()


# ----------------------------- MODO CARGAR ---------------------------
# El script NO navega: el operador deja la ventana de Chrome en el aula correcta,
# pestana Participantes. Por cada persona: Cargar Participante -> llenar ->
# Guardar participante (queda en la tabla). Recien al final del lote (X personas)
# se aprieta "Actualizar" UNA vez para confirmar todo y minimizar perdidas por lag.
#
# Todos los campos se ubican con .last: en la coleccion Symfony los participantes
# ya guardados quedan como inputs OCULTOS (indices 0..N-1) y el subform recien
# abierto es el ULTIMO (indice mas alto, el unico visible).

def _loc(page, clave):
    return page.locator(SEL[clave]).last


def _valor(page, clave):
    try:
        return (_loc(page, clave).input_value() or "").strip()
    except Exception:
        return ""


def _completar(page, clave, valor):
    if valor:
        _loc(page, clave).fill(valor)


def _escribir_cuil(page, cuil):
    """Escribe el CUIL (solo digitos) caracter por caracter para respetar la
    mascara 'mask_cuit' del campo."""
    loc = _loc(page, "cuil")
    loc.click()
    loc.type(val.solo_digitos(cuil), delay=60)


def _seleccionar(page, clave, texto):
    """Selecciona en el desplegable Select2 (el ULTIMO del DOM = el subform
    abierto) por su TEXTO, ignorando acentos/mayusculas. Setea value y dispara
    'change' (jQuery si esta). Devuelve True si encontro la opcion."""
    if not texto:
        return False
    return page.evaluate(
        """([sel, txt]) => {
            const norm = s => (s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                                .trim().toLowerCase();
            const els = [...document.querySelectorAll(sel)];
            const el = els[els.length - 1];
            if (!el) return false;
            const t = norm(txt);
            let opt = [...el.options].find(o => norm(o.text) === t);
            if (!opt) opt = [...el.options].find(o => norm(o.text).includes(t));
            if (!opt) return false;
            el.value = opt.value;
            if (window.jQuery) window.jQuery(el).trigger('change');
            else el.dispatchEvent(new Event('change', {bubbles: true}));
            return true;
        }""", [SEL[clave], texto])


def _select_sin_elegir(page, clave):
    """True si el Select2 'clave' (el ultimo) no tiene una opcion real elegida
    (esta en '-- Elija ... --' o vacio)."""
    try:
        return bool(page.evaluate(
            """(sel) => {
                const els = [...document.querySelectorAll(sel)];
                const el = els[els.length - 1];
                if (!el) return true;
                const t = (el.options[el.selectedIndex] || {}).text || '';
                return el.selectedIndex <= 0 || /elija/i.test(t);
            }""", SEL[clave]))
    except Exception:
        return True


def _esperar_autocompletado(page, segundos=12):
    """Espera hasta 'segundos' a que la plataforma autocomplete (aparece
    Nombre/Apellido). Devuelve True si autocompleto; False si tras el tiempo
    sigue vacio. Asi evitamos la carrera con el CUIL."""
    import time
    fin = time.time() + segundos
    while time.time() < fin:
        if _autocompletado(page):
            page.wait_for_timeout(600)   # dejar que terminen de cargar los campos
            return True
        page.wait_for_timeout(400)
    return False


def _n_opciones(page, clave):
    try:
        return int(page.evaluate(
            "(sel)=>{const e=[...document.querySelectorAll(sel)].pop(); return e?e.options.length:0;}",
            SEL[clave]))
    except Exception:
        return 0


def _esperar_opciones(page, clave, minimo=2, segundos=10):
    """Espera a que un Select2 tenga sus opciones cargadas (>= minimo)."""
    import time
    fin = time.time() + segundos
    while time.time() < fin:
        if _n_opciones(page, clave) >= minimo:
            return True
        page.wait_for_timeout(400)
    return False


def _autocompletado(page):
    """Tras el CUIL, decide si la plataforma autocompleto la persona:
    el campo Nombre o Apellido ya trae valor."""
    return bool(_valor(page, "nombre") or _valor(page, "apellido"))


def _seleccionar_localidad(page, partido, localidad):
    """Selecciona Partido y Localidad (ambos Select2). Espera a que carguen las
    opciones de Localidad antes de elegir. Fallback: el nombre del partido."""
    if partido and not _seleccionar(page, "partido", partido):
        print(f"  [aviso] no encontre el Partido '{partido}', revisalo a mano.")
    _esperar_opciones(page, "localidad", minimo=2, segundos=10)
    objetivo = localidad or partido
    if objetivo and not _seleccionar(page, "localidad", objetivo):
        if not _seleccionar(page, "localidad", partido):   # fallback
            print(f"  [aviso] no encontre la Localidad '{objetivo}', revisala a mano.")


def _comparar_con_hoja(page, datos):
    """Caso A (autocompletado): NO pisa nada. Compara lo cargado en pantalla
    contra la hoja y devuelve las diferencias para revision."""
    diffs = []
    for clave in ("dni", "nombre", "apellido", "email", "celular", "domicilio"):
        en_pantalla = _valor(page, clave)
        en_hoja = datos.get(clave, "")
        if en_hoja and en_pantalla and en_hoja.lower() != en_pantalla.lower():
            diffs.append(f"{clave}: hoja='{en_hoja}' / DIPA='{en_pantalla}'")
    return diffs


def _completar_profesion(page, texto):
    """Profesion/Oficio es un <textarea readonly 'search-dependent'>: NO acepta
    fill() (da 'element is not editable'). Se setea por JS quitando el readonly
    y disparando los eventos; el valor queda (verificado en vivo)."""
    return page.evaluate(
        """(txt) => {
            const ocs = [...document.querySelectorAll('textarea[name$=\"[ocupacion]\"]')];
            const oc = ocs[ocs.length - 1];
            if (!oc) return false;
            oc.removeAttribute('readonly');
            oc.value = txt;
            ['input','keyup','change','blur'].forEach(ev =>
                oc.dispatchEvent(new Event(ev, {bubbles: true})));
            return true;
        }""", texto)


def _resolver_profesion(page):
    """Si Profesion trae valor, se respeta. Si esta vacia, se pone '-'."""
    if not _valor(page, "profesion"):
        _completar_profesion(page, PROFESION_FALLBACK)


def _quitar_archivo_previo(page, clave):
    """En Caso A puede venir un archivo ya cargado (widget jasny en estado
    'fileinput-exists', con la X roja). Lo quitamos (data-dismiss=fileinput)
    para poder subir el nuevo. Devuelve True si quito algo."""
    try:
        return bool(page.evaluate(
            """(sel) => {
                const ins = [...document.querySelectorAll(sel)];
                const el = ins[ins.length - 1];
                if (!el) return false;
                const cont = el.closest('.fileinput');
                if (!cont) return false;
                const rm = cont.querySelector('[data-dismiss="fileinput"]');
                if (cont.className.includes('fileinput-exists') && rm) {
                    rm.click();
                    return true;
                }
                return false;
            }""", SEL[clave]))
    except Exception:
        return False


def _href_preview(page, clave):
    """Devuelve el href del preview del widget (ruta del archivo) o ''."""
    try:
        return page.evaluate(
            """(sel) => {
                const el = [...document.querySelectorAll(sel)].pop();
                const cont = el && el.closest('.fileinput');
                const a = cont && cont.querySelector('.fileinput-preview a');
                return a ? (a.getAttribute('href') || '') : '';
            }""", SEL[clave]) or ""
    except Exception:
        return ""


def _esperar_subida(page, clave, href_previo, segundos=15):
    """Espera a que la subida AJAX termine: el preview pasa a apuntar a un
    '/uploads/...' NUEVO (distinto del previo). Evita el falso positivo del
    href viejo que jasny deja oculto al quitar la imagen anterior."""
    import time
    fin = time.time() + segundos
    while time.time() < fin:
        h = _href_preview(page, clave)
        if "/uploads/" in h and h != href_previo:
            return True
        page.wait_for_timeout(400)
    return False


def _subir_uno(page, clave, ruta, nombre):
    """Sube UN archivo de forma fiel: abre el dialogo nativo (click en el input)
    y entrega el archivo por el file chooser de Playwright. Eso genera un evento
    'trusted' que dispara la subida AJAX del widget (filestyle+jasny). Si falla,
    cae a set_input_files. Devuelve True si confirmo la subida (/uploads/)."""
    if page.locator(SEL[clave]).count() == 0:
        print(f"       [aviso] no encontre el input de {nombre}; lo salteo.")
        return False
    href_previo = _href_preview(page, clave)
    try:
        with page.expect_file_chooser(timeout=8000) as fc:
            page.evaluate(
                "(sel)=>{const el=[...document.querySelectorAll(sel)].pop(); el.click();}",
                SEL[clave])
        fc.value.set_files(ruta)
    except Exception:
        try:
            page.locator(SEL[clave]).last.set_input_files(ruta, timeout=10000)
        except Exception as e:
            print(f"       [aviso] no pude poner {nombre}: {str(e).splitlines()[0]}")
            return False
    # change explicito por si el widget lo necesita
    try:
        page.evaluate(
            "(sel)=>{const el=[...document.querySelectorAll(sel)].pop();"
            "if(el && window.jQuery) window.jQuery(el).trigger('change');}", SEL[clave])
    except Exception:
        pass
    if _esperar_subida(page, clave, href_previo, segundos=15):
        print(f"       {nombre}: OK (subido al servidor).")
        return True
    print(f"       [aviso] {nombre}: la subida no se confirmo.")
    return False


def _subir_archivos(page, persona):
    """Sube foto carnet (jpeg) y constancia (pdf). Devuelve True si ambas OK."""
    ok_foto = _subir_uno(page, "file_foto", persona["jpeg"], "foto carnet")
    ok_doc = _subir_uno(page, "file_doc", persona["pdf"], "constancia")
    return ok_foto and ok_doc


def _hay_subform(page):
    """True si hay un subform de participante abierto (algun CUIL visible)."""
    return page.evaluate(
        "() => [...document.querySelectorAll('input[id$=\"datoPersona_cuit\"]')]"
        ".some(i => i.offsetParent !== null)")


def cargar_uno(page, persona):
    """Carga UNA persona en la TABLA del curso (Cargar Participante -> llenar ->
    Guardar participante). NO aprieta 'Actualizar': eso se hace una sola vez por
    lote. Devuelve ('EN_TABLA'|'ERROR', notas)."""
    d = persona["datos"]
    etiqueta = f"{d.get('apellido','')} {d.get('nombre','')}".strip()
    print(f"\n--- DNI {persona['dni']}  {etiqueta} ---")

    _loc(page, "abrir_form").click()
    _loc(page, "cuil").wait_for(state="visible", timeout=10000)

    # CUIL primero -> esperar (con sondeo) a ver si autocompleta. Apenas aparece
    # el Nombre -> Caso A (rapido). Si a los X seg sigue vacio, te PREGUNTA (por
    # si todavia estaba cargando): vos confirmas y el script espera tu respuesta.
    _escribir_cuil(page, d.get("cuil", ""))
    print(f"    - esperando autocompletado por CUIL (hasta {ESPERA_AUTOCOMPLETADO_S}s)...")
    autocompleto = _esperar_autocompletado(page, segundos=ESPERA_AUTOCOMPLETADO_S)
    if not autocompleto:
        print(f"    No detecte datos en {ESPERA_AUTOCOMPLETADO_S}s (puede seguir cargando).")
        resp = ""
        while resp not in ("s", "n"):
            resp = input("    ¿Se cargo el form? (aparecieron los datos)  s = si / n = es nuevo: ").strip().lower()
        if resp == "s":
            page.wait_for_timeout(500)
            autocompleto = True

    notas = []
    if autocompleto:
        print("  Caso A: autocompletado (no se pisan los datos de texto).")
        print("    - comparando con la hoja...")
        diffs = _comparar_con_hoja(page, d)
        if diffs:
            print("    DIFERENCIAS para revisar:")
            for x in diffs:
                print("      *", x)
            notas.append("dif: " + " ; ".join(diffs))
        # Partido/Localidad NO siempre se autocompletan (sobre todo Localidad):
        # si quedaron sin elegir, los completamos desde la hoja (son obligatorios).
        if _select_sin_elegir(page, "partido"):
            print("    - Partido vacio -> lo completo desde la hoja...")
            _seleccionar(page, "partido", d.get("partido", ""))
            page.wait_for_timeout(300)
        if _select_sin_elegir(page, "localidad"):
            print("    - Localidad vacia -> la completo desde la hoja...")
            _esperar_opciones(page, "localidad", minimo=2, segundos=10)
            if not _seleccionar(page, "localidad", d.get("localidad", "") or d.get("partido", "")):
                if not _seleccionar(page, "localidad", d.get("partido", "")):
                    print("      [aviso] no encontre la Localidad; revisala a mano.")
                    notas.append("localidad sin completar")
        print("    - resolviendo Profesion...")
        _resolver_profesion(page)
        # Si ya habia archivos cargados (autocompletado), quitarlos (X roja) antes.
        for clave in ("file_foto", "file_doc"):
            if _quitar_archivo_previo(page, clave):
                print(f"    - quite archivo previo de {clave} (X roja).")
                page.wait_for_timeout(300)
        if SUBIR_ARCHIVOS_AUTO:
            print("    - subiendo archivos (auto)...")
            archivos_ok = _subir_archivos(page, persona)
        else:
            archivos_ok = False   # los subis a mano en la pausa
    else:
        print("  Caso B: persona nueva, se completan los obligatorios.")
        if not _seleccionar(page, "tipo_doc", VALOR_TIPO_DOC):
            print("    [aviso] no pude fijar Tipo documento = DNI.")
        _completar(page, "nombre", d.get("nombre", ""))
        _completar(page, "apellido", d.get("apellido", ""))
        _completar(page, "dni", d.get("dni", ""))
        _completar(page, "email", d.get("email", ""))
        _completar(page, "celular", d.get("celular", ""))
        _completar(page, "domicilio", d.get("domicilio", ""))
        print("    - Partido/Localidad...")
        _seleccionar_localidad(page, d.get("partido", ""), d.get("localidad", ""))
        print("    - Profesion...")
        _completar_profesion(page, PROFESION_FALLBACK)
        if SUBIR_ARCHIVOS_AUTO:
            print("    - subiendo archivos (auto)...")
            archivos_ok = _subir_archivos(page, persona)
        else:
            archivos_ok = False   # los subis a mano en la pausa

    # Chequeo previo: avisar si quedaron selects obligatorios sin elegir
    # (la causa tipica de que "Guardar participante" no cierre el subform).
    faltan_sel = [k for k in ("tipo_doc", "partido", "localidad") if _select_sin_elegir(page, k)]
    if faltan_sel:
        print(f"    [aviso] quedan sin elegir: {', '.join(faltan_sel)} (puede fallar el guardado).")
        notas.append("sin elegir: " + ",".join(faltan_sel))

    # Pausa de revision: el subform sigue abierto. Aca subis los archivos a mano
    # (si no fueron por el script) y verificas todo antes de guardar.
    if MODO == "revision":
        if archivos_ok:
            print("  >>> Archivos subidos por el script. Revisa el formulario.")
        else:
            print("  >>> AHORA subi los 2 archivos A MANO: Foto carnet y Constancia.")
            print("  >>> (El subformulario esta abierto en el navegador.)")
        r = input("  >>> Cuando esten subidos: Enter para 'Guardar participante' | 's' SALTEAR: ").strip().lower()
        if r == "s":
            return "SALTEADO", " ; ".join(notas)

    # Guardar participante -> lo agrega a la tabla (NO persiste todavia)
    print("  Guardar participante (a la tabla)...")
    _loc(page, "guardar").click()
    try:
        page.wait_for_function(
            "() => ![...document.querySelectorAll('input[id$=\"datoPersona_cuit\"]')]"
            ".some(i => i.offsetParent !== null)", timeout=8000)
        print("       OK: agregado a la tabla.")
        return "EN_TABLA", " ; ".join(notas)
    except PWTimeout:
        print("       [aviso] el subform no se cerro (¿falto un dato obligatorio?).")
        notas.append("subform no se cerro")
        return "ERROR", " ; ".join(notas)


def _confirmar_curso(page):
    """Tras 'Actualizar' aparece un modal '¿Desea guardar el curso?'. Aprieta
    'Confirmar'. Sondea hasta 12s. Cubre <button>, <a> (rol link) y, por las
    dudas, cualquier button/a con ese texto. El titulo del modal tambien dice
    'Confirmar' pero es un div, asi que no lo toca."""
    import time
    getters = (
        lambda: page.get_by_role("button", name="Confirmar"),
        lambda: page.get_by_role("link", name="Confirmar"),
        lambda: page.locator('a:has-text("Confirmar")'),
        lambda: page.locator('button:has-text("Confirmar")'),
    )
    fin = time.time() + 12
    while time.time() < fin:
        for getter in getters:
            try:
                loc = getter()
                for i in range(min(loc.count(), 6)):
                    el = loc.nth(i)
                    if el.is_visible():
                        el.click()
                        return True
            except Exception:
                pass
        page.wait_for_timeout(400)
    return False


def _persistio(page, dni):
    """Tras Actualizar, verifica que el DNI aparezca en la pagina (tabla)."""
    try:
        return bool(page.evaluate("(d) => document.body.innerText.includes(d)", str(dni)))
    except Exception:
        return False


def modo_cargar(limite=None):
    personas = construir_personas(RUTA_HOJA, AULA, RUTA_CARPETA_AULA)
    listas = [p for p in personas if p["ok"]]
    bloqueadas = [p for p in personas if not p["ok"]]
    ya = dnis_ya_cargados(REGISTRO)
    pendientes = [p for p in listas if p["dni"] not in ya]
    if limite is not None:
        pendientes = pendientes[:limite]

    print("=" * 60)
    print(f"  CARGA AULA '{AULA}'  (modo: {MODO})")
    print("=" * 60)
    print(f"  Total en el aula     : {len(personas)}")
    print(f"  Listas para cargar   : {len(listas)}")
    print(f"  Ya cargadas (saltean): {len([p for p in listas if p['dni'] in ya])}")
    print(f"  A cargar en este lote: {len(pendientes)}" + (f"  (X = {limite})" if limite is not None else ""))
    if bloqueadas:
        print(f"  Bloqueadas (revisar) : {len(bloqueadas)}  -> corre el validador")
    print("=" * 60)
    if not pendientes:
        print("Nada para cargar. Fin.")
        return

    with sync_playwright() as p:
        ctx, es_cdp = abrir_contexto(p)

        if es_cdp:
            print("\n" + "=" * 60)
            print(">>> Conectado a TU Chrome (depuracion remota). Veo tus pestañas reales.")
            print(">>> Logueate (si hace falta) y entra al AULA:")
            print(f">>>   {URL_AULA}")
            print(">>> en la pestana 'Participantes' (que se vea 'Cargar Participante').")
            print("=" * 60)
        else:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            try:
                page.goto(URL_INICIO)
            except Exception:
                pass
            print("\n" + "=" * 60)
            print(">>> USA LA VENTANA QUE ACABA DE ABRIR EL SCRIPT.")
            print(">>> Logueate y entra al AULA en la pestana 'Participantes'.")
            print("=" * 60)
        # Detectar EXACTAMENTE la pestana del aula (por su id editar/<id>).
        while True:
            input(">>> Enter cuando estes en el AULA > Participantes... ")
            pg = _pagina_aula(ctx, URL_AULA)
            if pg is None:
                print(f"[!] No encuentro ninguna pestana en el aula pedida:")
                print(f"    {URL_AULA}")
                print("    Pestañas abiertas ahora mismo (lo que el script esta viendo):")
                for u in _listar_pestanas(ctx):
                    print("      -", u)
                print("    -> Entra a ESA aula EN LA VENTANA DEL SCRIPT y reintenta.")
                continue
            page = pg
            print(f"    Pestana del aula encontrada: {page.url}")
            if page.locator(SEL["abrir_form"]).count() == 0:
                print("    [!] Estoy en el aula pero no veo 'Cargar Participante'.")
                print("        Entra a la pestana 'Participantes' del aula y reintenta.")
                continue
            break
        print(">>> Aula detectada correctamente. Empiezo la carga.")

        en_tabla = []   # (persona, etiqueta, notas) agregados a la tabla, a confirmar
        for persona in pendientes:
            etiqueta = f"{persona['datos'].get('apellido','')} {persona['datos'].get('nombre','')}".strip()
            try:
                resultado, detalle = cargar_uno(page, persona)
            except PWTimeout as e:
                resultado, detalle = "ERROR", f"timeout: {e}"
                print("  [ERROR] timeout:", e)
            except Exception as e:
                resultado, detalle = "ERROR", str(e)
                print("  [ERROR]", e)
            print(f"  -> {resultado}")
            if resultado == "EN_TABLA":
                en_tabla.append((persona, etiqueta, detalle))
            else:
                registrar(REGISTRO, persona["dni"], etiqueta, resultado, detalle)

        # ---- Confirmacion del LOTE: un solo "Actualizar" ----
        if not en_tabla:
            print("\nNadie quedo en la tabla para confirmar.")
        else:
            print(f"\n>>> {len(en_tabla)} participante(s) en la tabla, listos para confirmar.")
            commit = True
            if MODO == "revision":
                resp = input(">>> Enter para apretar 'Actualizar' (confirma el lote) | 'c' cancelar: ")
                commit = resp.strip().lower() != "c"
            if not commit:
                print("Lote CANCELADO: no se apreto Actualizar, nada quedo en verde.")
            else:
                print("  Apretando 'Actualizar'...")
                _loc(page, "actualizar").click()
                page.wait_for_timeout(1200)
                # Modal '¿Desea guardar el curso?' -> Confirmar
                print("  Confirmando ('Confirmar')...")
                if _confirmar_curso(page):
                    print("       confirmado.")
                else:
                    print("       [aviso] no vi el modal 'Confirmar'. Si aparece, confirmalo a mano.")
                try:
                    page.wait_for_load_state("networkidle", timeout=20000)
                except PWTimeout:
                    pass
                page.wait_for_timeout(1500)
                print(f"  URL tras confirmar: {page.url}")
                # Verificar persistencia por DNI; si alguno no aparece, preguntar.
                no_vistos = [t for t in en_tabla if not _persistio(page, t[0]["dni"])]
                confirmar_todos = True
                if no_vistos:
                    print("  [!] No detecto en la pagina a:",
                          ", ".join(f"{t[1]} (DNI {t[0]['dni']})" for t in no_vistos))
                    r = input("  >>> Verifica en el navegador. ¿Se guardaron TODOS? (s/n): ")
                    confirmar_todos = r.strip().lower() == "s"
                for persona, etiqueta, detalle in en_tabla:
                    visto = _persistio(page, persona["dni"])
                    if visto or confirmar_todos:
                        registrar(REGISTRO, persona["dni"], etiqueta, "CARGADO", detalle)
                    else:
                        registrar(REGISTRO, persona["dni"], etiqueta, "ERROR",
                                  (detalle + " ; no aparece tras Actualizar").strip(" ;"))
                print(f"  Registro actualizado: {Path(REGISTRO).resolve()}")

        if es_cdp:
            print(">>> Listo. Tu Chrome queda abierto (no lo cierro).")
        else:
            input(">>> Enter para cerrar el navegador... ")
            ctx.close()


# -------------------------------- MAIN -------------------------------

if __name__ == "__main__":
    modo = sys.argv[1] if len(sys.argv) > 1 else ""
    if modo == "capturar":
        modo_capturar()
    elif modo == "cargar":
        # Cuantas personas: argumento opcional; si no, usa LOTE_DEFECTO del config.
        limite = LOTE_DEFECTO
        if len(sys.argv) > 2:
            arg = sys.argv[2].lower()
            if arg in ("todas", "todos", "all"):
                limite = None
            else:
                try:
                    limite = int(arg)
                except ValueError:
                    sys.exit("[ERROR] El limite debe ser un numero (o 'todas'). Ej: cargar 5")
        modo_cargar(limite)
    else:
        print(__doc__)
        print("Uso:  python cargador_dipa.py capturar")
        print(f"      python cargador_dipa.py cargar          (lote por defecto = {LOTE_DEFECTO})")
        print("      python cargador_dipa.py cargar 8        (lote de 8)")
        print("      python cargador_dipa.py cargar todas    (todas las pendientes)")
