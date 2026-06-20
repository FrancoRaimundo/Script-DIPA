@echo off
REM Abre UNA sola ventana de Google Chrome, controlable por el cargador (via
REM depuracion remota en el puerto 9222), apuntando al login de DIPA.
REM
REM Cierra TODO Chrome primero para que no haya ventanas "sueltas" en las que
REM uno navega por error y el script no ve. Esa era la causa de "no veo el aula".
REM
REM Uso: doble clic. Logueate y entra al AULA > pestana Participantes.
REM Luego, en la terminal:  python cargador_dipa.py cargar 1

setlocal
set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" (
  echo No encontre chrome.exe en las rutas habituales.
  echo Edita este .bat y pone la ruta correcta de Chrome.
  pause
  exit /b 1
)

echo ============================================================
echo  ATENCION: se van a CERRAR todas las ventanas de Chrome.
echo  Guarda tu trabajo (otras pestanas) antes de continuar.
echo ============================================================
pause

echo Cerrando Chrome...
taskkill /F /IM chrome.exe >nul 2>&1
REM Esperar a que liberen el perfil
ping -n 3 127.0.0.1 >nul

echo Abriendo la UNICA ventana de Chrome controlable (puerto 9222)...
start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%~dp0.perfil_dipa_cdp" "https://plataforma.maa.gba.gov.ar/logindpsit"

echo.
echo Listo. En ESA ventana (la unica abierta):
echo   1) Logueate (toda la cadena SSO).
echo   2) Entra al AULA y a la pestana 'Participantes'.
echo   3) En la terminal corre:  python cargador_dipa.py cargar 1
echo.
endlocal
