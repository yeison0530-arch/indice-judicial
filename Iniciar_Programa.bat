@echo off
title Generador de Indice Electronico
echo ========================================================
echo   INICIANDO GENERADOR DE INDICE ELECTRONICO
echo ========================================================
echo.

:: Comprobar si Python esta instalado en el sistema
python --version >nul 2>&1
IF %ERRORLEVEL% EQU 0 GOTO CHECK_ENV

echo [AVISO] Python no esta instalado en este equipo.
echo Iniciando descarga e instalacion automatica de Python...
echo Por favor, NO cierres esta ventana.
echo.

echo [Descargando Python 3.11...] Esto puede tardar un poco dependiendo del internet.
curl -# -L -o python_installer.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe

IF NOT EXIST "python_installer.exe" (
    echo [ERROR] No se pudo descargar el instalador de Python.
    echo Por favor, revisa tu conexion a internet o instala Python manualmente.
    pause
    exit /b
)

echo.
echo [Instalando Python de forma silenciosa...]
echo [NOTA: Si Windows te pide permisos de administrador, haz clic en SI]
    
:: Ejecutar la instalacion silenciosa
start "" /wait "python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_doc=0
    
:: Borrar el instalador una vez terminado
del "python_installer.exe"
    
:: Actualizar la variable PATH en esta ventana para que reconozca el comando "python" de inmediato
set PATH=%LocalAppData%\Programs\Python\Python311\Scripts\;%LocalAppData%\Programs\Python\Python311\;%PATH%
    
:: Volver a comprobar si la instalacion tuvo exito
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR CRITICO] La instalacion automatica de Python fallo.
    echo Es posible que el antivirus lo haya bloqueado o necesites reiniciar la computadora.
    pause
    exit /b
)
echo.
echo [EXITO] Python se instalo correctamente de forma automatica.
echo.

:CHECK_ENV
:: Comprobar si existe la carpeta del entorno virtual "entorno_app"
IF EXIST "entorno_app\Scripts\activate.bat" GOTO START_APP

echo [1/3] Creando entorno de trabajo local por primera vez...
python -m venv entorno_app
    
echo [2/3] Instalando componentes necesarios (esto puede tardar unos minutos)...
call entorno_app\Scripts\activate.bat
python -m pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt
echo.
echo [3/3] Instalacion de componentes completada exitosamente.
GOTO RUN_STREAMLIT

:START_APP
echo Iniciando entorno de la aplicacion...
call entorno_app\Scripts\activate.bat

:RUN_STREAMLIT
echo.
echo ========================================================
echo   ABRIENDO LA APLICACION EN TU NAVEGADOR...
echo   (No cierres esta ventana negra mientras la usas)
echo ========================================================
echo.

:: Omitir el mensaje de correo electronico de Streamlit
mkdir "%USERPROFILE%\.streamlit" >nul 2>&1
echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"

:: Ejecutar la aplicacion de Streamlit
streamlit run app.py

pause
