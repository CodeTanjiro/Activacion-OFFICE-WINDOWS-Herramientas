@echo off
title Activación y Herramientas
color 0A
mode con: cols=100 lines=30

:login
cls
echo Ingrese sus credenciales:
set /p username="Usuario: "
set /p password="Contraseña: "

rem Verifica las credenciales ingresadas
if "%username%"=="admin" (
    if "%password%"=="admin" (
        goto menu
    ) else (
        echo Credenciales incorrectas. Inténtelo de nuevo.
        pause
        goto login
    )
) else (
    echo Credenciales incorrectas. Inténtelo de nuevo.
    pause
    goto login
)

:menu
cls
echo  @ Creador: CodeTanjiro
echo.  Versión 0.1.2024
echo.
echo   ACTIVACIÓN Y HERRAMIENTAS  
echo.
echo         Microsoft 365  ---  Office
echo                      MICROSOFT 365 -------- ACTIVACIÓN 
echo 1. Act. Método 1 (Licencia en Volumen) --- Microsoft
echo 2. Act. Método 2 (Licencia en Volumen) --- Externo
echo 3. Act. Método 3 (Reactivación) --- Microsoft
echo.
echo                             WINDOWS -------- ACTIVACIÓN 
echo 4. Activación de Windows 10 (Licencia en Volumen) --- Microsoft
echo.
echo                                       HERRAMIENTAS 
echo 5. Eliminar clave de producto de Office
echo 6. Verificar activación de Windows y Office
echo 7. Convertir archivos a una extensión específica
echo 8. Verificar puertos abiertos
echo 9. Ver mi dirección IP
echo 10. Guardar especificaciones de la PC
echo 11. Salir
echo.

set /p choice="Elige Tipo de Activación: "
if "%choice%"=="1" goto script1
if "%choice%"=="2" goto script2
if "%choice%"=="3" goto script3
if "%choice%"=="4" goto script4
if "%choice%"=="5" goto script5
if "%choice%"=="6" goto script6
if "%choice%"=="7" goto script7
if "%choice%"=="8" goto script8
if "%choice%"=="9" goto script9
if "%choice%"=="10" goto script10
if "%choice%"=="11" exit
goto menu

:script1
rem Activación de Microsoft Office usando licencia en volumen (Método 1)
rem Reemplazar la clave de producto por la correcta antes de ejecutar
cd /d %ProgramFiles%\Microsoft Office\Office16
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
)
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:e8.us.to
cscript ospp.vbs /act
pause
goto menu

:script2
rem Act. Método 2 (Licencia en Volumen) --- Externo
powershell -NoExit -Command "irm https://massgrave.dev/get | iex"
pause
goto menu

:script3
rem Reactivación de Microsoft Office
cd /d %ProgramFiles%\Microsoft Office\Office16
cscript ospp.vbs /act
pause
goto menu

:script4
rem Activación de Windows 10 usando licencia en volumen
rem Proporcionar opciones para diferentes ediciones de Windows 10
cls
echo Seleccione el tipo de activación de Windows 10:
echo.
echo 1. Windows 10 Home
echo 2. Windows 10 Home N
echo 3. Windows 10 Professional
echo 4. Windows 10 Professional N
echo 5. Windows 10 Enterprise
echo 6. Windows 10 Enterprise N
echo 7. Windows 10 Education
echo 8. Windows 10 Education N
echo.
set /p opcion="Ingrese el número correspondiente al tipo de activación: "

if "%opcion%" geq "1" if "%opcion%" leq "8" (
    for /f "tokens=%opcion%" %%a in (
        "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
        "3KHY7-WNT83-DGQKR-F7HPR-844BM"
        "W269N-WFGWX-YVC9B-4J6C9-T83GX"
        "MH37W-N47XK-V7XM9-C7227-GCQG9"
        "NPPR9-FWDCX-D2C8J-H872K-2YT43"
        "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
        "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
        "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
    ) do slmgr /ipk %%a
) else (
    echo Opción no válida. Intente nuevamente.
    pause
    goto script4
)

echo.
echo Presione una tecla para continuar...
pause

slmgr /skms kms.digiboy.ir
echo.
echo Espere hasta que acepte la ventana emergente...
pause

slmgr /ato
echo.
echo Espere hasta que acepte la ventana emergente...
pause
goto menu

:script5
rem Eliminar clave de producto de Office
cd /d %ProgramFiles%\Microsoft Office\Office16
cscript ospp.vbs /unpkey:6MWKP >nul
echo La clave de producto de Office se ha eliminado correctamente.
pause
goto menu

:script6
rem Verificar activación de Windows y Office
cls
echo Verificando la activación de Windows y Office...
echo.
cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus
echo.
echo.
slmgr /xpr
echo.
pause
goto menu

:script7
rem Convertir archivos en una carpeta a una extensión específica
cls
echo Convertir archivos en una carpeta a una extensión específica
echo.

set /p "carpeta=Ruta de la carpeta con archivos: "
set /p "extension=Extensión de destino (sin punto): "

echo.

if not exist "%carpeta%" (
    echo Carpeta no existe.
    pause
    goto menu
)

dir /b "%carpeta%\*" > nul 2>&1
if errorlevel 1 (
    echo Carpeta vacía.
    pause
    goto menu
)

echo Convirtiendo archivos en la carpeta "%carpeta%" a la extensión .%extension%...

for %%i in ("%carpeta%\*.*") do (
    set "archivo=%%~ni"
    set "extension_actual=%%~xi"
    
    if /i not "%extension_actual:~1%"=="%extension%" (
        echo Convirtiendo "%%i"...
        type "%%i" >> "%%~dpni.%extension%"
    ) else (
        echo El archivo "%%i" ya tiene la extensión .%extension%.
    )
)

echo.
echo Conversión completada.
pause
goto menu

:script8
rem Verificar puertos abiertos y guardar en un archivo
@echo off

netstat -ano | findstr /c:"ESTABLISHED" > temp_hosts_abiertos.txt

for /f "tokens=2,3,5 delims= " %%a in (temp_hosts_abiertos.txt) do (
    echo Dirección Remota: %%a, Puerto Remoto: %%b >> hosts_abiertos.txt
)

del temp_hosts_abiertos.txt

echo Info. sobre hosts abiertos en hosts_abiertos.txt
pause
goto menu

:script9
rem Ver mi dirección IP y guardar en un archivo
ipconfig /all > direccion_ip.txt
echo Dirección IP detallada en direccion_ip.txt
pause
goto menu

:script10
rem Guardar especificaciones de la PC en un archivo
cls
echo Guardando especificaciones de la PC...
echo.
systeminfo > "Especificaciones_PC.txt"
echo Especificaciones guardadas en 'Especificaciones_PC.txt'.
echo.
pause
goto menu
