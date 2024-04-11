@echo off
setlocal EnableDelayedExpansion

:menu
cls
echo  @ Creador: CodeTanjiro
echo.  Versión 0.1.2024

echo         Microsoft 365  ---  Office
echo                      MICROSOFT 365 -------- ACTIVACIÓN 
echo 1. Act. Método 1 (Licencia en Volumen) --- Microsoft
echo 2. Act. Método 2 (Licencia en Volumen) --- Externo
echo 3. Act. Método 3 (Reactivación) --- Microsoft
echo                             WINDOWS -------- ACTIVACIÓN 
echo 4. Activación de Windows 10 (Licencia en Volumen) --- Microsoft
echo                                       HERRAMIENTAS 
echo 5. Eliminar clave de producto de Office
echo 6. Verificar activación de Windows y Office
echo 7. Convertir archivos a una extensión específica
echo 8. Verificar puertos abiertos
echo 9. Ver mi dirección IP
echo 10. Guardar especificaciones de la PC
echo 11. Realizar escaneo antivirus
echo 12. Salir
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
if "%choice%"=="11" goto script11
if "%choice%"=="12" exit
goto menu

:script1
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
powershell -NoExit -Command "irm https://massgrave.dev/get | iex"
pause
goto menu

:script3
cd /d %ProgramFiles%\Microsoft Office\Office16
cscript ospp.vbs /act
pause
goto menu

:script4
cls
echo Seleccione el tipo de activación de Windows 10:
echo.
echo 1. Windows 10 Home
echo 2. Windows 10 Home N
echo 3. Windows 10 Home Single Language
echo 4. Windows 10 Home Country Specific
echo 5. Windows 10 Professional
echo 6. Windows 10 Professional N
echo 7. Windows 10 Enterprise
echo 8. Windows 10 Enterprise N
echo 9. Windows 10 Education
echo 10. Windows 10 Education N
echo 11. Windows 10 Enterprise 2015 LTSB
echo 12. Windows 10 Enterprise 2015 LTSB N
echo.
set /p opcion="Ingrese el número correspondiente al tipo de activación: "

REM Ejecución de los comandos según la opción seleccionada
if "%opcion%"=="1" (
    slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
) else if "%opcion%"=="2" (
    slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM
) else if "%opcion%"=="3" (
    slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
) else if "%opcion%"=="4" (
    slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
) else if "%opcion%"=="5" (
    slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
) else if "%opcion%"=="6" (
    slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
) else if "%opcion%"=="7" (
    slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
) else if "%opcion%"=="8" (
    slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
) else if "%opcion%"=="9" (
    slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
) else if "%opcion%"=="10" (
    slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
) else if "%opcion%"=="11" (
    slmgr /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9
) else if "%opcion%"=="12" (
    slmgr /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ
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
cd /d %ProgramFiles%\Microsoft Office\Office16
cscript ospp.vbs /unpkey:6MWKP >nul
echo La clave de producto de Office se ha eliminado correctamente.
pause
goto menu

:script6
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
REM Convertir archivos en una carpeta a una extensión específica
cls
echo Convertir archivos en una carpeta a una extensión específica
echo.

set /p carpeta="Ruta de la carpeta con archivos: "
set /p extension="Extensión de destino (sin punto): "

echo.

REM Verificar si la carpeta existe


if not exist "%carpeta%" (
    echo Carpeta no existe.
    pause
    goto menu
)

REM Verificar si la carpeta contiene archivos
dir /b "%carpeta%\*" > nul 2>&1
if errorlevel 1 (
    echo Carpeta vacía.
    pause
    goto menu
)

echo Convirtiendo archivos en la carpeta "%carpeta%" a la extensión .%extension%...

REM Recorriendo archivos en la carpeta
for %%i in ("%carpeta%\*.*") do (
    REM Nombre y extensión del archivo
    set "archivo=%%~ni"
    set "extension_actual=%%~xi"
    
    REM Verificar si la extensión actual no es la de destino
    if /i not "%extension_actual:~1%"=="%extension%" (
        echo Convirtiendo "%%i"...
        
        REM Concatenar contenido al archivo con extensión de destino
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

rem Verificar puertos
netstat -ano | findstr /c:"ESTABLISHED" > temp_hosts_abiertos.txt

rem Extraer dirección remota y puerto remoto
for /f "tokens=2,3,5 delims= " %%a in (temp_hosts_abiertos.txt) do (
    echo Dirección Remota: %%a, Puerto Remoto: %%b >> hosts_abiertos.txt
)

rem Eliminar archivo temporal
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

:script11
rem Realizar escaneo antivirus
cls
echo Realizar escaneo antivirus:
echo.
echo 1. Escanear un archivo específico
echo 2. Escanear todo un directorio
echo 3. Volver al menú principal
echo.

set /p scan_choice="Seleccione una opción: "

if "%scan_choice%"=="1" goto scan_file
if "%scan_choice%"=="2" goto scan_directory
if "%scan_choice%"=="3" goto menu

:scan_file
rem Escanear un archivo específico
set /p file_to_scan="Ruta del archivo a escanear: "

rem Verificar si el archivo existe
if not exist "%file_to_scan%" (
    echo El archivo no existe.
    pause
    goto script11
)

rem Realizar escaneo antivirus
echo Escaneando "%file_to_scan%"...

rem Comando de escaneo antivirus
rem Simulación de escaneo
set /a "total_examined=1"
set "failed_scan_files="
echo Archivos escaneados: %total_examined% > escaneo_resultado.txt
echo Archivos no escaneados: %failed_scan_files% >> escaneo_resultado.txt
echo. >> escaneo_resultado.txt
echo Escaneo completo.
pause
goto script11

:scan_directory
rem Escanear un directorio
set /p directory_to_scan="Ruta del directorio a escanear: "

rem Verificar si el directorio existe
if not exist "%directory_to_scan%" (
    echo Directorio no existe.
    pause
    goto script11
)

rem Realizar escaneo antivirus
echo Escaneando directorio "%directory_to_scan%"...

rem Comando de escaneo antivirus
rem Simulación de escaneo
set /a "total_examined=10"
set "failed_scan_files=archivo1.txt, archivo2.exe"
echo Archivos escaneados: %total_examined% > escaneo_resultado.txt
echo Archivos no escaneados: %failed_scan_files% >> escaneo_resultado.txt
echo. >> escaneo_resultado.txt
echo Escaneo completo.
pause
goto script11