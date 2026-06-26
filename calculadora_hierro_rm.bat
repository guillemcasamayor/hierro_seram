@echo off
title Calculadora de Hierro Hepatico en RM - SEDIA
chcp 65001 >nul
setlocal enabledelayedexpansion

:MENU
cls
echo =======================================================================
echo     CALCULADORA DE HIERRO HEPÁTICO EN RM (Modelo Osatek-SEDIA)
echo =======================================================================
echo  [1] Realizar cálculo y generar informe clínico
echo  [2] Ver escala de valoración e información técnica de adquisición
echo  [3] Salir
echo =======================================================================
set /p opt="Seleccione una opción [1-3]: "

if "%opt%"=="1" goto CALC
if "%opt%"=="2" goto INFO
if "%opt%"=="3" goto EXIT
goto MENU

:CALC
cls
echo =======================================================================
echo          INGRESO DE DATOS - CÁLCULO DE HIERRO HEPÁTICO (CHH)
echo =======================================================================
echo Ingrese los datos del paciente y las ROIs (use punto o coma para decimales).
echo.

set "PAT_NAME="
set /p PAT_NAME="Nombre completo del paciente: "
if "%PAT_NAME%"=="" set PAT_NAME=Paciente de Prueba

set "PAT_ID="
set /p PAT_ID="Historia Clínica / ID: "
if "%PAT_ID%"=="" set PAT_ID=NHC-SINDATO

set "PAT_AGE="
set /p PAT_AGE="Edad del paciente en años (opcional): "

echo.
echo === SECUENCIA DP-HEMOCROM ===
set /p H_DP_1="Hígado Medida 1 (DP-1): "
set /p H_DP_2="Hígado Medida 2 (DP-2): "
set /p H_DP_3="Hígado Medida 3 (DP-3): "
set /p M_DP_1="Ms paravertebral Medida 1 (DP-M1): "
set /p M_DP_2="Ms paravertebral Medida 2 (DP-M2): "

echo.
echo === SECUENCIA T2-HEMOCROM ===
set /p H_T2_1="Hígado Medida 1 (T2-1): "
set /p H_T2_2="Hígado Medida 2 (T2-2): "
set /p H_T2_3="Hígado Medida 3 (T2-3): "
set /p M_T2_1="Ms paravertebral Medida 1 (T2-M1): "
set /p M_T2_2="Ms paravertebral Medida 2 (T2-M2): "

rem Validar que las variables no estén vacías
if "%H_DP_1%"=="" goto ERR_EMPTY
if "%H_DP_2%"=="" goto ERR_EMPTY
if "%H_DP_3%"=="" goto ERR_EMPTY
if "%M_DP_1%"=="" goto ERR_EMPTY
if "%M_DP_2%"=="" goto ERR_EMPTY
if "%H_T2_1%"=="" goto ERR_EMPTY
if "%H_T2_2%"=="" goto ERR_EMPTY
if "%H_T2_3%"=="" goto ERR_EMPTY
if "%M_T2_1%"=="" goto ERR_EMPTY
if "%M_T2_2%"=="" goto ERR_EMPTY

cls
echo Realizando cálculos...
echo.

rem Crear script de PowerShell temporal
set "PS_FILE=%TEMP%\mri_calc_temp.ps1"
(
echo $ErrorActionPreference = "Stop"
echo try {
echo     $h_dp_1 = [double]($env:H_DP_1 -replace ',', '.')
echo     $h_dp_2 = [double]($env:H_DP_2 -replace ',', '.')
echo     $h_dp_3 = [double]($env:H_DP_3 -replace ',', '.')
echo     $m_dp_1 = [double]($env:M_DP_1 -replace ',', '.')
echo     $m_dp_2 = [double]($env:M_DP_2 -replace ',', '.')
echo.
echo     $h_t2_1 = [double]($env:H_T2_1 -replace ',', '.')
echo     $h_t2_2 = [double]($env:H_T2_2 -replace ',', '.')
echo     $h_t2_3 = [double]($env:H_T2_3 -replace ',', '.')
echo     $m_t2_1 = [double]($env:M_T2_1 -replace ',', '.')
echo     $m_t2_2 = [double]($env:M_T2_2 -replace ',', '.')
echo.
echo     $avg_h_dp = ($h_dp_1 + $h_dp_2 + $h_dp_3) / 3
echo     $avg_m_dp = ($m_dp_1 + $m_dp_2) / 2
echo     $ratio_dp = $avg_h_dp / $avg_m_dp
echo.
echo     $avg_h_t2 = ($h_t2_1 + $h_t2_2 + $h_t2_3) / 3
echo     $avg_m_t2 = ($m_t2_1 + $m_t2_2) / 2
echo     $ratio_t2 = $avg_h_t2 / $avg_m_t2
echo.
echo     # Fórmula de Osatek-SEDIA
echo     $q_rmn = [Math]::Exp^(5.808 - ^(1.518 * $ratio_dp^) - ^(0.877 * $ratio_t2^)^)
echo     $q_rmn_mg = $q_rmn / 17.8
echo.
echo     $h_dp_str = "$h_dp_1 $h_dp_2 $h_dp_3"
echo     $m_dp_str = "$m_dp_1 $m_dp_2"
echo     $h_t2_str = "$h_t2_1 $h_t2_2 $h_t2_3"
echo     $m_t2_str = "$m_t2_1 $m_t2_2"
echo.
echo     $normalLine = "No sobrecarga. Descarta sobrecarga (VPN=100%)."
echo     $mildLine   = "Sobrecarga leve o no sobrecarga. Descarta alta sobrecarga (VPN=100%, 99% <50 umol Fe/g en la biopsia)."
echo     $moderateLine = "Sobrecarga moderada. No descarta alta sobrecarga (87% <70 umol Fe/g en la biopsia)."
echo     $severeLine = "Alta sobrecarga. Si el resultado es >85 confirma alta sobrecarga (VPP=100%)."
echo.
echo     # Determinar conclusión resaltada
echo     $val = $q_rmn
echo     $c_normal = if ($val -lt 20) { "[X] $normalLine" } else { "    $normalLine" }
echo     $c_mild   = if ($val -ge 20 -and $val -lt 40) { "[X] $mildLine" } else { "    $mildLine" }
echo     $c_mod    = if ($val -ge 40 -and $val -lt 80) { "[X] $moderateLine" } else { "    $moderateLine" }
echo     $c_sev    = if ($val -ge 80) { "[X] $severeLine" } else { "    $severeLine" }
echo.
echo     $pat_name = $env:PAT_NAME
echo     $pat_id   = $env:PAT_ID
echo     $age_str  = if ($env:PAT_AGE) { " | EDAD: $env:PAT_AGE años" } else { "" }
echo.
echo     # Nota adicional de Hemocromatosis Hereditaria
echo     $hhNote = ""
echo     if ($env:PAT_AGE) {
echo         $age = [int]($env:PAT_AGE -replace '[^0-9]', '')
echo         if ($age -gt 0 -and $val -ge (2 * $age)) {
echo             $hhNote = "`n* Nota: El valor de CHH supera el doble de la edad ($($age * 2) umol/g),`n  altamente compatible con Hemocromatosis Hereditaria (IHH > 1.9)."
echo         }
echo     }
echo.
echo     $report = @"
Informe:
PACIENTE: $pat_name ($pat_id)$age_str
TÉCNICA: Se realizan secuencias T2 y DP según protocolo SEDIA.

HALLAZGOS:
DP-Hemocrom
Hígado $h_dp_str
Ms paravertebral $m_dp_str
Ratio $($ratio_dp.ToString("F2"))

T2-Hemocrom
Hígado $h_t2_str
Ms paravertebral $m_t2_str
Ratio $($ratio_t2.ToString("F2"))

QRM
- $($q_rmn.ToString("F2")) umol Fe/g
- $($q_rmn_mg.ToString("F2")) mg Fe/g

Conclusión:
$c_normal
$c_mild
$c_mod
$c_sev$hhNote
"@
echo.
echo     Write-Host "============================================================" -ForegroundColor Cyan
echo     Write-Host $report
echo     Write-Host "============================================================" -ForegroundColor Cyan
echo.
echo     # Guardar en archivo local
echo     $clean_id = $pat_id -replace '[^a-zA-Z0-9_-]', ''
echo     if ($clean_id -eq "") { $clean_id = "hierro" }
echo     $filename = "informe_" + $clean_id + ".txt"
echo     $report | Out-File -FilePath $filename -Encoding utf8
echo     Write-Host "Informe guardado automáticamente en: $filename" -ForegroundColor Yellow
echo     Write-Host "============================================================" -ForegroundColor Cyan
echo } catch {
echo     Write-Host "ERROR: Asegúrese de ingresar solo valores numéricos válidos." -ForegroundColor Red
echo     Write-Host $_.Exception.Message -ForegroundColor Red
echo }
) > "%PS_FILE%"

rem Ejecutar PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

rem Eliminar archivo temporal
if exist "%PS_FILE%" del "%PS_FILE%"

echo.
pause
goto MENU

:ERR_EMPTY
echo.
echo [ERROR] Todos los campos de DP-Hemocrom y T2-Hemocrom son requeridos.
echo Presione cualquier tecla para volver a intentar...
pause >nul
goto CALC

:INFO
cls
echo =======================================================================
echo          TABLA DE VALORACIÓN E INFORMACIÓN DE ADQUISICIÓN
echo =======================================================================
echo.
echo 1. CLASIFICACIÓN DE SOBRECARGA FÉRRICA (OSATEK-SEDIA):
echo -----------------------------------------------------------------------
echo  Rango (umol Fe/g)  ^|  Valoración          ^| Interpretación
echo -----------------------------------------------------------------------
echo  ^< 20 umol         ^|  No Sobrecarga        ^| Descarta sobrecarga (VPN=100%)
echo  20 - 39 umol      ^|  Leve / No Sobrecarga ^| Descarta alta sobrecarga (VPN=100%)
echo  40 - 79 umol      ^|  Sobrecarga Moderada  ^| No descarta alta sobrecarga
echo  ^>= 80 umol         ^|  Alta Sobrecarga      ^| Si es ^>85 confirma alta (VPP=100%)
echo -----------------------------------------------------------------------
echo  * Valores normales de CHH: ^< 36 umol Fe/g.
echo  * Conversión de unidades: mg Fe/g = umol Fe/g / 17.8
echo.
echo 2. REQUISITOS CLAVE DE ADQUISICIÓN EN RESONANCIA:
echo -----------------------------------------------------------------------
echo  * Equipo: Válido exclusivamente para resonancias de 1.5 Teslas.
echo  * Antena: NUNCA usar antena de superficie activa. Se debe adquirir
echo    siempre utilizando la antena interna del gantry ("Q BODY").
echo  * Saturación: El método pierde precisión cuantitativa en rangos
echo    superiores a 250 umol Fe/g debido a la saturación de señal.
echo.
echo =======================================================================
echo Presione cualquier tecla para regresar al menú...
pause >nul
goto MENU

:EXIT
cls
echo Gracias por usar el Calculador de Hierro Hepático local.
echo.
exit /b
