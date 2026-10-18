@echo off
chcp 65001 >nul
echo ======================================================================
echo   DANG CAP NHAT VA KICH HOAT COPY CHEMDRAW SANG WORD (BAN V5)...
echo ======================================================================

set "SUITE_DIR=%USERPROFILE%\Downloads\ONLYOFFICE-Open-Suite"
if exist "%USERPROFILE%\Desktop\ONLYOFFICE-Open-Suite\Cai-Dat-Len-May-Moi.cmd" set "SUITE_DIR=%USERPROFILE%\Desktop\ONLYOFFICE-Open-Suite"
if exist "%USERPROFILE%\Downloads\ONLYOFFICE-Open-Suite\ONLYOFFICE-Open-Suite\Cai-Dat-Len-May-Moi.cmd" set "SUITE_DIR=%USERPROFILE%\Downloads\ONLYOFFICE-Open-Suite\ONLYOFFICE-Open-Suite"

if not exist "%SUITE_DIR%\portable" mkdir "%SUITE_DIR%\portable"
if not exist "%SUITE_DIR%\icons" mkdir "%SUITE_DIR%\icons"

echo [1/3] Dang tai ban cap nhat moi nhat (khong bi trang hinh)...
set "TMP_ZIP=%TEMP%\oo_ota_update.zip"
set "TMP_EXT=%TEMP%\oo_ota_extract"
if exist "%TMP_EXT%" rmdir /s /q "%TMP_EXT%" >nul 2>&1
mkdir "%TMP_EXT%" >nul 2>&1

curl.exe -sL -H "Cache-Control: no-cache" "https://raw.githubusercontent.com/thanhAI408/onlyoffice-ota-update/main/Gui-Cho-May-Khac-Sieu-Nhe.zip?v=%RANDOM%%RANDOM%" -o "%TMP_ZIP%"
tar.exe -xf "%TMP_ZIP%" -C "%TMP_EXT%" >nul 2>&1

if exist "%TMP_EXT%\ONLYOFFICE-Open-Suite\portable\OnlyOfficeLauncher.cs" (
    xcopy /s /e /y /q "%TMP_EXT%\ONLYOFFICE-Open-Suite\*" "%SUITE_DIR%\" >nul 2>&1
)
curl.exe -sL -H "Cache-Control: no-cache" "https://raw.githubusercontent.com/thanhAI408/onlyoffice-ota-update/main/OnlyOfficeLauncher.cs?v=%RANDOM%%RANDOM%" -o "%SUITE_DIR%\portable\OnlyOfficeLauncher.cs"

echo [2/3] Dang bien dich va kich hoat cau noi ChemDraw 24/7...
taskkill /f /im OnlyOfficeLauncher.exe >nul 2>&1

set "CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if exist "%SUITE_DIR%\icons\OfficeSuite.ico" (
    "%CSC%" /nologo /target:winexe /optimize+ /out:"%SUITE_DIR%\portable\OnlyOfficeLauncher.exe" /win32icon:"%SUITE_DIR%\icons\OfficeSuite.ico" /r:System.dll,System.Drawing.dll,System.Windows.Forms.dll,System.IO.Compression.FileSystem.dll "%SUITE_DIR%\portable\OnlyOfficeLauncher.cs" >nul
) else (
    "%CSC%" /nologo /target:winexe /optimize+ /out:"%SUITE_DIR%\portable\OnlyOfficeLauncher.exe" /r:System.dll,System.Drawing.dll,System.Windows.Forms.dll,System.IO.Compression.FileSystem.dll "%SUITE_DIR%\portable\OnlyOfficeLauncher.cs" >nul
)

echo [3/3] Dang khoi dong dich vu ngam...
"%SUITE_DIR%\portable\OnlyOfficeLauncher.exe" --install

rmdir /s /q "%TMP_EXT%" >nul 2>&1
del /f /q "%TMP_ZIP%" >nul 2>&1

echo.
echo ======================================================================
echo   HOAN TAT 100%! DA SUA LOI KHUNG TRANG KHI DAN CHEMDRAW!
echo   Ban co the dan (Ctrl+V) vao Word NGAY BAY GIO!
echo ======================================================================
timeout /t 3
