@echo off
chcp 65001 >nul
title रजत पीडीएफ कन्वर्टर (Rajat PDF Converter Pro)
color 0B
echo ====================================================================
echo             रजत पीडीएफ कन्वर्टर (Rajat PDF Converter Pro)
echo ====================================================================
echo.
echo सॉफ्टवेयर लोड हो रहा है, कृपया प्रतीक्षा करें...
echo.

cd /d "%~dp0"

where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    py -3.9 run.py
    if %ERRORLEVEL% neq 0 (
        py run.py
    )
    goto end
)

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    python run.py
    goto end
)

echo [त्रुटि] Python नहीं मिला! कृपया Python 3.9+ स्थापित करें।
pause

:end
