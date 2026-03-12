@echo off
echo Select version:
echo   1. Local OCR  (no API key required)
echo   2. Claude API (requires Anthropic API key)
set /p choice="Enter 1 or 2: "
if "%choice%"=="2" (
    python "%~dp0table_converter.py"
) else (
    python "%~dp0table_converter_local.py"
)
pause
