@echo off
chcp 65001 >nul

pwsh -ExecutionPolicy Bypass -File "%~dp0build.ps1"

echo.
echo Appuyez sur une touche pour fermer...
pause >nul