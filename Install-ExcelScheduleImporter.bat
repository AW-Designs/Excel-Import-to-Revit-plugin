@echo off
REM Excel Schedule Importer - one-click installer
REM ------------------------------------------------
REM Double-click this file. It detects your installed Revit versions
REM (2024/2025/2026), downloads the matching add-in from the latest GitHub
REM release, and registers it under your Windows user profile.
REM No admin rights needed. Safe to run again later (it just re-installs
REM the current release), though the add-in also updates itself automatically.

title Excel Schedule Importer - Installer

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; iex (irm 'https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest/download/install.ps1')"

echo.
echo ============================================================
echo  Done. Start Revit - look for "Import Excel Schedule" on the
echo  Add-Ins tab. First launch may ask about an unsigned add-in -
echo  click "Always Load".
echo ============================================================
echo.
pause
