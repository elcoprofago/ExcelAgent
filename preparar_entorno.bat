@echo off
title ExcelAgent - preparar entorno
rem Deja andando ..\.venv con el Python de esta PC (ver preparar_entorno.ps1).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0preparar_entorno.ps1"
