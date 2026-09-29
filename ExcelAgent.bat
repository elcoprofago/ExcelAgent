@echo off
title ExcelAgent
rem Rutas relativas a este archivo: el entorno virtual esta en la carpeta de arriba (..\.venv).
rem Asi funciona igual en D:\REPOS\EXCEL (notebook) que en F:\source\repos\EXCEL.
if not exist "%~dp0..\.venv\Scripts\pythonw.exe" goto sinvenv

rem Delega en el lanzador .vbs, que arranca la aplicacion sin consola y sin esperarla.
wscript //NoLogo "%~dp0ExcelAgent.vbs"
exit /b 0
:sinvenv
echo Falta el entorno virtual en "%~dp0..\.venv" (ver README, Instalacion).
pause
exit /b 1
