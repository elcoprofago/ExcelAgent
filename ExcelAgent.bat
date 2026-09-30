@echo off
title ExcelAgent
rem Rutas relativas a este archivo: Python esta en la carpeta de arriba (..\python del instalador, o ..\.venv).
rem Asi funciona igual en D:\REPOS\EXCEL (notebook) que en F:\source\repos\EXCEL.
if exist "%~dp0..\python\pythonw.exe" goto lanzar
if not exist "%~dp0..\.venv\Scripts\pythonw.exe" goto sinvenv
:lanzar

rem Delega en el lanzador .vbs, que arranca la aplicacion sin consola y sin esperarla.
wscript //NoLogo "%~dp0ExcelAgent.vbs"
exit /b 0
:sinvenv
echo Falta Python: ni "%~dp0..\python" (instalador) ni "%~dp0..\.venv" (ver README, Entorno).
pause
exit /b 1
