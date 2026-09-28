@echo off
title ExcelAgent
if not exist D:\REPOS\EXCEL\.venv\Scripts\pythonw.exe goto sinvenv

rem Delega en el lanzador .vbs, que arranca la aplicacion sin consola y sin esperarla.
wscript //NoLogo D:\REPOS\EXCEL\ExcelAgent\ExcelAgent.vbs
exit /b 0
:sinvenv
echo Falta el entorno virtual en D:\REPOS\EXCEL\.venv
:fin
