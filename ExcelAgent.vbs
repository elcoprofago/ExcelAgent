' Lanzador sin consola para ExcelAgent.
' Doble clic aca, o en un acceso directo que apunte a este archivo.
Option Explicit
Dim shell, comando
Set shell = CreateObject("WScript.Shell")
If Not CreateObject("Scripting.FileSystemObject").FileExists("D:\REPOS\EXCEL\.venv\Scripts\pythonw.exe") Then
  MsgBox "Falta el entorno virtual en D:\REPOS\EXCEL\.venv", 16, "ExcelAgent"
  WScript.Quit 1
End If
comando = Chr(34) + "D:\REPOS\EXCEL\.venv\Scripts\pythonw.exe" + Chr(34) + " " + Chr(34) + "D:\REPOS\EXCEL\ExcelAgent\gui.py" + Chr(34)
' 1 = ventana normal. pythonw.exe no tiene consola, asi que no hay parpadeo.
shell.Run comando, 1, False
