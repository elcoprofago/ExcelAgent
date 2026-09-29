' Lanzador sin consola para ExcelAgent.
' Doble clic aca, o en un acceso directo que apunte a este archivo.
' Rutas relativas a este archivo: el entorno virtual esta en la carpeta de arriba (..\.venv).
Option Explicit
Dim shell, fso, aqui, py, comando
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
aqui = fso.GetParentFolderName(WScript.ScriptFullName)
py = fso.BuildPath(fso.GetParentFolderName(aqui), ".venv\Scripts\pythonw.exe")
If Not fso.FileExists(py) Then
  MsgBox "Falta el entorno virtual en " & fso.GetParentFolderName(py) & " (ver README, Instalacion).", 16, "ExcelAgent"
  WScript.Quit 1
End If
comando = Chr(34) & py & Chr(34) & " " & Chr(34) & fso.BuildPath(aqui, "gui.py") & Chr(34)
' 1 = ventana normal. pythonw.exe no tiene consola, asi que no hay parpadeo.
shell.Run comando, 1, False
