' Lanzador sin consola para ExcelAgent.
' Doble clic aca, o en un acceso directo que apunte a este archivo.
' Busca Python en la carpeta de arriba, en este orden:
'   ..\python   el Python propio que trae el instalador (no depende de lo que haya instalado en la PC);
'   ..\.venv    el entorno virtual de desarrollo (ver README, Entorno).
Option Explicit
Dim shell, fso, aqui, raiz, py, comando
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
aqui = fso.GetParentFolderName(WScript.ScriptFullName)
raiz = fso.GetParentFolderName(aqui)
py = fso.BuildPath(raiz, "python\pythonw.exe")
If Not fso.FileExists(py) Then
  py = fso.BuildPath(raiz, ".venv\Scripts\pythonw.exe")
  If Not fso.FileExists(py) Then
    MsgBox "No encuentro Python para ExcelAgent: falta " & fso.BuildPath(raiz, "python") & vbCrLf & vbCrLf & _
      "Volver a correr el instalador de ExcelAgent.", 16, "ExcelAgent"
    WScript.Quit 1
  End If
  ' Un entorno copiado de otra PC apunta (pyvenv.cfg, linea home) a un Python que aca no existe.
  Dim cfg, archivo, linea, base
  cfg = fso.BuildPath(raiz, ".venv\pyvenv.cfg")
  If fso.FileExists(cfg) Then
    Set archivo = fso.OpenTextFile(cfg, 1)
    Do Until archivo.AtEndOfStream
      linea = Trim(archivo.ReadLine)
      If LCase(Left(linea, 4)) = "home" Then
        base = Trim(Mid(linea, InStr(linea, "=") + 1))
        If Not fso.FileExists(fso.BuildPath(base, "pythonw.exe")) Then
          MsgBox "El entorno virtual apunta a un Python que no esta en esta PC:" & vbCrLf & base & vbCrLf & vbCrLf & _
            "Instalar ExcelAgent con su instalador, o ejecutar preparar_entorno.bat (en esta misma carpeta).", 16, "ExcelAgent"
          WScript.Quit 1
        End If
      End If
    Loop
    archivo.Close
  End If
End If
' -E -s: sin variables PYTHON* ni bibliotecas del usuario (%APPDATA%\Python); solo las del Python elegido.
comando = Chr(34) & py & Chr(34) & " -E -s " & Chr(34) & fso.BuildPath(aqui, "gui.py") & Chr(34)
' 1 = ventana normal. pythonw.exe no tiene consola, asi que no hay parpadeo.
shell.Run comando, 1, False