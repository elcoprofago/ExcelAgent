; Instalador de ExcelAgent (Inno Setup 6). No se compila a mano: lo arma construir.ps1, que prepara la carpeta
; de etapa y pasa MiVersion (de version.py) y Etapa por la linea de comandos.
;
; Instala, por usuario y sin pedir administrador, en %LOCALAPPDATA%\Programs\ExcelAgent:
;   ExcelAgent\   el programa (el boton Actualizar lo pisa con cada version nueva; por eso la carpeta es del usuario)
;   python\       un Python propio con las bibliotecas ya instaladas: no depende del Python que tenga la PC
;   tesseract\    el motor de OCR
; La configuracion (API key, etc.) vive en %APPDATA% y no la toca ni el instalador ni el desinstalador.

#ifndef MiVersion
  #error Compilar con construir.ps1 (falta MiVersion)
#endif
#ifndef Etapa
  #error Compilar con construir.ps1 (falta Etapa)
#endif

[Setup]
AppId={{6F1C2A4E-8B3D-4E5F-9A7B-2C1D0E3F4A5B}
AppName=ExcelAgent
AppVersion={#MiVersion}
AppVerName=ExcelAgent {#MiVersion}
VersionInfoVersion={#MiVersion}
DefaultDirName={autopf}\ExcelAgent
DefaultGroupName=ExcelAgent
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#Etapa}\ExcelAgent.ico
UninstallDisplayIcon={app}\ExcelAgent.ico
UninstallDisplayName=ExcelAgent
CloseApplications=yes

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear un acceso directo en el escritorio"

[InstallDelete]
; El Python propio es todo del instalador: se reemplaza entero para que no queden bibliotecas de una version anterior.
Type: filesandordirs; Name: "{app}\python"

[Files]
Source: "{#Etapa}\ExcelAgent\*"; DestDir: "{app}\ExcelAgent"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Etapa}\python\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Etapa}\tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Etapa}\ExcelAgent.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\ExcelAgent"; Filename: "{sys}\wscript.exe"; Parameters: "//NoLogo ""{app}\ExcelAgent\ExcelAgent.vbs"""; WorkingDir: "{app}\ExcelAgent"; IconFilename: "{app}\ExcelAgent.ico"
Name: "{autodesktop}\ExcelAgent"; Filename: "{sys}\wscript.exe"; Parameters: "//NoLogo ""{app}\ExcelAgent\ExcelAgent.vbs"""; WorkingDir: "{app}\ExcelAgent"; IconFilename: "{app}\ExcelAgent.ico"; Tasks: escritorio

[Run]
Filename: "{sys}\wscript.exe"; Parameters: "//NoLogo ""{app}\ExcelAgent\ExcelAgent.vbs"""; Description: "Abrir ExcelAgent"; Flags: postinstall nowait skipifsilent

[UninstallDelete]
; Lo que Python genera al correr dentro de carpetas del instalador. No se borran logs\, _versiones_anteriores\ ni
; _excelagent_respaldos\: son del usuario, y quedan en la carpeta despues de desinstalar.
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\ExcelAgent\__pycache__"
Type: filesandordirs; Name: "{app}\ExcelAgent\tests\__pycache__"
Type: filesandordirs; Name: "{app}\ExcelAgent\tests\modulos\__pycache__"

[Code]
// VERSION de un version.py ya instalado ('' si no hay). El boton Actualizar pudo haber dejado una version mas nueva
// que la de este instalador: en ese caso se avisa antes de volver atras.
function VersionInstalada(): String;
var
  Lineas: TArrayOfString;
  I, P: Integer;
  L: String;
begin
  Result := '';
  if not LoadStringsFromFile(ExpandConstant('{app}\ExcelAgent\version.py'), Lineas) then Exit;
  for I := 0 to GetArrayLength(Lineas) - 1 do begin
    L := Trim(Lineas[I]);
    if Copy(L, 1, 7) = 'VERSION' then begin
      P := Pos('''', L);
      if P > 0 then begin
        L := Copy(L, P + 1, Length(L));
        P := Pos('''', L);
        if P > 0 then Result := Copy(L, 1, P - 1);
      end;
      Exit;
    end;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Instalada: String;
  A, B: Int64;
begin
  Result := True;
  if CurPageID <> wpReady then Exit;
  Instalada := VersionInstalada();
  if (Instalada = '') or not StrToVersion(Instalada, A) or not StrToVersion('{#MiVersion}', B) then Exit;
  if ComparePackedVersion(A, B) > 0 then
    Result := SuppressibleMsgBox('Ya esta instalada la version ' + Instalada + ', mas nueva que esta ({#MiVersion}).' + #13#10 +
      'Si sigue, se vuelve a la {#MiVersion}. Seguir?', mbConfirmation, MB_YESNO, IDNO) = IDYES;
end;
