[Setup]
; Información del instalador
AppName=AutoPersonalHonorario
AppVersion=1.11
DefaultDirName={pf}\AutoPersonalHonorario
DefaultGroupName=AutoPersonalHonorario
OutputDir=Output
OutputBaseFilename=Instalador_AutoPersonalHonorario
Compression=lzma
SolidCompression=yes

[Files]
; Archivos que se copiarán
Source: "C:\Users\jfz\OneDrive - Municipalidad de Vitacura\Documentos\Auto Personas\cx_Freeze\build\exe.win-amd64-3.10\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Crear un acceso directo en el escritorio
Name: "{commondesktop}\AutoPersonalHonorario"; Filename: "{app}\mi_app_autopersonas.exe"
; Crear un acceso directo en el menú de inicio
Name: "{group}\AutoPersonalHonorario"; Filename: "{app}\mi_app_autopersonas.exe"

[Run]
; Ejecutar la aplicación al finalizar la instalación
Filename: "{app}\mi_app_autopersonas.exe"; Description: "{cm:LaunchProgram,AutoPersonalHonorario}"; Flags: nowait postinstall skipifsilent