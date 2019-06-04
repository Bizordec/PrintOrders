; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "PrintOrders"
#define MyAppVersion "3.0"
#define MyAppPublisher "Ilia B."
#define MyAppExeName "PrintOrders.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{C4156B3B-9B0C-479E-A960-88991CD91A63}
AppName="Print Orders"
AppVersion={#MyAppVersion}
AppMutex=PrintOrders
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=D:\Mirea_stuff\PrintJobs\PrintOrders
OutputBaseFilename=PrintOrders_setup{#MyAppVersion}
; "ArchitecturesInstallIn64BitMode=x64" requests that the install be
; done in "64-bit mode" on x64, meaning it should use the native
; 64-bit Program Files directory and the 64-bit view of the registry.
; On all other architectures it will install in "32-bit mode".
;rchitecturesInstallIn64BitMode=x64
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableWelcomePage=no

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
; Install MyProg-x64.exe if running in 64-bit mode (x64; see above),
; MyProg.exe otherwise.
Source: "PrintOrders.exe"; DestDir: "{app}"; Flags: ignoreversion
;Source: "PrintOrders-x32.exe"; DestDir: "{app}"; Check: not Is64BitInstallMode; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
; dll used to check running notepad at install time
Source: psvince.dll; flags: dontcopy
;psvince is installed in {app} folder, so it will be
;loaded at uninstall time ;to check if notepad is running
Source: psvince.dll; DestDir: {app}

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commonstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\11.0\Word\Options"; ValueType: dword; ValueName: "ForceSetCopyCount"; ValueData: "1"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Microsoft\Office\12.0\Word\Options"; ValueType: dword; ValueName: "ForceSetCopyCount"; ValueData: "1"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\Word\Options"; ValueType: dword; ValueName: "ForceSetCopyCount"; ValueData: "1"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Options"; ValueType: dword; ValueName: "ForceSetCopyCount"; ValueData: "1"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Options"; ValueType: dword; ValueName: "ForceSetCopyCount"; ValueData: "1"; Flags: createvalueifdoesntexist

;[UninstallRun]
;Filename: "{cmd}"; Parameters: "/C ""taskkill /im {#MyAppExeName} /f /t"

[Code]
function IsAppRunning(const FileName: string): Boolean;
var
  FWMIService: Variant;
  FSWbemLocator: Variant;
  FWbemObjectSet: Variant;
begin
  Result := false;
  FSWbemLocator := CreateOleObject('WBEMScripting.SWBEMLocator');
  FWMIService := FSWbemLocator.ConnectServer('', 'root\CIMV2', '', '');
  FWbemObjectSet := FWMIService.ExecQuery(Format('SELECT Name FROM Win32_Process Where Name="%s"',[FileName]));
  Result := (FWbemObjectSet.Count > 0);
  FWbemObjectSet := Unassigned;
  FWMIService := Unassigned;
  FSWbemLocator := Unassigned;
end;

function InitializeSetup: boolean;
begin
  Result := not IsAppRunning('notepad.exe');
  if not Result then
  MsgBox('notepad.exe is running. Please close the application before running the installer ', mbError, MB_OK);
end;