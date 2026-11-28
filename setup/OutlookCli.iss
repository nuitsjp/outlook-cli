; The installer is intentionally unsigned. The release workflow does not
; configure an Inno Setup SignTool.

#define AppName "outlook-cli"
#define ShimDir "{localappdata}\Microsoft\WinGet\Links"
#define ShimRegistry "Software\OutlookCli\Shim"

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#ifndef SourceDir
  #define SourceDir "..\src\OutlookCli\bin\Release\net481"
#endif

#ifndef ShimFile
  #define ShimFile "..\src\OutlookCli.Shim\bin\Release\net481\outlook-cli-shim.exe"
#endif

#ifndef OutputDir
  #define OutputDir "..\dist"
#endif

[Setup]
AppId={{B1D86A11-FA5E-4C1C-B8CD-3B5A0E51C43E}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=outlook-cli
DefaultDirName={localappdata}\Programs\{#AppName}
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename={#AppName}-{#AppVersion}-setup
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\outlook-cli.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoDescription=outlook-cli installer
VersionInfoProductName={#AppName}
VersionInfoCompany=outlook-cli

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#ShimFile}"; DestDir: "{#ShimDir}"; DestName: "outlook-cli.exe"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "{#ShimRegistry}"; ValueType: string; ValueName: "Target"; ValueData: "{app}\outlook-cli.exe"; Flags: uninsdeletevalue uninsdeletekeyifempty

[Code]
const
  NetFramework481Release = 533320;

function IsNetFramework481OrLaterInstalled(): Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(
    HKLM64,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release',
    Release) and (Release >= NetFramework481Release);
end;

function InitializeSetup(): Boolean;
begin
  Result := IsNetFramework481OrLaterInstalled();
  if not Result then
    MsgBox(
      'outlook-cli requires .NET Framework 4.8.1 or later. Install that prerequisite and run this installer again.',
      mbError,
      MB_OK);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ExistingShim: TFindRec;
begin
  Result := '';
  if CompareText(RemoveBackslashUnlessRoot(ExpandConstant('{app}')), ExpandConstant('{#ShimDir}')) = 0 then
  begin
    Result := 'Choose an installation directory outside WinGet Links.';
    Exit;
  end;

  if FindFirst(ExpandConstant('{#ShimDir}\outlook-cli.exe'), ExistingShim) then
  begin
    FindClose(ExistingShim);
    if not RegValueExists(HKCU, '{#ShimRegistry}', 'Target') then
      Result := 'An existing outlook-cli.exe in WinGet Links is managed outside this installer.';
  end;
end;
