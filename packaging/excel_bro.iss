#ifndef SourceDir
  #error SourceDir command-line define is required
#endif

#ifndef OutputDir
  #error OutputDir command-line define is required
#endif

#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

#define AppName "Excel Bro"
#define AppPublisher "Excel Bro"
#define AppExeName "ExcelBro.exe"
#define AppIdValue "{{F5E94C17-4F55-42A2-ABDC-D2D1CE5A6E91}"
#define ManifestId "9c758d40-c2b8-42d8-a6bc-735bd5c4f34c"
#define CatalogId "{{41f62f5c-cd95-44f2-a0c6-c8cb847fe4e0}"

[Setup]
AppId={#AppIdValue}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableDirPage=no
DisableProgramGroupPage=no
DisableReadyPage=no
DisableFinishedPage=no
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
CloseApplications=force
CloseApplicationsFilter={#AppExeName}
RestartApplications=no
AllowNetworkDrive=no
AllowRootDirectory=no
AllowUNCPath=no
OutputDir={#OutputDir}
OutputBaseFilename=Excel-Bro-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExeName}
CreateUninstallRegKey=yes
UsePreviousAppDir=yes
SetupLogging=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
Type: files; Name: "{app}\uninstall.ps1"

[Icons]
Name: "{group}\Excel Bro"; Filename: "{app}\{#AppExeName}"
Name: "{group}\卸载 Excel Bro"; Filename: "{uninstallexe}"

[Registry]
Root: HKCU; \
    Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "Excel Bro"; \
    ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue
Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{#CatalogId}"; \
    ValueType: string; ValueName: "Id"; ValueData: "{#CatalogId}"; \
    Flags: uninsdeletekey
Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{#CatalogId}"; \
    ValueType: string; ValueName: "Url"; ValueData: "{code:GetCatalogUrl}"; \
    Flags: uninsdeletekey
Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{#CatalogId}"; \
    ValueType: dword; ValueName: "Flags"; ValueData: "1"; \
    Flags: uninsdeletekey
; 直连注册（等同 office-addin-debugging 旁加载）：Excel 重启后功能区自动出现
; Excel Bro 标签，无需用户先在"共享文件夹"里手动添加一次。值名 = 清单 Id，
; 值数据 = 本机清单文件路径。卸载时由 uninsdeletevalue 自动移除。
Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\Wef\Developer"; \
    ValueType: string; ValueName: "{#ManifestId}"; \
    ValueData: "{app}\catalog\manifest.xml"; \
    Flags: uninsdeletevalue

[Code]
function GetCatalogUrl(Param: String): String;
begin
  Result := '\\' + GetEnv('COMPUTERNAME') + '\ExcelBroAddins';
end;

function PowerShellPath(): String;
begin
  Result := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
end;

function IntegrationParameters(Action: String): String;
begin
  Result :=
    '-NoProfile -ExecutionPolicy Bypass -File "' +
    ExpandConstant('{app}\install_tasks.ps1') +
    '" -Action ' + Action +
    ' -InstallDir "' + ExpandConstant('{app}') + '"';
end;

procedure RemoveLegacyRegistration();
begin
  RegDeleteValue(
    HKCU,
    'Software\Microsoft\Office\16.0\Wef\Developer',
    '{#ManifestId}'
  );
  RegDeleteKeyIncludingSubkeys(
    HKCU,
    'Software\Microsoft\Office\16.0\Wef\Developer\{#ManifestId}'
  );
  RegDeleteKeyIncludingSubkeys(
    HKCU,
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel Bro'
  );
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    RemoveLegacyRegistration();
  end;

  if CurStep = ssPostInstall then
  begin
    if (not Exec(
      PowerShellPath(),
      IntegrationParameters('Install'),
      ExpandConstant('{app}'),
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode
    )) or (ResultCode <> 0) then
    begin
      RaiseException(
        'Excel Bro 系统集成失败。安装未完成，请查看安装日志后重试。'
      );
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    if (not Exec(
      PowerShellPath(),
      IntegrationParameters('Uninstall'),
      ExpandConstant('{app}'),
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode
    )) or (ResultCode <> 0) then
    begin
      RaiseException(
        'Excel Bro 卸载未完成。请允许管理员授权后重试。'
      );
    end;
  end;
end;
