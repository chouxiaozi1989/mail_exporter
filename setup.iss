; Inno Setup Compiler配置文件
; 用于将邮件导出工具打包成Windows安装程序

#define MyAppName "邮件导出工具"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "chouxiaozi1989"
#define MyAppURL "https://github.com/chouxiaozi1989/mail_exporter"
#define MyAppExeName "mail_exporter_gui.exe"
#define MyAppDescription "邮箱邮件导出工具 - 支持批量导出邮件到CSV文件"

[Setup]
; 注意: AppId的值唯一标识此应用程序。
; 不要为其他安装程序使用相同的AppId值。
; (要生成新的GUID，请在IDE中单击 工具|生成GUID。)
AppId={{8B5A2C3D-4E6F-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; 移除以下行以在管理安装模式下运行（为所有用户安装）。
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=dist
OutputBaseFilename=邮件导出工具_v{#MyAppVersion}_安装程序
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
; 主程序文件
Source: "dist\mail_exporter_gui.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\mail_exporter.exe"; DestDir: "{app}"; Flags: ignoreversion
; Python运行时和依赖库（如果使用PyInstaller打包）
Source: "dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 文档文件
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
; 注意: 不要在任何共享系统文件上使用"Flags: ignoreversion"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\*.pyc"

[Code]
// 自定义安装过程代码
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

// 检查是否已安装旧版本
function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

// 卸载旧版本
function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  Result := 0;
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

// 安装前检查
function InitializeSetup(): Boolean;
begin
  if IsUpgrade() then
  begin
    if MsgBox('检测到已安装的旧版本，是否要先卸载？', mbConfirmation, MB_YESNO) = IDYES then
    begin
      UnInstallOldVersion();
    end;
  end;
  Result := True;
end;

// 安装完成后的操作
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // 可以在这里添加安装后的自定义操作
  end;
end;