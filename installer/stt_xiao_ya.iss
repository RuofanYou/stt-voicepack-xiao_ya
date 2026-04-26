; STT 可爱少女 (xiao_ya) Windows SAPI5 Voice Pack — Inno Setup script
; Single-file Setup-v6.exe that:
;   - 自动卸载任何旧版本（PrepareToInstall 阶段静默调旧 uninstaller）
;   - 强制装到 C:\STT-xiao_ya-VoicePack（UsePreviousAppDir=no, 不沿用旧目录）
;   - 静默装 VC++ redist 2015-2022 x64
;   - regsvr32 注册 SAPI5 COM DLL
;   - 干净卸载 + 注册表清理

#define MyAppName "STT 可爱少女 (xiao_ya) 中文语音包"
#define MyAppVersion "0.6.0"
#define MyAppVersionTag "v6"
#define MyAppPublisher "STT Voice Pack Team"
#define MyAppURL "https://github.com/RuofanYou/stt-voicepack-xiao_ya"

[Setup]
AppId={{5C7A9D4A-3B6F-4E18-9A2D-7B1E2F8C6A31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
VersionInfoVersion={#MyAppVersion}

DefaultDirName=C:\STT-xiao_ya-VoicePack
; 关键: 同 AppId 的旧版本不再沿用旧路径, 强制走 DefaultDirName
UsePreviousAppDir=no
DefaultGroupName=STT xiao_ya VoicePack
DisableProgramGroupPage=yes
UninstallDisplayName={#MyAppName} {#MyAppVersionTag}

PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Compression=lzma2/ultra64
SolidCompression=yes

OutputDir=output
OutputBaseFilename=STT-xiao_ya-Setup-{#MyAppVersionTag}

WizardStyle=modern
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "payload\stt_xiao_ya_sapi5.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\sherpa-onnx-c-api.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\sherpa-onnx-cxx-api.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\onnxruntime.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\onnxruntime_providers_shared.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\cargs.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\xiao_ya\*"; DestDir: "{app}\xiao_ya"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "payload\VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
Filename: "{tmp}\VC_redist.x64.exe"; \
  Parameters: "/install /quiet /norestart"; \
  StatusMsg: "正在安装 Microsoft Visual C++ Redistributable..."; \
  Flags: waituntilterminated

Filename: "{sys}\regsvr32.exe"; \
  Parameters: "/s ""{app}\stt_xiao_ya_sapi5.dll"""; \
  StatusMsg: "正在注册 SAPI5 语音引擎..."; \
  Flags: runascurrentuser waituntilterminated

[UninstallRun]
Filename: "{sys}\regsvr32.exe"; \
  Parameters: "/u /s ""{app}\stt_xiao_ya_sapi5.dll"""; \
  Flags: runascurrentuser waituntilterminated

[Code]
// =============================================================
// 自动卸载旧版本：装 v6 之前先静默调旧 uninstaller
// =============================================================
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1';
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

procedure UnInstallOldVersion();
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then
  begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    Exec(sUnInstallString,
         '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART',
         '', SW_HIDE, ewWaitUntilTerminated, iResultCode);
  end;
end;

function IsX64System: Boolean;
begin
  Result := IsWin64;
end;

function InitializeSetup(): Boolean;
begin
  if not IsX64System then
  begin
    MsgBox('本语音包仅支持 64 位 Windows。', mbError, MB_OK);
    Result := False;
    exit;
  end;
  Result := True;
end;

// 在文件被复制之前自动卸载已存在的旧版本
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  if IsUpgrade() then
    UnInstallOldVersion();
  Result := '';
end;
