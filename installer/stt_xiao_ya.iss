; STT 可爱少女 (xiao_ya) Windows SAPI5 Voice Pack — Inno Setup script
; Builds a single-file Setup.exe that:
;   - requires admin
;   - installs DLL + model bundle to Program Files\STT-xiao_ya-VoicePack
;   - silently installs VC++ redist 2015-2022 x64
;   - registers the DLL via regsvr32 (so SAPI5 picks up the voice)
;   - cleanly unregisters on uninstall

#define MyAppName "STT 可爱少女 (xiao_ya) 中文语音包"
#define MyAppVersion "0.2.0"
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
DefaultGroupName=STT xiao_ya VoicePack
DisableProgramGroupPage=yes
UninstallDisplayName={#MyAppName}

PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Compression=lzma2/ultra64
SolidCompression=yes

OutputDir=output
OutputBaseFilename=STT-xiao_ya-Setup

WizardStyle=modern
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Default.isl"

[Files]
; Our SAPI5 DLL (dynamically linked to sherpa-onnx; delay-loaded)
Source: "payload\stt_xiao_ya_sapi5.dll"; DestDir: "{app}"; Flags: ignoreversion

; sherpa-onnx runtime DLLs — placed next to our DLL. DllMain calls
; SetDllDirectoryW({app}) so delay-load stubs can find them without PATH
; pollution or System32 writes.
Source: "payload\sherpa-onnx-c-api.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\sherpa-onnx-cxx-api.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\onnxruntime.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\onnxruntime_providers_shared.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\cargs.dll"; DestDir: "{app}"; Flags: ignoreversion

; xiao_ya model bundle (copied recursively)
Source: "payload\xiao_ya\*"; DestDir: "{app}\xiao_ya"; Flags: ignoreversion recursesubdirs createallsubdirs

; VC++ redistributable — installed then deleted
Source: "payload\VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
; 1) VC++ Redist 2015-2022 x64 — silent, no reboot
Filename: "{tmp}\VC_redist.x64.exe"; \
  Parameters: "/install /quiet /norestart"; \
  StatusMsg: "正在安装 Microsoft Visual C++ Redistributable..."; \
  Flags: waituntilterminated

; 2) Register our SAPI5 COM DLL
Filename: "{sys}\regsvr32.exe"; \
  Parameters: "/s ""{app}\stt_xiao_ya_sapi5.dll"""; \
  StatusMsg: "正在注册 SAPI5 语音引擎..."; \
  Flags: runascurrentuser waituntilterminated

[UninstallRun]
; Unregister before files are removed
Filename: "{sys}\regsvr32.exe"; \
  Parameters: "/u /s ""{app}\stt_xiao_ya_sapi5.dll"""; \
  Flags: runascurrentuser waituntilterminated

[Code]
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
