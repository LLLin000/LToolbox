; L的工具箱 —— Inno Setup 6 安装包脚本（每用户安装，无需管理员）
; 构建：先跑 tools/build_exe.py 与 tools/build_ppam.py --mode exe，再
;   iscc /DAppVersion=1.0.0 installer\LToolbox.iss
; 或直接改下面的 #define。
#define AppName "L的工具箱"
#define AppIdName "LToolbox"
#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif
#define ExeFile "..\dist\LToolbox.exe"
#define PpamFile "..\release\LToolbox.ppam"

[Setup]
AppId={{7C4A1B2E-9D3F-4A55-8E6B-2F1C0D9A7B34}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=LLLin000
AppPublisherURL=https://github.com/LLLin000/LToolbox
DefaultDirName={localappdata}\LToolbox
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=LToolbox-{#AppVersion}-setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName={#AppName}

[Files]
Source: {#ExeFile}; DestDir: {app}; Flags: ignoreversion
Source: {#PpamFile}; DestDir: {userappdata}\Microsoft\AddIns; DestName: "LToolbox.ppam"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox"; ValueType: string; \
  ValueName: "Path"; ValueData: "{userappdata}\Microsoft\AddIns\LToolbox.ppam"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox"; ValueType: string; \
  ValueName: "Title"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox"; ValueType: string; \
  ValueName: "Description"; ValueData: "图片标定 / 比例尺 / 原图像素裁取"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox"; ValueType: dword; \
  ValueName: "AutoLoad"; ValueData: "$FFFFFFFF"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\AddIns\LToolbox.ppam"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  if FindWindowByClassName('PPTFrameClass') <> 0 then
  begin
    if MsgBox('检测到 PowerPoint 正在运行。继续安装前请先完全退出 PowerPoint，否则加载项可能无法写入。' + #13#10 + #13#10 + '现在继续？', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;
