; 데이터 보안처리 프로그램 — 사용자 설치형(관리자 권한 불필요) 인스톨러
; Inno Setup 6 스크립트.  컴파일: ISCC.exe installer.iss
; 결과: installer_output\데이터_보안처리_프로그램_설치.exe

#define AppName "데이터 보안처리 프로그램"
#define AppVer  "1.0"
#define AppExe  "data_security_program.exe"

[Setup]
AppId={{A7E3C1F2-N2SF-4D5E-9A0B-EXELMASK0001}
AppName={#AppName}
AppVersion={#AppVer}
AppPublisher=KEPCO
; --- 관리자 권한 불필요(사용자 설치): 사내망 잠긴 PC에서도 설치 가능 ---
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=데이터_보안처리_프로그램_설치
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 작업:"
Name: "startupicon"; Description: "Windows 시작 시 자동 실행(트레이 상주)"; GroupDescription: "추가 작업:"; Flags: checkedonce

[Files]
Source: "dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{#AppName} 제거"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Registry]
; 시작 시 자동 실행(현재 사용자) — 제거 시 함께 삭제
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#AppName}"; ValueData: """{app}\{#AppExe}"""; \
  Tasks: startupicon; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#AppExe}"; Description: "지금 실행"; Flags: nowait postinstall skipifsilent
