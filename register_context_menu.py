r"""
우클릭 컨텍스트 메뉴 등록 — "개인정보 마스킹본 만들기"
────────────────────────────────────────────────────────────
탐색기에서 .xlsx / .xls 파일을 우클릭했을 때 나오는 메뉴에
"개인정보 마스킹본 만들기" 항목을 추가한다. 누르면 mask_cli.py 가
그 파일을 마스킹한다.

[동작 원리 — 설명가능]
Windows 탐색기는 파일을 우클릭할 때 '레지스트리'라는 시스템 설정 저장소를
읽어 메뉴를 그린다. 특정 확장자에 메뉴를 붙이려면 아래 위치에 값을 심는다:

  HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.xlsx\shell\<동사>
      (MUIVerb) = 메뉴에 보일 이름
      (Icon)    = 메뉴 아이콘(선택)
      \command
          (기본값) = "<pythonw.exe>" "<mask_cli.py>" "%1"

· HKEY_CURRENT_USER 아래라 '관리자 권한 불필요'(내 계정에만 등록).
· SystemFileAssociations\.xlsx 는 엑셀의 파일연결(Excel.Sheet.12)을 건드리지
  않고 '확장자'에만 메뉴를 얹는 안전한 위치.
· "%1" 은 사용자가 우클릭한 '그 파일의 전체 경로'로 치환된다.
· 값을 지우면(=unregister) 메뉴도 사라진다. 되돌리기 쉽다.
"""
import os
import sys
import winreg

APP_DIR = os.path.dirname(os.path.abspath(__file__))
VERB = "FtcMaskExport"                       # 내부 식별자(레지스트리 키 이름)
LABEL = "🔒 개인정보 마스킹 사본 만들기"        # 사용자에게 보이는 메뉴 이름
EXTS = (".xlsx", ".xls", ".csv")


def _command():
    """실행 명령행. 배포(exe)면 'exe --mask %1', 개발(소스)이면 pythonw + mask_cli.py."""
    if getattr(sys, "frozen", False):        # PyInstaller 등으로 exe 빌드된 경우
        exe = sys.executable
        return f'"{exe}" --mask "%1"'         # main.py 가 --mask 로 분기
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = sys.executable                 # pythonw 없으면 python 사용(콘솔 잠깐 뜸)
    cli = os.path.join(APP_DIR, "mask_cli.py")
    return f'"{pyw}" "{cli}" "%1"'


def _icon_path():
    """메뉴 아이콘 경로. exe 배포 시엔 exe 자체 아이콘, 소스면 옆의 .ico."""
    if getattr(sys, "frozen", False):
        return sys.executable                # onefile exe에 내장된 아이콘 사용
    ico = os.path.join(APP_DIR, "FTC_Client_Main.ico")
    return ico if os.path.exists(ico) else ""


def register():
    cmd = _command()
    icon = _icon_path()
    for ext in EXTS:
        key_path = rf"Software\Classes\SystemFileAssociations\{ext}\shell\{VERB}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
            winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, LABEL)
            if icon:
                winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\command") as c:
            winreg.SetValueEx(c, None, 0, winreg.REG_SZ, cmd)   # (기본값)에 명령행
    print("✅ 등록 완료 — .xlsx/.xls 우클릭 메뉴에 '개인정보 마스킹본 만들기' 추가")
    print(f"   실행 명령: {cmd}")
    print("   (탐색기가 이미 떠 있으면 새 창에서 반영됩니다.)")


def unregister():
    for ext in EXTS:
        base = rf"Software\Classes\SystemFileAssociations\{ext}\shell\{VERB}"
        for sub in (base + r"\command", base):     # 자식부터 삭제
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub)
            except FileNotFoundError:
                pass
    print("✅ 해제 완료 — 우클릭 메뉴에서 항목 제거")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("-u", "--uninstall", "remove"):
        unregister()
    else:
        register()
