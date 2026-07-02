"""
전체 원복(제거) — 레지스트리 + 생성 데이터까지 되돌린다.
────────────────────────────────────────────────────────────
회사에서 "사용 중지" 지시가 내려와도 흔적 없이 원복할 수 있게 한다.

[이 스크립트가 되돌리는 것]
1. 우클릭 컨텍스트 메뉴 (HKEY_CURRENT_USER 레지스트리 키)     ← 항상 제거
2. 프로그램이 만든 데이터 폴더(감시폴더/감사로그)             ← 기본은 '보존', 명시 요청 시만 삭제

[안전 원칙 — 설명가능]
· 레지스트리는 HKCU 아래 '우리 키(FtcMaskExport)'만 지운다.
  다른 프로그램 설정·엑셀 파일연결은 절대 건드리지 않는다.
· 감사로그(감사로그/*.jsonl)는 개인정보보호법상 안전조치 '증빙'이다.
  기본값으로는 지우지 않는다. 정말 지우려면:  uninstall.py --purge-data --yes
· 원본 옆에 만들어진 '_마스킹.xlsx' 결과물은 사용자 파일이므로 건드리지 않는다
  (여기저기 흩어져 있어 자동 수거가 위험 → 사용자가 직접 판단해 삭제).

[사용법]
  python uninstall.py                     # 레지스트리만 제거(데이터 보존) — 기본·안전
  python uninstall.py --purge-data --yes  # 데이터 폴더까지 삭제(감사로그 포함) — 주의
"""
import os
import sys
import shutil
import winreg

VERB = "FtcMaskExport"
EXTS = (".xlsx", ".xls", ".csv")


def _data_dir():
    """mask_cli/app_tray 와 동일 규칙 — 데이터가 어디 쌓였는지 찾기 위함."""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.expanduser("~"), "Documents", "엑셀파일_개인정보_마스킹")
    return os.path.dirname(os.path.abspath(__file__))


# ── 1) 레지스트리 원복 ─────────────────────────────────────────
def remove_registry():
    removed = []
    for ext in EXTS:
        base = rf"Software\Classes\SystemFileAssociations\{ext}\shell\{VERB}"
        for sub in (base + r"\command", base):     # 자식(command)부터 삭제
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub)
                removed.append("HKCU\\" + sub)
            except FileNotFoundError:
                pass
    if removed:
        print("✅ 레지스트리 원복 — 우클릭 메뉴 제거:")
        for r in removed:
            print("   - " + r)
    else:
        print("ℹ️  레지스트리에 등록된 우클릭 메뉴가 없습니다(이미 제거됨).")
    return removed


def _folder_size(path):
    total = 0
    for root, _dirs, files in os.walk(path):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return total


# ── 2) 데이터 폴더 처리 (기본: 보존) ───────────────────────────
def report_data(purge=False):
    root = _data_dir()
    audit = os.path.join(root, "감사로그")
    watch = os.path.join(root, "마스킹_감시폴더")

    print("\n📂 프로그램이 만든 데이터:")
    for label, path, is_audit in (("감사로그(증빙)", audit, True),
                                  ("감시폴더", watch, False)):
        if os.path.isdir(path):
            kb = _folder_size(path) / 1024
            print(f"   - {label}: {path}  ({kb:,.1f} KB)")
        else:
            print(f"   - {label}: (없음)")

    if not purge:
        print("\n🔒 데이터는 보존했습니다(기본값).")
        print("   · 감사로그는 개인정보보호법 안전조치 '증빙'이라 보관을 권장합니다.")
        print("   · 정말 삭제하려면:  python uninstall.py --purge-data --yes")
        return

    # purge 모드 — 감시폴더는 삭제, 감사로그는 한 번 더 경고 후 삭제
    if os.path.isdir(watch):
        shutil.rmtree(watch, ignore_errors=True)
        print(f"\n🗑️  감시폴더 삭제: {watch}")
    if os.path.isdir(audit):
        shutil.rmtree(audit, ignore_errors=True)
        print(f"🗑️  감사로그 삭제(증빙 소멸 주의): {audit}")


def main():
    args = set(sys.argv[1:])
    purge = "--purge-data" in args
    confirmed = "--yes" in args

    print("── 엑셀 개인정보 마스킹 — 원복(제거) ──")
    remove_registry()

    if purge and not confirmed:
        print("\n⚠️  --purge-data 는 감사로그(증빙)까지 삭제합니다.")
        print("    실수 방지를 위해 --yes 를 함께 붙여야 실제 삭제됩니다.")
        report_data(purge=False)
        return
    report_data(purge=purge)

    print("\n※ 참고:")
    print("  · 원본 옆 '_마스킹.xlsx' 결과물은 사용자 파일이라 자동 삭제하지 않습니다.")
    print("  · 트레이 상주앱(설치형)을 깔았다면 [제어판 > 프로그램 제거]에서 지우세요.")
    print("  · 소스 코드(이 폴더)는 그대로 둡니다 — 필요 시 폴더째 삭제하세요.")
    print("✅ 원복 완료.")


if __name__ == "__main__":
    main()
