"""
통합 진입점 — 하나의 exe가 인자에 따라 분기한다.
────────────────────────────────────────────────────────────
빌드하면 이 파일 하나가 exe가 되어, 파이썬이 없는 업무망에서도
아래 모든 기능을 수행한다(개별 .py 배포 불필요).

  (인자 없음)              → 트레이 상주앱 실행 (기존 동작)
  --mask <파일...>         → 우클릭 컨텍스트 메뉴 핸들러(파일 마스킹)
  --register              → 우클릭 메뉴 '개인정보 마스킹 사본 만들기' 등록
  --unregister            → 우클릭 메뉴 해제
  --uninstall [옵션]       → 전체 원복(레지스트리+데이터). 옵션은 uninstall.py 참고
  --manual                → 사용설명서(HTML)를 새로 만들어 브라우저로 열기

[설계 의도] 개별 스크립트(mask_cli/register/uninstall)는 순수 파이썬이라
업무망(파이썬 없음)에서 못 돈다. 이 진입점으로 묶어 exe 하나로 배포한다.
각 기능 모듈은 필요할 때만 import 한다(트레이앱의 폴더 생성 등 부작용 회피).
"""
import sys


def main():
    args = sys.argv[1:]
    cmd = args[0] if args else ""

    if cmd == "--mask":
        import mask_cli
        rc = 0
        for path in args[1:]:
            rc = mask_cli.mask_file(path) or rc
        sys.exit(rc)

    if cmd == "--register":
        import register_context_menu
        register_context_menu.register()
        return

    if cmd == "--unregister":
        import register_context_menu
        register_context_menu.unregister()
        return

    if cmd == "--uninstall":
        import uninstall
        sys.argv = [sys.argv[0]] + args[1:]   # uninstall.py 가 자기 인자를 읽도록
        uninstall.main()
        return

    if cmd == "--manual":
        import make_manual
        make_manual.open_manual()
        return

    # 인자 없음 → 트레이 상주앱 (기존 동작)
    import app_tray
    app_tray.main()


if __name__ == "__main__":
    main()
