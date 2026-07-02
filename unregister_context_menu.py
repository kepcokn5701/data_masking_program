"""
우클릭 컨텍스트 메뉴 해제 — register_context_menu.py 로 심은 항목을 제거한다.
(register_context_menu.py 의 unregister() 를 그대로 호출)

되돌리기 원칙: HKEY_CURRENT_USER 아래 우리 키(FtcMaskExport)만 지운다.
다른 프로그램 설정·엑셀 파일연결은 건드리지 않는다.
"""
from register_context_menu import unregister

if __name__ == "__main__":
    unregister()
