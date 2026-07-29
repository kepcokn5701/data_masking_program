"""
데이터 보안처리 프로그램 — 트레이 통합 앱 (메인 실행 파일)
────────────────────────────────────────────────────────────
· 작업표시줄(시계 옆) 트레이에 상주한다.
· 트레이 아이콘을 클릭하면 '마스킹 도구' 창이 열린다(파일 선택→마스킹).
· 동시에 '마스킹_감시폴더'를 자동 감시 — 넣은 엑셀을 자동 마스킹(_완료 폴더).
· 마스킹 로직은 공용 엔진(masking_engine)을 그대로 사용 — 웹서버와 동일.
"""
import os
import sys
import time
import queue
import threading

import pystray
from pystray import Menu, MenuItem
from PIL import Image, ImageDraw

from excel_masking import MaskingApp
from masking_engine import (read_table, analyze_dataframe, mask_dataframe,
                            write_workbook, count_detections, set_rules_path)
import register_context_menu as ctxmenu


def _data_dir():
    """설정·규칙·감사로그 폴더 — 공용 코어(masking_engine)의 정의를 그대로 쓴다.
    네 군데에 같은 코드를 복사해 두면 한쪽만 고쳐져 서로 다른 폴더를 보게 된다."""
    from masking_engine import data_dir
    return data_dir()


WATCH = os.path.join(_data_dir(), "자동_마스킹_폴더")
DONE = os.path.join(WATCH, "_완료(마스킹된파일)")
SRC_KEEP = os.path.join(WATCH, "_원본보관")
ERR = os.path.join(WATCH, "_오류")
for _d in (WATCH, DONE, SRC_KEEP, ERR):
    os.makedirs(_d, exist_ok=True)

# 조직 학습형 규칙 파일(GUI·우클릭·감시폴더가 같은 규칙 공유)
set_rules_path(os.path.join(_data_dir(), "masking_rules.json"))

POLL_SEC = 3
state = {"paused": False, "running": True, "done": 0}
_seen = {}     # {path: 직전 폴링 크기} — 복사 중 파일 거르기

# ── 최초 실행 표시 파일 ────────────────────────────────────────
# 이 앱은 트레이(시계 옆)에 조용히 상주해서, 실행해도 창이 뜨지 않는다.
# 그 탓에 직원들이 '프로그램이 켜지긴 한 건가?' 하고 알아채지 못했다.
# → 처음 한 번만 사용설명서를 자동으로 띄워 '켜졌다'는 걸 눈으로 알려준다.
#   (파일 하나로 기억하므로 두 번째부터는 뜨지 않는다. 지우면 다시 뜬다.)
FIRST_RUN_FLAG = os.path.join(_data_dir(), "최초실행안내완료.txt")


# ── 트레이 아이콘 그림 ─────────────────────────────────────────
def _icon_image(active=True):
    img = Image.new("RGBA", (64, 64), (30, 58, 95, 255))
    d = ImageDraw.Draw(img)
    body = (22, 163, 74) if active else (148, 163, 184)
    d.rounded_rectangle([16, 30, 48, 54], radius=4, fill=body)        # 자물쇠 몸통
    d.arc([22, 14, 42, 38], start=180, end=360, fill=body, width=5)   # 고리
    d.ellipse([29, 38, 35, 44], fill=(15, 23, 42, 255))               # 열쇠구멍
    return img


def _notify(icon, title, msg):
    try:
        icon.notify(msg[:200], title)
    except Exception:
        pass


# ── 폴더 자동 감시 ─────────────────────────────────────────────
def _unique(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}({i}){ext}"):
        i += 1
    return f"{base}({i}){ext}"


def _move(path, folder):
    try:
        os.replace(path, _unique(os.path.join(folder, os.path.basename(path))))
    except OSError:
        pass


def _process(path, icon):
    name = os.path.basename(path)
    base = os.path.splitext(name)[0]
    try:
        table = read_table(path)
        targets = [c["name"] for c in analyze_dataframe(table) if c["suggest"]]
        result, report, ref = mask_dataframe(table, targets)
        write_workbook(_unique(os.path.join(DONE, f"{base}_마스킹.xlsx")), result, report, ref)
        _move(path, SRC_KEEP)
        state["done"] += 1
        _notify(icon, "마스킹 완료",
                f"{name} → 완료 폴더 (개인정보 {count_detections(report)}건 마스킹)")
    except Exception as e:
        _move(path, ERR)
        _notify(icon, "처리 실패", f"{name}: {e}")


def _watch_loop(icon):
    while state["running"]:
        if not state["paused"]:
            try:
                for fn in os.listdir(WATCH):
                    if state["paused"] or not state["running"]:
                        break
                    p = os.path.join(WATCH, fn)
                    if not os.path.isfile(p):
                        continue
                    if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xls")):
                        continue
                    try:
                        sz = os.path.getsize(p)
                    except OSError:
                        continue
                    if _seen.get(p) == sz and sz > 0:     # 크기 안정(복사 완료) 후 처리
                        _seen.pop(p, None)
                        _process(p, icon)
                    else:
                        _seen[p] = sz
            except Exception:
                pass
        time.sleep(POLL_SEC)


# ── 메인: tkinter(메인 스레드) + pystray(별 스레드) ────────────
def main():
    app = MaskingApp()
    app.withdraw()                                    # 트레이 상주 — 시작 시 창 숨김
    app.protocol("WM_DELETE_WINDOW", app.withdraw)    # 창 닫기 → 숨김(트레이 유지)

    cmds = queue.Queue()                              # 트레이 스레드 → tk 스레드 안전 전달

    def _pump():
        try:
            while True:
                cmds.get_nowait()()
        except queue.Empty:
            pass
        app.after(120, _pump)
    app.after(120, _pump)

    def show(icon=None, item=None):
        cmds.put(lambda: (app.deiconify(), app.state("normal"),
                          app.lift(), app.focus_force()))

    def open_folder(folder):
        return lambda icon, item: os.startfile(folder)

    def toggle_pause(icon, item):
        state["paused"] = not state["paused"]
        icon.icon = _icon_image(not state["paused"])

    def open_manual(icon, item):
        import make_manual
        try:
            make_manual.open_manual()
        except Exception as e:
            _notify(icon, "설명서 열기 실패", str(e))

    def register_menu(icon, item):
        try:
            ctxmenu.register()
            _notify(icon, "오른쪽클릭 메뉴 추가됨",
                    "이제 파일에서 마우스 오른쪽 클릭 → '개인정보 마스킹 사본 만들기' 사용 가능")
        except Exception as e:
            _notify(icon, "추가 실패", str(e))

    def unregister_menu(icon, item):
        try:
            ctxmenu.unregister()
            _notify(icon, "오른쪽클릭 메뉴에서 뺐음", "오른쪽 클릭 메뉴에서 항목이 사라졌습니다.")
        except Exception as e:
            _notify(icon, "빼기 실패", str(e))

    def quit_all(icon, item):
        state["running"] = False
        try:
            icon.stop()
        except Exception:
            pass
        cmds.put(app.destroy)

    icon = pystray.Icon(
        "n2sf_mask",
        icon=_icon_image(True),
        title="데이터 보안처리 프로그램",
        menu=Menu(
            MenuItem("보안처리 도구 열기", show, default=True),
            MenuItem("📖 사용설명서 열기", open_manual),
            Menu.SEPARATOR,
            MenuItem("자동 마스킹 폴더 열기 (넣으면 자동 처리)", open_folder(WATCH)),
            MenuItem("완료된 파일 폴더 열기", open_folder(DONE)),
            MenuItem("자동 마스킹 잠시 멈춤", toggle_pause, checked=lambda i: state["paused"]),
            Menu.SEPARATOR,
            MenuItem("오른쪽클릭에 '마스킹 사본 만들기' 추가", register_menu),
            MenuItem("오른쪽클릭 메뉴에서 빼기", unregister_menu),
            Menu.SEPARATOR,
            MenuItem("종료", quit_all),
        ),
    )
    threading.Thread(target=icon.run, daemon=True).start()
    threading.Thread(target=_watch_loop, args=(icon,), daemon=True).start()

    # ── 처음 실행이면 사용설명서를 자동으로 띄운다 ─────────────────
    # 트레이 상주앱이라 실행해도 화면에 아무 변화가 없어서, 직원들이
    # 켜진 줄 모르고 그냥 지나쳤다. 처음 한 번은 설명서를 열어 알려준다.
    def _first_run_guide():
        if os.path.exists(FIRST_RUN_FLAG):
            return                                   # 이미 안내했음 → 조용히 넘어감
        try:
            import make_manual
            make_manual.open_manual()                # 설명서 HTML 생성 후 브라우저로 열기
            _notify(icon, "데이터 보안처리 프로그램 — 설치 완료",
                    "시계 옆 트레이에 상주합니다.\n"
                    "아이콘을 클릭하면 마스킹 도구가 열립니다.")
        except Exception as e:
            # 설명서를 못 열어도 최소한 '켜졌다'는 것은 알려야 한다
            _notify(icon, "데이터 보안처리 프로그램 실행됨",
                    f"시계 옆 트레이 아이콘을 클릭해 사용하세요.\n(설명서 열기 실패: {e})")
        try:
            with open(FIRST_RUN_FLAG, "w", encoding="utf-8") as f:
                f.write("이 파일이 있으면 최초 실행 안내를 다시 띄우지 않습니다.\n"
                        "설명서를 다시 보려면 트레이 아이콘 우클릭 → '사용설명서 열기'.\n")
        except Exception:
            pass                                     # 기록 실패가 실행을 막지는 않도록

    app.after(1500, _first_run_guide)                # 트레이 아이콘이 뜬 뒤 안내
    app.mainloop()


if __name__ == "__main__":
    main()
