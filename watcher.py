"""
N2SF 마스킹 폴더 감시 — 트레이 상주 앱
────────────────────────────────────────────────────────────
'마스킹_감시폴더'에 엑셀을 넣으면 자동으로 마스킹본을 '_완료'에 생성한다.
작업표시줄 우측(시계 옆) 트레이 아이콘으로 상주하며, 우클릭 메뉴로 제어한다.
마스킹 로직은 공용 엔진(masking_engine.py)을 그대로 사용 — 데스크톱·웹과 동일.
"""
import os
import sys
import time
import threading

from masking_engine import (read_table, analyze_dataframe, mask_dataframe,
                            write_workbook, count_detections)
import pystray
from pystray import Menu, MenuItem
from PIL import Image, ImageDraw


def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


WATCH = os.path.join(_base_dir(), "마스킹_감시폴더")
DONE = os.path.join(WATCH, "_완료")
SRC_KEEP = os.path.join(WATCH, "_원본")
ERR = os.path.join(WATCH, "_오류")
for _d in (WATCH, DONE, SRC_KEEP, ERR):
    os.makedirs(_d, exist_ok=True)

POLL_SEC = 3
state = {"paused": False, "running": True, "done": 0}
_seen_size = {}     # {path: 직전 폴링 때 크기} — 복사 중(불완전) 파일 거르기


def _icon_image(active=True):
    """자물쇠 모양 트레이 아이콘 생성."""
    img = Image.new("RGBA", (64, 64), (30, 58, 95, 255))
    d = ImageDraw.Draw(img)
    body = (22, 163, 74) if active else (148, 163, 184)
    d.rounded_rectangle([16, 30, 48, 54], radius=4, fill=body)        # 자물쇠 몸통
    d.arc([22, 14, 42, 38], start=180, end=360, fill=body, width=5)   # 고리
    d.ellipse([29, 38, 35, 44], fill=(15, 23, 42, 255))               # 열쇠구멍
    return img


def _unique(path):
    """이미 있으면 (1), (2)… 붙여 덮어쓰기 방지."""
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
        out = _unique(os.path.join(DONE, f"{base}_마스킹.xlsx"))
        write_workbook(out, result, report, ref)
        _move(path, SRC_KEEP)
        state["done"] += 1
        n = count_detections(report)
        _notify(icon, "마스킹 완료", f"{name} → _완료 폴더 ({n}건 마스킹)")
    except Exception as e:
        _move(path, ERR)
        _notify(icon, "처리 실패", f"{name}: {e}")


def _notify(icon, title, msg):
    try:
        icon.notify(msg[:200], title)
    except Exception:
        pass


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
                    low = fn.lower()
                    if fn.startswith("~$") or not low.endswith((".xlsx", ".xls")):
                        continue
                    # 복사 중 파일 거르기: 크기가 직전 폴링과 같아야(=안정) 처리
                    try:
                        sz = os.path.getsize(p)
                    except OSError:
                        continue
                    if _seen_size.get(p) == sz and sz > 0:
                        _seen_size.pop(p, None)
                        _process(p, icon)
                    else:
                        _seen_size[p] = sz
            except Exception:
                pass
        time.sleep(POLL_SEC)


# ── 트레이 메뉴 동작 ───────────────────────────────────────────
def _open(folder):
    return lambda icon, item: os.startfile(folder)


def _toggle_pause(icon, item):
    state["paused"] = not state["paused"]
    icon.icon = _icon_image(active=not state["paused"])
    icon.title = "N2SF 마스킹 감시 — " + ("일시정지" if state["paused"] else "동작 중")


def _quit(icon, item):
    state["running"] = False
    icon.stop()


def main():
    icon = pystray.Icon(
        "n2sf_watcher",
        icon=_icon_image(True),
        title="N2SF 마스킹 감시 — 동작 중",
        menu=Menu(
            MenuItem(lambda i: f"감시폴더 열기 ({state['done']}건 처리됨)", _open(WATCH)),
            MenuItem("완료폴더 열기", _open(DONE)),
            MenuItem("원본보관 폴더 열기", _open(SRC_KEEP)),
            Menu.SEPARATOR,
            MenuItem("일시정지", _toggle_pause, checked=lambda i: state["paused"]),
            MenuItem("종료", _quit),
        ),
    )
    threading.Thread(target=_watch_loop, args=(icon,), daemon=True).start()
    icon.run()


if __name__ == "__main__":
    main()
