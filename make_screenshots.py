# -*- coding: utf-8 -*-
"""
매뉴얼용 실제 스크린샷 캡처 → base64 자산(manual_assets.py) 생성.
────────────────────────────────────────────────────────────
· 앱 창을 실제로 띄워 '창 영역만' 캡처한다(전체화면 X, 사생활 보호).
· PNG는 매일 삭제되므로 파일로 두지 않고 base64로 manual_assets.py(영속 .py)에 저장한다.
· make_manual.py 가 이 자산을 <img>로 인라인 → 단일 HTML 완성.
      python make_screenshots.py
"""
import os
import io
import sys
import time
import base64

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import ImageGrab
import masking_engine as E
from excel_masking import MaskingApp

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(APP_DIR, "파일예시", "전기사용신청서_가상데이터.xlsx")


def _pump(win, n=10, delay=0.05):
    for _ in range(n):
        win.update()
        time.sleep(delay)


def _grab(widget):
    """위젯(창) 영역만 캡처 → PIL Image."""
    widget.update()
    x, y = widget.winfo_rootx(), widget.winfo_rooty()
    w, h = widget.winfo_width(), widget.winfo_height()
    return ImageGrab.grab(bbox=(x, y, x + w, y + h))


def capture():
    # 규칙 창 예시가 채워져 보이도록 임시 규칙 파일 사용
    E.set_rules_path(os.path.join(os.environ.get("TEMP", APP_DIR), "_shot_rules.json"))
    try:
        os.remove(E._RULES_PATH)
    except OSError:
        pass
    E.set_rules_path(E._RULES_PATH)
    E.add_column_rule("계약번호", mode="digits")

    shots = {}

    app = MaskingApp()
    app.geometry("900x740+60+30")
    app.attributes("-topmost", True)
    app.table = E.read_table(SAMPLE)
    app.filepath.set(SAMPLE)
    app._analyze_and_render()
    _pump(app, 16)
    shots["main"] = _grab(app)

    # 규칙 창 열기 → 마지막 Toplevel 캡처
    app._manage_rules()
    _pump(app, 14)
    tops = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel"]
    if tops:
        top = tops[-1]
        top.geometry("+520+120")
        top.attributes("-topmost", True)
        _pump(app, 12)
        shots["rules"] = _grab(top)
        top.destroy()

    app.destroy()
    return shots


def main():
    shots = capture()
    lines = ["# -*- coding: utf-8 -*-",
             '"""매뉴얼용 스크린샷(base64 data URI). make_screenshots.py 로 재생성."""',
             "IMAGES = {"]
    sp = os.environ.get("TEMP", APP_DIR)
    for name, img in shots.items():
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        raw = buf.getvalue()
        img.save(os.path.join(sp, f"_shot_{name}.png"))   # 확인용(임시)
        b64 = base64.b64encode(raw).decode("ascii")
        lines.append(f'    "{name}": "data:image/png;base64,{b64}",')
        print(f"  캡처: {name}  {img.size}  {len(raw)//1024} KB")
    lines.append("}")
    out = os.path.join(APP_DIR, "manual_assets.py")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("생성됨:", out)


if __name__ == "__main__":
    main()
