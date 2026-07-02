# -*- coding: utf-8 -*-
"""
매뉴얼용 '실제 화면 캡처' → base64 자산(manual_assets_shots.py) 생성.
────────────────────────────────────────────────────────────
· 사내에서 손으로 캡처한 실제 화면(트레이/메뉴/폴더 등)을 영속 .py 로 굽는다.
  (원본 PNG/JPG는 매일 삭제되므로 base64로 보존 → 삭제돼도 매뉴얼 재생성 가능)
· 캡처 원본 폴더: ../스크린샷/  (파일명 매핑은 아래 SHOTS)
      python make_shot_assets.py
"""
import os
import io
import base64
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "스크린샷")

# key : (파일명, 최대폭, crop(box) 또는 None)
SHOTS = {
    "tray_icon": ("작업표시줄 아이콘.JPG", 360, None),
    "tray_menu": ("작업표시줄 아이콘 우클릭.JPG", 640, (78, 52, 700, 478)),
    "rightclick": ("엑셀파일 우클릭.png", 620, None),
    "folders": ("폴더별 안내.JPG", 840, None),
    "mainwin": ("프로그램 실행시.png", 720, None),
}


def encode(path, maxw, crop):
    im = Image.open(path).convert("RGB")
    if crop:
        im = im.crop(crop)
    if im.width > maxw:
        im = im.resize((maxw, round(im.height * maxw / im.width)))
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii"), im.size, len(buf.getvalue())


def main():
    out = os.path.join(HERE, "manual_assets_shots.py")
    lines = ["# -*- coding: utf-8 -*-",
             '"""매뉴얼용 실제 화면 캡처(base64). make_shot_assets.py 로 재생성."""',
             "SHOTS = {"]
    for key, (fn, maxw, crop) in SHOTS.items():
        p = os.path.join(SRC, fn)
        if not os.path.exists(p):
            print("  (없음, 건너뜀):", fn)
            continue
        b64, size, nbytes = encode(p, maxw, crop)
        lines.append(f'    "{key}": "data:image/png;base64,{b64}",')
        print(f"  {key}: {size}  {nbytes//1024} KB")
    lines.append("}")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("생성됨:", out, "·", round(os.path.getsize(out) / 1024, 1), "KB")


if __name__ == "__main__":
    main()
