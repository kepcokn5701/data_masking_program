# -*- coding: utf-8 -*-
"""
매뉴얼용 Lottie 애니메이션 생성기 (영속 소스).
────────────────────────────────────────────────────────────
· python-lottie 로 라이선스 깨끗한 애니메이션을 '직접 생성'한다(외부 저작물 X).
· lottie 재생기(lottie_light)와 애니메이션 JSON을 base64/문자열로 manual_assets_lottie.py 에 굽는다.
· make_manual.py 가 이를 단일 HTML에 인라인 → 인터넷 없이(업무망) 모션 재생.
      python make_lottie.py
재생기 원본: npm 'lottie-web'(MIT). 애니메이션: 본 스크립트로 자체 생성(제3자 IP 없음).
"""
import os
import json
import base64

from lottie import objects, Point, Color
from lottie.objects import shapes

BLUE = Color(0.145, 0.388, 0.922)
GREEN = Color(0.082, 0.639, 0.290)
WHITE = Color(1, 1, 1)


def _round(stroke):
    for attr, val in (("line_cap", 2), ("line_join", 2)):
        try:
            setattr(stroke, attr, val)
        except Exception:
            pass
    return stroke


def anim_click():
    """‘여기 클릭’ 펄스 — 커지며 사라지는 링 2개 + 중앙 점 (반복)."""
    an = objects.Animation(60, 30)
    an.width = an.height = 200
    layer = objects.ShapeLayer()
    an.add_layer(layer)

    for delay in (0, 20):
        grp = shapes.Group()
        ring = shapes.Ellipse()
        ring.size.value = Point(60, 60)
        ring.position.value = Point(0, 0)
        grp.shapes.insert(0, ring)
        grp.shapes.insert(1, _round(shapes.Stroke(BLUE, 10)))
        tr = grp.transform
        tr.position.value = Point(100, 100)
        tr.scale.add_keyframe(delay, Point(20, 20))
        tr.scale.add_keyframe(delay + 40, Point(150, 150))
        tr.opacity.add_keyframe(delay, 90)
        tr.opacity.add_keyframe(delay + 40, 0)
        layer.shapes.append(grp)

    dot = shapes.Group()
    e = shapes.Ellipse()
    e.size.value = Point(46, 46)
    e.position.value = Point(100, 100)
    dot.shapes.insert(0, e)
    dot.shapes.insert(1, shapes.Fill(BLUE))
    layer.shapes.append(dot)
    return an


def anim_check():
    """완료 체크 — 초록 원이 튀어오르고 흰 체크가 그려짐."""
    an = objects.Animation(50, 30)
    an.width = an.height = 200

    # 배경 원 (팝인)
    bg_layer = objects.ShapeLayer()
    an.add_layer(bg_layer)
    bg = shapes.Group()
    c = shapes.Ellipse()
    c.size.value = Point(170, 170)
    c.position.value = Point(0, 0)
    bg.shapes.insert(0, c)
    bg.shapes.insert(1, shapes.Fill(GREEN))
    bt = bg.transform
    bt.position.value = Point(100, 100)
    bt.scale.add_keyframe(0, Point(0, 0))
    bt.scale.add_keyframe(10, Point(115, 115))
    bt.scale.add_keyframe(16, Point(100, 100))
    bg_layer.shapes.append(bg)

    # 체크 표시 (트림으로 그려짐)
    ck_layer = objects.ShapeLayer()
    an.add_layer(ck_layer)
    grp = shapes.Group()
    path = shapes.Path()
    bez = objects.Bezier()
    bez.add_point(Point(58, 104))
    bez.add_point(Point(88, 134))
    bez.add_point(Point(146, 68))
    bez.closed = False
    path.shape.value = bez
    grp.shapes.append(path)
    grp.shapes.append(_round(shapes.Stroke(WHITE, 16)))
    trim = shapes.Trim()
    trim.start.value = 0
    trim.end.add_keyframe(12, 0)
    trim.end.add_keyframe(30, 100)
    grp.shapes.append(trim)
    ck_layer.shapes.append(grp)
    return an


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    anims = {"click": anim_click(), "check": anim_check()}

    # 재생기(lottie_light) 찾기: scratchpad 또는 node_modules
    player_b64 = ""
    candidates = [
        os.path.join(os.environ.get("TEMP", ""), "..", "claude"),  # placeholder
    ]
    found = None
    for root in [
        r"C:\Users\Admin\AppData\Local\Temp\claude\c--Users-Admin-Desktop-project-ftc-data-masking\c97b4517-db4c-461f-a7cd-dcaa6295808a\scratchpad\package\build\player\lottie_light.min.js",
        os.path.join(here, "lottie_light.min.js"),
    ]:
        if os.path.exists(root):
            found = root
            break
    if found:
        with open(found, "rb") as f:
            player_b64 = base64.b64encode(f.read()).decode("ascii")
        print("재생기 임베드:", round(os.path.getsize(found) / 1024, 1), "KB")
    else:
        print("경고: lottie_light.min.js 를 못 찾음 — PLAYER 비어 있음")

    out = os.path.join(here, "manual_assets_lottie.py")
    with open(out, "w", encoding="utf-8") as f:
        f.write("# -*- coding: utf-8 -*-\n")
        f.write('"""매뉴얼용 Lottie 자산(재생기 base64 + 애니 JSON). make_lottie.py 로 재생성."""\n')
        f.write(f'PLAYER_B64 = "{player_b64}"\n')
        f.write("ANIM = {\n")
        for name, an in anims.items():
            f.write(f'    "{name}": {json.dumps(json.dumps(an.to_dict()), ensure_ascii=False)},\n')
        f.write("}\n")
    print("생성됨:", out, "·", round(os.path.getsize(out) / 1024, 1), "KB")


if __name__ == "__main__":
    main()
