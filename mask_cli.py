"""
FTC 반출 직전 원클릭 마스킹 — 우클릭 컨텍스트 메뉴용 (headless one-shot)
────────────────────────────────────────────────────────────
탐색기에서 엑셀 파일을 우클릭 → "개인정보 마스킹본 만들기" 를 누르면
이 스크립트가 실행되어 원본 옆에 `파일명_마스킹.xlsx` 를 생성한다.
사용자는 그 마스킹본을 FTC 로 반출한다.

[배경] FTC(휴네시온 자료전송시스템)는 상용 바이너리라 전송 순간을
       직접 가로챌 수 없다. 그래서 '전송 순간 자동'이 아니라
       '전송 동선에 밀착한 원클릭'으로 마스킹을 유도한다. (강제 아님)
       → 상세: docs/FTC_전송_마스킹_연계_검토_및_로드맵.md

[설명가능·근거·검증가능] 처리 후 무엇을/왜 마스킹했는지 요약 창으로 보여주고,
       메타데이터만(셀 본문 제외) 감사로그(JSONL)에 남긴다.

마스킹 로직은 공용 엔진(masking_engine)을 그대로 사용 — GUI/웹/트레이와 동일.
"""
import os
import sys
import json
import socket
import getpass
from datetime import datetime, timezone, timedelta

from masking_engine import (
    read_table, analyze_dataframe, mask_dataframe,
    write_workbook, write_csv, count_detections, set_rules_path,
    list_rules, add_column_rule, remove_column_rule, MODE_LABEL,
)

KST = timezone(timedelta(hours=9))


# ── 데이터/로그 경로 (app_tray._data_dir() 와 동일 규칙) ──────────
def _data_dir():
    """설치형(frozen)은 사용자 '문서' 폴더, 개발 시엔 스크립트 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.expanduser("~"), "Documents", "엑셀파일_개인정보_마스킹")
    return os.path.dirname(os.path.abspath(__file__))


AUDIT_DIR = os.path.join(_data_dir(), "감사로그")
os.makedirs(AUDIT_DIR, exist_ok=True)

# 조직 학습형 규칙 파일(사용자가 지정한 '항상 마스킹' 컬럼) — GUI와 같은 파일 공유
set_rules_path(os.path.join(_data_dir(), "masking_rules.json"))


def _unique(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}({i}){ext}"):
        i += 1
    return f"{base}({i}){ext}"


# ── 감사로그 (메타데이터만 · 셀 본문 절대 기록 안 함) ─────────────
# 기존 docs/사용이력_로깅_도입안.md 의 원칙(JSONL·메타데이터만)과 일관.
def _audit(event, *, src="", dst="", status="ok", detail="",
           masked_cnt=0, cols=None, types=None):
    rec = {
        "ts": datetime.now(KST).isoformat(timespec="seconds"),
        "event": event,                       # mask_export / none / error
        "user": getpass.getuser(),            # 윈도우 로그인 계정
        "hostname": socket.gethostname(),
        "src_name": os.path.basename(src),    # 파일명만 (경로/본문 제외)
        "dst_name": os.path.basename(dst),
        "status": status,
        "detail": detail,
        "masked_cnt": masked_cnt,             # 총 마스킹 건수
        "cols": cols or [],                   # 마스킹된 컬럼명(메타데이터)
        "types": types or {},                 # {유형: 건수}
    }
    day = datetime.now(KST).strftime("%Y-%m")
    path = os.path.join(AUDIT_DIR, f"masking-{day}.jsonl")
    try:
        with open(path, "a", encoding="utf-8") as fp:
            fp.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass   # 로깅 실패가 마스킹 자체를 막지 않도록


# ── 결과 요약 창 (설명가능·검증가능) ─────────────────────────────
def _dialog(kind, title, msg):
    """작은 안내 창. 콘솔 없이(pythonw) 실행돼도 동작."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        {"info": messagebox.showinfo,
         "warn": messagebox.showwarning,
         "error": messagebox.showerror}.get(kind, messagebox.showinfo)(title, msg)
        root.destroy()
    except Exception:
        print(f"[{title}] {msg}")


def _summ_types(report_rows):
    """report_rows(헤더 포함) → {유형: 건수} 와 컬럼 목록."""
    types, cols = {}, []
    for r in report_rows[1:]:
        col, typ, _grade, n, *_ = r
        if isinstance(n, int) and n > 0:
            types[typ] = types.get(typ, 0) + n
            if col not in cols and col != "—":
                cols.append(col)
    return types, cols


# ── 핵심: 파일 하나를 마스킹본으로 만든다 ─────────────────────────
def mask_file(path):
    name = os.path.basename(path)

    low = name.lower()
    if low.endswith(("_마스킹.xlsx", "_마스킹.csv")):
        _dialog("warn", "이미 마스킹된 사본",
                f"'{name}' 은(는) 이미 마스킹된 사본입니다.\n"
                "원본 파일에서 마우스 오른쪽 버튼을 눌러 주세요.")
        return 2

    if not low.endswith((".xlsx", ".xls", ".csv")):
        _dialog("warn", "지원하지 않는 형식",
                "엑셀/CSV 파일(.xlsx / .xls / .csv)만 마스킹할 수 있습니다.")
        return 2

    is_csv = low.endswith(".csv")

    try:
        table = read_table(path)
    except Exception as e:
        _audit("error", src=path, status="error", detail=str(e))
        _dialog("error", "파일을 읽을 수 없음",
                f"'{name}' 을(를) 열 수 없습니다.\n\n{e}\n\n"
                "· 파일이 Excel에서 열려 있으면 닫고 다시 시도하세요.\n"
                "· 문서보안(DRM) 파일이면 .xlsx로 다시 저장 후 시도하세요.")
        return 1

    targets = [c["name"] for c in analyze_dataframe(table) if c["suggest"]]

    if not targets:
        _audit("none", src=path, status="ok", detail="민감정보 미탐지")
        _dialog("warn", "민감정보 미탐지",
                f"'{name}' 에서 마스킹할 개인정보를 찾지 못했습니다.\n\n"
                "마스킹본을 만들지 않았습니다. 원본을 직접 확인한 뒤\n"
                "반출 여부를 판단하세요. (자동 판정이므로 누락 가능성 있음)")
        return 0

    result, report_rows, ref_rows = mask_dataframe(table, targets)
    ext = "_마스킹.csv" if is_csv else "_마스킹.xlsx"   # 입력 형식 유지
    dst = _unique(os.path.join(os.path.dirname(path),
                               os.path.splitext(name)[0] + ext))
    if is_csv:
        write_csv(dst, result)                          # CSV는 단일 표로 저장
    else:
        write_workbook(dst, result, report_rows, ref_rows)

    total = count_detections(report_rows)
    types, cols = _summ_types(report_rows)
    _audit("mask_export", src=path, dst=dst, status="ok",
           masked_cnt=total, cols=cols, types=types)

    detail = "\n".join(f"   · {t} × {n}건" for t, n in types.items()) or "   · (건수 0)"
    view = "" if is_csv else "\n※ 결과 파일의 [분류리포트] 시트에서\n   무엇을 왜 가렸는지 직접 확인하세요."
    _dialog("info", "마스킹 사본 생성 완료",
            f"✅ 개인정보를 가린 사본을 만들었습니다.\n\n"
            f"만들어진 파일:\n{os.path.basename(dst)}\n"
            f"(원본과 같은 폴더에 새로 생겼어요)\n\n"
            f"가린 개인정보 (총 {total}건):\n{detail}\n\n"
            f"▶ FTC로는 이 '_마스킹' 파일을 내보내세요."
            f"{view}")
    return 0


# ── 규칙 조작 명령 (AI 비서가 붙는 자리) ─────────────────────────
# 직원이 "고객번호는 뒤 4자리는 남겨줘" 라고 말하면, AI 비서는 아래 명령 한 줄만 실행한다.
#   python mask_cli.py rules add 고객번호 last4
# 엑셀 '내용'은 AI에게 전혀 넘어가지 않는다 — AI는 '규칙'만 만들고,
# 실제 마스킹은 지금까지처럼 이 PC 안의 엔진이 수행한다.
# 결과를 JSON으로 출력하는 이유: 사람도 읽을 수 있고, AI도 그대로 해석할 수 있어서.

def _out(obj):
    """결과를 UTF-8 JSON 한 덩어리로 출력(한글이 깨지지 않게)."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")   # 윈도우 콘솔 한글 깨짐 방지
    except Exception:
        pass
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def rules_command(argv):
    """rules list | add <칸이름> <방식> | remove <칸이름>  → 성공 0, 실패 2"""
    action = argv[0] if argv else "list"

    if action == "list":
        _out({"ok": True,
              "rules": list_rules(),                 # 지금 저장된 규칙들
              "modes": MODE_LABEL})                   # 고를 수 있는 방식 목록(코드: 설명)
        return 0

    if action == "add":
        if len(argv) < 3:
            _out({"ok": False,
                  "error": "사용법: rules add <칸이름> <방식>",
                  "modes": MODE_LABEL})
            return 2
        header, mode = argv[1], argv[2]
        if mode not in MODE_LABEL:                    # 없는 방식이면 조용히 'full'로 바뀌지 않도록 먼저 막는다
            _out({"ok": False,
                  "error": f"'{mode}' 는 없는 방식입니다.",
                  "modes": MODE_LABEL})
            return 2
        add_column_rule(header, mode=mode)
        _out({"ok": True,
              "message": f"‘{header}’ 칸을 앞으로 ‘{MODE_LABEL[mode]}’ 방식으로 가립니다.",
              "rules": list_rules()})
        return 0

    if action == "remove":
        if len(argv) < 2:
            _out({"ok": False, "error": "사용법: rules remove <칸이름>"})
            return 2
        header = argv[1]
        before = {r.get("match") for r in list_rules()}
        if header not in before:
            _out({"ok": False, "error": f"‘{header}’ 규칙이 없습니다.",
                  "rules": list_rules()})
            return 2
        remove_column_rule(header)
        _out({"ok": True, "message": f"‘{header}’ 규칙을 지웠습니다.",
              "rules": list_rules()})
        return 0

    _out({"ok": False, "error": f"모르는 명령: {action}",
          "usage": ["rules list", "rules add <칸이름> <방식>", "rules remove <칸이름>"]})
    return 2


def main():
    # 첫 인자가 'rules' 면 규칙 관리 모드 (그 외에는 지금까지처럼 파일 마스킹)
    if len(sys.argv) >= 2 and sys.argv[1] == "rules":
        return rules_command(sys.argv[2:])

    if len(sys.argv) < 2:
        _dialog("warn", "사용법",
                "엑셀 파일을 우클릭 → '개인정보 마스킹본 만들기' 로 실행하세요.\n"
                "(또는 명령행: mask_cli.py <엑셀파일경로>)")
        return 2
    # 여러 파일을 한 번에 선택해 우클릭한 경우도 처리
    rc = 0
    for arg in sys.argv[1:]:
        rc = mask_file(arg) or rc
    return rc


if __name__ == "__main__":
    sys.exit(main())
