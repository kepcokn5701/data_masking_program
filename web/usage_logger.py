"""사용 이력 로깅 — 행위 메타데이터만 JSONL로 append. (개인정보 본문은 기록 안 함)"""
import os
import sys
import json
import socket
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
# exe(PyInstaller)로 묶이면 로그를 exe 옆에 남긴다(임시 추출폴더는 종료 시 사라지므로).
if getattr(sys, "frozen", False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(_BASE, "logs")
os.makedirs(LOG_DIR, exist_ok=True)


def _hostname(ip):
    """IP → 호스트명 역방향 조회(실패해도 무시)."""
    try:
        return socket.gethostbyaddr(ip)[0]
    except Exception:
        return ""


def log_event(event, *, user="", client_ip="", req_id="", filename="",
              status="ok", detail="", **extra):
    rec = {
        "ts": datetime.now(KST).isoformat(timespec="seconds"),
        "event": event,
        "user": (user or "(미입력)").strip(),
        "client_ip": client_ip,
        "hostname": _hostname(client_ip) if client_ip else "",
        "req_id": req_id,
        "filename": filename,
        "status": status,
        "detail": detail,
    }
    rec.update(extra)  # rows, cols_total, cols_c, cols_s, selected_cnt 등
    day = datetime.now(KST).strftime("%Y-%m-%d")
    path = os.path.join(LOG_DIR, f"usage-{day}.jsonl")
    with open(path, "a", encoding="utf-8") as fp:
        fp.write(json.dumps(rec, ensure_ascii=False) + "\n")
