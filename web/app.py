"""
N2SF 마스킹 웹서버
────────────────────────────────────────────────────────────
이 PC에서 구동하고, 같은 망의 다른 PC가 브라우저로 접속(로컬 IP)하는 구조.
마스킹 로직은 상위 폴더의 공용 엔진(masking_engine.py)을 그대로 사용한다.

실행:
  python app.py                  → http://<이 PC IP>:5000
  (업무망 이관 시 production)     → waitress-serve --listen=0.0.0.0:5000 app:app
"""
import io
import os
import re
import sys
import glob
import json
import uuid
import socket
import zipfile
import tempfile
from collections import Counter

from flask import (Flask, request, jsonify, send_file, render_template,
                   session, redirect, url_for)

# 상위 폴더의 공용 엔진 import
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from masking_engine import (  # noqa: E402
    read_table, analyze_dataframe, mask_dataframe, write_workbook)
from usage_logger import log_event, LOG_DIR  # noqa: E402

def _resource(rel):
    """exe(PyInstaller)로 묶이면 임시 추출 폴더, 아니면 스크립트 폴더 기준 경로."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


app = Flask(__name__, template_folder=_resource("templates"))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 업로드 50MB 제한
app.secret_key = os.urandom(24)                        # 관리자 세션용

ALLOWED_EXT = (".xlsx", ".xls")
TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")
TMP_DIR = os.path.join(tempfile.gettempdir(), "n2sf_masking")
os.makedirs(TMP_DIR, exist_ok=True)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# 관리자 비밀번호 (환경변수 N2SF_ADMIN_PW로 변경 가능, 기본값 제공)
ADMIN_PW = os.environ.get("N2SF_ADMIN_PW", "n2sf-admin")


def _token_path(token):
    """토큰 검증 후 임시 파일 경로 반환(경로 조작 방지)."""
    if not token or not TOKEN_RE.match(token):
        return None
    return os.path.join(TMP_DIR, token + ".xlsx")   # openpyxl이 확장자 검증함


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/analyze")
def analyze():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(ALLOWED_EXT):
        return jsonify(error="엑셀 파일(.xlsx/.xls)을 올려주세요."), 400
    try:
        data = f.read()
        table = read_table(io.BytesIO(data))
    except Exception as e:
        log_event("analyze", client_ip=request.remote_addr,
                  user=request.form.get("user", ""),
                  filename=getattr(f, "filename", ""), status="error",
                  detail="파일 읽기 실패")
        return jsonify(error=f"파일을 읽을 수 없습니다: {e}"), 400

    token = uuid.uuid4().hex
    with open(_token_path(token), "wb") as fp:    # 처리 사이 임시 보관(분석→마스킹)
        fp.write(data)

    columns = analyze_dataframe(table)
    for c in columns:                              # types dict → "유형×n" 표시용 문자열
        c["types_text"] = ", ".join(f"{t}×{n}" for t, n in c["types"].items()) or "—"
    c_cnt = sum(1 for c in columns if c["grade"] == "C")
    s_cnt = sum(1 for c in columns if c["grade"] == "S")

    log_event("analyze",
              user=request.form.get("user", ""),
              client_ip=request.remote_addr,
              req_id=token,
              filename=f.filename,
              rows=table.nrows, cols_total=len(columns), cols_c=c_cnt, cols_s=s_cnt)
    return jsonify(token=token, filename=f.filename, rows=table.nrows,
                   total=len(columns), c=c_cnt, s=s_cnt, columns=columns,
                   risk_max=max((c["risk"] for c in columns), default=0))


@app.post("/discard")
def discard():
    """초기화: 마스킹 전 임시 파일을 서버에서 삭제."""
    payload = request.get_json(silent=True) or {}
    for token in payload.get("tokens", []):
        path = _token_path(token)
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass
    return jsonify(ok=True)


def _mask_one(item, user, client_ip):
    """단일 파일 마스킹 → (download_name, xlsx_bytes) 또는 (None, 오류메시지)."""
    path = _token_path(item.get("token"))
    if not path or not os.path.exists(path):
        return None, "세션이 만료되었습니다. 파일을 다시 올려주세요."
    selected = [str(s) for s in item.get("selected", [])]
    if not selected:
        return None, "마스킹할 컬럼을 하나 이상 선택해 주세요."

    try:
        table = read_table(path)
        result, report_rows, ref_rows = mask_dataframe(table, selected)
        bio = io.BytesIO()
        write_workbook(bio, result, report_rows, ref_rows)
    except Exception as e:
        return None, f"파일을 처리할 수 없습니다: {e}"

    try:
        os.remove(path)                            # 처리 후 즉시 삭제(서버에 잔존 X)
    except OSError:
        pass

    log_event("mask", user=(user or ""), client_ip=client_ip,
              req_id=item.get("token", ""), filename=item.get("filename", ""),
              selected_cnt=len(selected))

    base = os.path.splitext(os.path.basename(item.get("filename", "결과")))[0]
    return f"{base}_마스킹.xlsx", bio.getvalue()


@app.post("/mask")
def mask():
    payload = request.get_json(silent=True) or {}
    user = payload.get("user") or ""
    # 단일/다중 모두 지원: items 리스트가 있으면 다중, 없으면 단일 호환
    items = payload.get("items")
    if not items:
        items = [{"token": payload.get("token"), "filename": payload.get("filename"),
                  "selected": payload.get("selected", [])}]
    if not items:
        return jsonify(error="처리할 파일이 없습니다."), 400

    results = []
    for item in items:
        name, data = _mask_one(item, user, request.remote_addr)
        if name is None:
            return jsonify(error=data), 400        # 하나라도 실패하면 중단
        results.append((name, data))

    if len(results) == 1:                          # 1개 → 엑셀 그대로
        name, data = results[0]
        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=name, mimetype=XLSX_MIME)

    # 여러 개 → ZIP 묶음
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in results:
            zf.writestr(name, data)
    zbuf.seek(0)
    return send_file(zbuf, as_attachment=True,
                     download_name="마스킹결과.zip", mimetype="application/zip")


# ── 관리자 로그 모니터링 ───────────────────────────────────────

def _read_logs(limit=3000):
    """logs/usage-*.jsonl 들을 최신순으로 읽어 레코드 리스트 반환."""
    recs = []
    for fp in sorted(glob.glob(os.path.join(LOG_DIR, "usage-*.jsonl")), reverse=True):
        try:
            with open(fp, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            recs.append(json.loads(line))
                        except ValueError:
                            pass
        except OSError:
            pass
        if len(recs) >= limit:
            break
    recs.sort(key=lambda r: r.get("ts", ""), reverse=True)
    return recs[:limit]


@app.get("/admin")
def admin_redirect():
    return redirect(url_for("admin_logs"))


@app.post("/admin/login")
def admin_login():
    if (request.form.get("pw") or "") == ADMIN_PW:
        session["admin"] = True
    else:
        return render_template("admin.html", authed=False, error="비밀번호가 올바르지 않습니다.")
    return redirect(url_for("admin_logs"))


@app.get("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin_logs"))


@app.get("/admin/logs")
def admin_logs():
    if not session.get("admin"):
        return render_template("admin.html", authed=False, error="")
    recs = _read_logs()
    today = ""
    try:
        from datetime import datetime, timezone, timedelta
        today = datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d")
    except Exception:
        pass
    stats = {
        "total": len(recs),
        "users": len({r.get("user", "") for r in recs}),
        "today": sum(1 for r in recs if str(r.get("ts", "")).startswith(today)),
        "mask": sum(1 for r in recs if r.get("event") == "mask"),
        "analyze": sum(1 for r in recs if r.get("event") == "analyze"),
        "errors": sum(1 for r in recs if r.get("status") == "error"),
        "by_user": Counter(r.get("user", "(미입력)") for r in recs).most_common(10),
    }
    return render_template("admin.html", authed=True, recs=recs, stats=stats)


@app.errorhandler(413)
def too_large(_e):
    return jsonify(error="파일이 너무 큽니다(최대 50MB)."), 413


@app.errorhandler(Exception)
def _json_errors(e):
    """어떤 오류든 JSON으로 응답(프론트가 항상 파싱 가능 → 'not valid json' 방지)."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return jsonify(error=e.description), e.code
    return jsonify(error=f"서버 처리 오류: {e}"), 500


def _lan_ip():
    """같은 망에서 접속할 이 PC의 IP 추정."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except Exception:
            ip = "127.0.0.1"
    finally:
        s.close()
    return ip


if __name__ == "__main__":
    port = 5000
    ip = _lan_ip()
    print("\n" + "=" * 52)
    print("  N2SF 마스킹 웹서버 실행 중")
    print(f"  - 이 PC에서:    http://127.0.0.1:{port}")
    print(f"  - 같은 망 PC:   http://{ip}:{port}")
    print("  (종료: Ctrl + C)")
    print("=" * 52 + "\n")
    try:
        from waitress import serve            # 운영용 WSGI 서버(권장)
        serve(app, host="0.0.0.0", port=port)
    except ImportError:
        app.run(host="0.0.0.0", port=port)    # 개발용 폴백
