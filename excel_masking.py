"""
엑셀파일 개인정보 마스킹 — 데스크톱 GUI
────────────────────────────────────────────────────────────
마스킹 로직은 공용 엔진(masking_engine.py)에 있으며 웹서버(web/app.py)와 공유한다.
이 파일은 tkinter 화면만 담당한다.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading

import make_manual

from masking_engine import (
    STANDARD_NAME, STANDARD_TIMELINE, GUIDE_NOTE, TREND_NOTES,
    GRADE_LABEL, GRADE_COLOR, GRADE_POLICY, GRADE_DEF, RISK_HIGH,
    read_table, analyze_dataframe, mask_dataframe, write_workbook, count_detections,
    MODE_LABEL, list_rules, add_column_rule, remove_column_rule, preview_mask,
)


def _risk_style(pct):
    """오탐 추정 %에 따른 (배경색, 글자색)."""
    if pct >= RISK_HIGH:
        return "#fee2e2", "#b91c1c"      # 높음
    if pct >= 15:
        return "#fef3c7", "#92400e"      # 보통
    return "#dcfce7", "#166534"          # 낮음


class MaskingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("엑셀파일 개인정보 마스킹")
        self.geometry("880x720")
        self.minsize(820, 640)
        self.configure(bg="#f4f6f9")

        self.filepath = tk.StringVar()
        self.table = None
        self.col_checks = {}   # {col: BooleanVar}
        self.ack_var = tk.BooleanVar(value=False)   # 확인 후 사용(게이트)
        self._build_ui()

    # ── UI 빌드 ───────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg="#1e3a5f", pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔒  엑셀파일 개인정보 마스킹",
                 font=("맑은 고딕", 15, "bold"), fg="white", bg="#1e3a5f").pack()
        tk.Label(hdr, text="개인정보(이름·전화·주민번호 등)를 별표(*)로 가려 안전한 사본을 만듭니다",
                 font=("맑은 고딕", 9), fg="#bcd0e8", bg="#1e3a5f").pack(pady=(2, 0))

        # 근거 바 (외부 링크 없이 안내만)
        srcbar = tk.Frame(self, bg="#eef2f7")
        srcbar.pack(fill="x")
        tk.Label(srcbar, text="근거:", font=("맑은 고딕", 8, "bold"),
                 fg="#475569", bg="#eef2f7").pack(side="left", padx=(18, 4), pady=4)
        tk.Button(srcbar, text="ℹ️ N2SF 기준·근거 보기", command=self._show_sources,
                  bg="#dbeafe", fg="#1e3a5f", relief="flat",
                  font=("맑은 고딕", 8, "bold"), padx=8).pack(side="left", padx=2, pady=2)
        tk.Label(srcbar, text="원문: 국정원(국가사이버안보센터) 홈페이지에서 확인",
                 font=("맑은 고딕", 8), fg="#64748b", bg="#eef2f7").pack(side="left", padx=6)
        tk.Button(srcbar, text="📖 사용방법", command=self._open_manual,
                  bg="#16a34a", fg="white", relief="flat",
                  font=("맑은 고딕", 8, "bold"), padx=10).pack(side="right", padx=(2, 18), pady=2)

        # ① 파일 선택
        frm = tk.LabelFrame(self, text=" ① 파일 선택 ", font=("맑은 고딕", 10, "bold"),
                            bg="#f4f6f9", padx=10, pady=8)
        frm.pack(fill="x", padx=18, pady=(12, 6))
        tk.Entry(frm, textvariable=self.filepath, font=("맑은 고딕", 10),
                 width=58, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(frm, text="파일 열기", command=self._open_file,
                  bg="#2563eb", fg="white", font=("맑은 고딕", 10),
                  relief="flat", padx=12, pady=4).pack(side="left")

        # ② 분류 결과
        outer = tk.LabelFrame(self, text=" ② 자동 분류 결과 (체크한 칸만 마스킹) ",
                              font=("맑은 고딕", 10, "bold"), bg="#f4f6f9", padx=10, pady=8)
        outer.pack(fill="both", expand=True, padx=18, pady=6)
        topbar = tk.Frame(outer, bg="#f4f6f9")
        topbar.pack(fill="x", pady=(0, 4))
        self.summary_var = tk.StringVar(value="파일을 열면 등급별 분류 결과가 표시됩니다.")
        tk.Label(topbar, textvariable=self.summary_var, font=("맑은 고딕", 9, "bold"),
                 fg="#334155", bg="#f4f6f9").pack(side="left")
        tk.Button(topbar, text="전체 선택/해제", command=self._toggle_all,
                  bg="#e2e8f0", relief="flat", font=("맑은 고딕", 8), padx=6).pack(side="right")
        tk.Button(topbar, text="📌 항상 마스킹할 칸 정하기", command=self._manage_rules,
                  bg="#dbeafe", fg="#1e3a5f", relief="flat",
                  font=("맑은 고딕", 8, "bold"), padx=6).pack(side="right", padx=4)

        self.canvas = tk.Canvas(outer, bg="#ffffff", highlightthickness=1,
                                highlightbackground="#e2e8f0")
        sb = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.list_frame = tk.Frame(self.canvas, bg="#ffffff")
        self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.placeholder = tk.Label(self.list_frame, text="아직 불러온 파일이 없습니다.",
                                    font=("맑은 고딕", 10), fg="#94a3b8", bg="#ffffff")
        self.placeholder.pack(pady=30)

        # ③ 오탐 가능성 안내 (추정·근거) — 유형별 한 줄씩
        nf = tk.LabelFrame(self, text=" 오탐 가능성 안내 (추정·근거) ",
                           font=("맑은 고딕", 9, "bold"), bg="#fffbeb",
                           fg="#92400e", padx=10, pady=6)
        nf.pack(fill="x", padx=18, pady=(0, 6))
        self.notice_inner = tk.Frame(nf, bg="#fffbeb")
        self.notice_inner.pack(fill="x")
        tk.Label(self.notice_inner, text="파일을 열면 표시됩니다.", font=("맑은 고딕", 8),
                 fg="#a16207", bg="#fffbeb", anchor="w").pack(anchor="w")

        # ④ 확인 체크 (자체 줄)
        ackrow = tk.Frame(self, bg="#eef2f7")
        ackrow.pack(fill="x", padx=18)
        tk.Checkbutton(ackrow, text="  분류 결과와 오탐 가능성을 확인했습니다 (체크해야 실행 가능)",
                       variable=self.ack_var, command=self._gate,
                       font=("맑은 고딕", 9, "bold"), fg="#334155",
                       bg="#eef2f7", activebackground="#eef2f7").pack(side="left", pady=4)

        # ⑤ 상태 + 실행 버튼 (자체 줄)
        btm = tk.Frame(self, bg="#f4f6f9")
        btm.pack(fill="x", padx=18, pady=(4, 12))
        self.status_var = tk.StringVar(value="파일을 선택해 주세요.")
        tk.Label(btm, textvariable=self.status_var, font=("맑은 고딕", 9),
                 fg="#64748b", bg="#f4f6f9").pack(side="left")
        self.run_btn = tk.Button(btm, text="   마스킹 실행 및 저장   ", command=self._run_masking,
                                 bg="#16a34a", fg="white", font=("맑은 고딕", 11, "bold"),
                                 relief="flat", padx=18, pady=8, state="disabled",
                                 disabledforeground="#dcfce7")
        self.run_btn.pack(side="right")

    def _gate(self):
        self.run_btn.config(state="normal" if self.ack_var.get() else "disabled")

    # ── N2SF 기준·근거 팝업 (설명가능성/신뢰성, 외부 링크 없음) ──
    def _show_sources(self):
        win = tk.Toplevel(self)
        win.title("N2SF 기준 · 근거")
        win.geometry("640x540")
        win.configure(bg="#ffffff")

        tk.Label(win, text="(참고) 국가 망 보안체계(N2SF) 등급 기준",
                 font=("맑은 고딕", 13, "bold"), fg="#1e3a5f", bg="#ffffff").pack(pady=(16, 2))
        tk.Label(win, text="※ 이 도구의 핵심은 '마스킹'이며, N2SF는 향후 전환 대비 참고 분류입니다.",
                 font=("맑은 고딕", 8), fg="#94a3b8", bg="#ffffff").pack()
        tk.Label(win, text=STANDARD_NAME, font=("맑은 고딕", 9), fg="#475569",
                 bg="#ffffff").pack(pady=(8, 0))
        tk.Label(win, text=STANDARD_TIMELINE, font=("맑은 고딕", 9), fg="#64748b",
                 bg="#ffffff").pack(pady=(0, 10))

        for g in ("C", "S", "O"):
            row = tk.Frame(win, bg="#ffffff")
            row.pack(fill="x", padx=24, pady=3)
            tk.Label(row, text=f" {g} ", bg=GRADE_COLOR[g], fg="white",
                     font=("맑은 고딕", 10, "bold"), width=3).pack(side="left")
            tk.Label(row, text=f" {GRADE_LABEL[g]} · {GRADE_POLICY[g]}",
                     font=("맑은 고딕", 9, "bold"), bg="#ffffff", anchor="w").pack(side="left")
            tk.Label(row, text=GRADE_DEF[g], font=("맑은 고딕", 8), fg="#64748b",
                     bg="#ffffff", anchor="w", wraplength=560, justify="left").pack(
                side="left", padx=6)

        tk.Label(win, text="정식 근거 (원문)", font=("맑은 고딕", 10, "bold"),
                 fg="#1e3a5f", bg="#ffffff").pack(pady=(14, 4))
        tk.Label(win, text=f"「{STANDARD_NAME}」\n{GUIDE_NOTE}",
                 font=("맑은 고딕", 9), fg="#475569", bg="#ffffff",
                 justify="center").pack()

        tk.Label(win, text="참고 동향정보 (정식 공문 아님)", font=("맑은 고딕", 10, "bold"),
                 fg="#64748b", bg="#ffffff").pack(pady=(14, 4))
        for note in TREND_NOTES:
            tk.Label(win, text="• " + note, font=("맑은 고딕", 8), fg="#64748b",
                     bg="#ffffff", anchor="w", wraplength=580,
                     justify="left").pack(fill="x", padx=24)

        tk.Label(win, text="※ 전국 일괄 시행일은 단일 날짜로 공표되지 않았으며 단계적으로 적용됩니다.",
                 font=("맑은 고딕", 8), fg="#94a3b8", bg="#ffffff",
                 wraplength=580).pack(pady=(14, 6), padx=24)
        tk.Button(win, text="닫기", command=win.destroy, bg="#e2e8f0",
                  relief="flat", padx=16, pady=4).pack(pady=8)

    # ── 사용설명서 열기 (프로그램 내장) ──────────────────────────
    def _open_manual(self):
        try:
            make_manual.open_manual()   # HTML을 새로 만들어 기본 브라우저로 연다
        except Exception as e:
            messagebox.showerror("오류", f"사용설명서를 열 수 없습니다.\n{e}")

    # ── 항상 마스킹 규칙 관리 (조직 학습형) ──────────────────────
    def _manage_rules(self):
        win = tk.Toplevel(self)
        win.title("항상 마스킹할 칸 정하기")
        win.geometry("560x520")
        win.configure(bg="#ffffff")

        tk.Label(win, text="📌 항상 마스킹할 칸 정하기",
                 font=("맑은 고딕", 13, "bold"), fg="#1e3a5f", bg="#ffffff").pack(pady=(14, 2))
        tk.Label(win, text="자동으로 안 잡히는 사내 고유 칸(계약번호·상호 등)을\n"
                           "한 번 지정하면 다음부터 항상 자동으로 마스킹됩니다.",
                 font=("맑은 고딕", 8), fg="#64748b", bg="#ffffff", justify="center").pack()

        # 현재 규칙 목록
        listwrap = tk.LabelFrame(win, text=" 지금까지 정한 칸 ", font=("맑은 고딕", 9, "bold"),
                                 bg="#ffffff", padx=8, pady=6)
        listwrap.pack(fill="both", expand=True, padx=16, pady=(10, 6))

        label2mode = {v: k for k, v in MODE_LABEL.items()}

        def refresh():
            for w in listwrap.winfo_children():
                w.destroy()
            rules = list_rules()
            if not rules:
                tk.Label(listwrap, text="아직 등록된 규칙이 없습니다.", font=("맑은 고딕", 9),
                         fg="#94a3b8", bg="#ffffff").pack(pady=16)
            for r in rules:
                row = tk.Frame(listwrap, bg="#f8fafc")
                row.pack(fill="x", pady=1)
                tk.Label(row, text=f"‘{r.get('match','')}’ 칸", width=22, anchor="w",
                         font=("맑은 고딕", 9, "bold"), bg="#f8fafc").pack(side="left", padx=4)
                tk.Label(row, text="→ " + MODE_LABEL.get(r.get("mode", "full"), "전체 가림"),
                         width=14, anchor="w", font=("맑은 고딕", 8),
                         fg="#475569", bg="#f8fafc").pack(side="left")
                # 예전에 정해 둔 규칙도 '어떻게 보이는지' 바로 알 수 있게 예시를 붙인다
                ex = sample_value(r.get("match", ""))
                if ex:
                    tk.Label(row, text=f"{ex[:12]} → {preview_mask(ex, r.get('mode','full'))[:12]}",
                             anchor="w", font=("맑은 고딕", 8),
                             fg="#16a34a", bg="#f8fafc").pack(side="left", padx=2)
                tk.Button(row, text="삭제", command=lambda m=r.get("match"): (
                              remove_column_rule(m), refresh()),
                          bg="#fee2e2", fg="#b91c1c", relief="flat",
                          font=("맑은 고딕", 8), padx=6).pack(side="right", padx=4)

        # 추가 폼
        addfrm = tk.LabelFrame(win, text=" 새로 정하기 ", font=("맑은 고딕", 9, "bold"),
                               bg="#ffffff", padx=8, pady=8)
        addfrm.pack(fill="x", padx=16, pady=(0, 10))

        tk.Label(addfrm, text="칸(열)", font=("맑은 고딕", 8), bg="#ffffff").grid(row=0, column=0, padx=2)
        cols = list(self.table.headers) if self.table else []
        col_var = tk.StringVar(value=cols[0] if cols else "")
        if cols:
            col_widget = ttk.Combobox(addfrm, textvariable=col_var, values=cols,
                                      width=22, state="readonly")
        else:
            col_widget = tk.Entry(addfrm, textvariable=col_var, width=24)  # 파일 없으면 직접 입력
        col_widget.grid(row=0, column=1, padx=4)

        tk.Label(addfrm, text="방식", font=("맑은 고딕", 8), bg="#ffffff").grid(row=0, column=2, padx=2)
        mode_var = tk.StringVar(value=MODE_LABEL["full"])
        ttk.Combobox(addfrm, textvariable=mode_var, values=list(MODE_LABEL.values()),
                     width=14, state="readonly").grid(row=0, column=3, padx=4)

        # ── 결과 미리보기 ────────────────────────────────────────
        # 방식 이름('숫자만 가림')만 봐서는 결과가 어떻게 될지 알 수 없다.
        # 그래서 지금 열린 파일에서 그 칸의 '진짜 값'을 하나 가져와,
        # 고른 방식을 적용하면 어떻게 보이는지 그 자리에서 보여준다.
        preview = tk.Label(addfrm, text="", font=("맑은 고딕", 9), anchor="w",
                           bg="#ffffff", justify="left")
        preview.grid(row=1, column=0, columnspan=5, sticky="w", pady=(8, 0))

        def sample_value(header):
            """열린 파일에서 그 칸의 비어 있지 않은 첫 값 하나. 없으면 None."""
            if self.table is None or header not in self.table.headers:
                return None
            for v in self.table.column(header):
                if v is not None and str(v).strip() not in ("", "-"):
                    return str(v)
            return None

        def refresh_preview(*_):
            """칸이나 방식을 바꿀 때마다 '원래값 → 가린값' 을 다시 그린다."""
            val = sample_value(col_var.get().strip())
            if not val:
                preview.config(text="미리보기: 파일을 열면 이 칸의 실제 값으로 보여드립니다.",
                               fg="#94a3b8")
                return
            after = preview_mask(val, label2mode.get(mode_var.get(), "full")) or val
            preview.config(text=f"미리보기:   {val[:22]}   →   {after[:22]}", fg="#16a34a")

        col_var.trace_add("write", refresh_preview)
        mode_var.trace_add("write", refresh_preview)
        refresh_preview()

        def do_add():
            header = col_var.get().strip()
            if not header:
                messagebox.showwarning("알림", "컬럼명을 입력/선택해 주세요.", parent=win)
                return
            add_column_rule(header, mode=label2mode.get(mode_var.get(), "full"))
            refresh()
            if self.table is not None:      # 열린 파일이 있으면 즉시 반영
                self._analyze_and_render()

        tk.Button(addfrm, text="＋ 추가", command=do_add, bg="#16a34a", fg="white",
                  relief="flat", font=("맑은 고딕", 9, "bold"), padx=10).grid(
            row=0, column=4, padx=6)

        tk.Label(win, text="※ 규칙은 masking_rules.json에 저장되어 우클릭·자동감시에도 함께 적용됩니다.",
                 font=("맑은 고딕", 8), fg="#94a3b8", bg="#ffffff", wraplength=520).pack(padx=16)
        tk.Button(win, text="닫기", command=win.destroy, bg="#e2e8f0",
                  relief="flat", padx=16, pady=4).pack(pady=8)
        refresh()

    # ── 파일 열기 ─────────────────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        try:
            self.table = read_table(path)
            self.filepath.set(path)
            self.status_var.set(
                f"✅ {self.table.nrows}행 × {self.table.ncols}열 로드 완료 · 분류 중…")
            self.update()
            self._analyze_and_render()
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽을 수 없습니다.\n\n{e}")

    # ── 분석 + 목록 렌더링 ────────────────────────────────────
    def _analyze_and_render(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        self.col_checks.clear()

        c_cnt = s_cnt = 0
        notice = {}   # 유형 → (최고 오탐%, 근거)
        for i, c in enumerate(analyze_dataframe(self.table)):
            col, grade, counts = c["name"], c["grade"], c["types"]
            before, after = c["before"], c["after"]
            if grade == "C":
                c_cnt += 1
            elif grade == "S":
                s_cnt += 1
            for it in c["risk_items"]:
                if it["type"] not in notice or it["pct"] > notice[it["type"]][0]:
                    notice[it["type"]] = (it["pct"], it["reason"])
            check_var = tk.BooleanVar(value=c["suggest"])
            self.col_checks[col] = check_var
            self._render_row(i, col, grade, counts, before, after, check_var, c["risk"])

        total = self.table.ncols
        self.summary_var.set(
            f"총 {total}열 — 기밀(C) {c_cnt} · 민감(S) {s_cnt} · 공개(O) {total - c_cnt - s_cnt}")
        for w in self.notice_inner.winfo_children():
            w.destroy()
        if notice:
            for t, (p, r) in sorted(notice.items(), key=lambda kv: -kv[1][0]):
                line = tk.Frame(self.notice_inner, bg="#fffbeb")
                line.pack(fill="x", anchor="w", pady=1)
                rbg, rfg = _risk_style(p)
                tk.Label(line, text=f"오탐 {p}%", bg=rbg, fg=rfg, width=8,
                         font=("맑은 고딕", 8, "bold")).pack(side="left", padx=(0, 8))
                tk.Label(line, text=f"{t} — {r}", bg="#fffbeb", fg="#7c2d12",
                         font=("맑은 고딕", 8), anchor="w", justify="left",
                         wraplength=740).pack(side="left", fill="x")
        else:
            tk.Label(self.notice_inner, text="탐지된 민감정보가 없습니다.",
                     font=("맑은 고딕", 8), fg="#a16207", bg="#fffbeb",
                     anchor="w").pack(anchor="w")
        # 새 파일마다 확인 절차 재요구
        self.ack_var.set(False)
        self._gate()
        self.status_var.set("✅ 분류 완료 — 아래 확인 체크 후 [마스킹 실행 및 저장]")

    def _render_row(self, i, col, grade, counts, before, after, check_var, risk):
        bg = "#ffffff" if i % 2 == 0 else "#f8fafc"
        row = tk.Frame(self.list_frame, bg=bg)
        row.pack(fill="x", padx=2, pady=1)
        tk.Checkbutton(row, variable=check_var, bg=bg).pack(side="left", padx=(4, 2))
        tk.Label(row, text=f" {grade} ", bg=GRADE_COLOR[grade], fg="white",
                 font=("맑은 고딕", 9, "bold"), width=3).pack(side="left", padx=4)
        tk.Label(row, text=str(col), width=14, anchor="w", font=("맑은 고딕", 9, "bold"),
                 bg=bg).pack(side="left", padx=2)
        types = ", ".join(f"{t}×{n}" for t, n in counts.items()) if counts else "—"
        tk.Label(row, text=types, width=22, anchor="w", font=("맑은 고딕", 8),
                 fg="#475569", bg=bg).pack(side="left", padx=2)
        # 오탐 추정 배지
        if risk:
            rbg, rfg = _risk_style(risk)
            tk.Label(row, text=f"오탐 {risk}%", bg=rbg, fg=rfg,
                     font=("맑은 고딕", 8, "bold"), width=8).pack(side="left", padx=2)
        else:
            tk.Label(row, text="—", width=8, font=("맑은 고딕", 8),
                     fg="#cbd5e1", bg=bg).pack(side="left", padx=2)
        ex = f"{before[:12]} → {after[:16]}" if before else "(마스킹 없음)"
        tk.Label(row, text=ex, anchor="w", font=("맑은 고딕", 8),
                 fg="#16a34a" if before else "#94a3b8", bg=bg).pack(side="left", padx=2)

    def _toggle_all(self):
        if not self.col_checks:
            return
        new = not all(v.get() for v in self.col_checks.values())
        for v in self.col_checks.values():
            v.set(new)

    # ── 마스킹 실행 ───────────────────────────────────────────
    def _run_masking(self):
        if self.table is None:
            messagebox.showwarning("알림", "먼저 파일을 선택해 주세요.")
            return
        targets = [c for c, v in self.col_checks.items() if v.get()]
        if not targets:
            messagebox.showwarning("알림", "마스킹할 컬럼을 하나 이상 선택해 주세요.")
            return
        base, _ = os.path.splitext(self.filepath.get())
        save_path = filedialog.asksaveasfilename(
            title="저장 위치 선택",
            initialfile=os.path.basename(base + "_마스킹.xlsx"),
            defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if not save_path:
            return
        self.status_var.set("⏳ 마스킹 처리 중…")
        self.update()
        threading.Thread(target=self._worker, args=(targets, save_path), daemon=True).start()

    def _worker(self, targets, save_path):
        try:
            result, report_rows, ref_rows = mask_dataframe(self.table, targets)
            write_workbook(save_path, result, report_rows, ref_rows)

            total = count_detections(report_rows)
            self.after(0, lambda: (
                self.status_var.set(f"✅ 저장 완료 → {os.path.basename(save_path)}"),
                messagebox.showinfo("완료",
                    f"마스킹이 완료되었습니다!\n\n"
                    f"· 마스킹 건수: 총 {total}건\n"
                    f"· 시트: [마스킹결과] + [분류리포트] + [N2SF근거]\n\n"
                    f"저장 위치:\n{save_path}")))
        except Exception as e:
            self.after(0, lambda: (
                self.status_var.set("❌ 오류 발생"),
                messagebox.showerror("오류", f"처리 중 오류가 발생했습니다.\n{e}")))


if __name__ == "__main__":
    MaskingApp().mainloop()
