import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os
import threading

# ── 마스킹 함수들 ──────────────────────────────────────────────

def mask_name(value):
    """이름: 홍*동 (가운데 글자 *)"""
    s = str(value).strip()
    if len(s) == 2:
        return s[0] + "*"
    elif len(s) >= 3:
        return s[0] + "*" * (len(s) - 2) + s[-1]
    return s

def mask_phone(value):
    """전화번호: 010-****-5678"""
    s = re.sub(r"[^\d]", "", str(value))
    if len(s) == 11:
        return f"{s[:3]}-****-{s[7:]}"
    elif len(s) == 10:
        return f"{s[:3]}-***-{s[6:]}"
    return re.sub(r"\d{3,4}(?=\d{4})", "****", str(value))

def mask_email(value):
    """이메일: ho***@gmail.com"""
    s = str(value).strip()
    if "@" in s:
        local, domain = s.split("@", 1)
        keep = max(2, len(local) // 3)
        return local[:keep] + "*" * (len(local) - keep) + "@" + domain
    return s

def mask_rrn(value):
    """주민등록번호: ******-*******"""
    s = re.sub(r"[^\d]", "", str(value))
    if len(s) == 13:
        return "******-*******"
    return re.sub(r"\d", "*", str(value))

def mask_address(value):
    """주소: 서울시 ** ** *** (시/도 뒤 마스킹)"""
    s = str(value).strip()
    pattern = r"((?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)(?:특별시|광역시|특별자치시|도|특별자치도)?)\s*(.*)"
    m = re.match(pattern, s)
    if m:
        region = m.group(1)
        rest = m.group(2)
        masked_rest = re.sub(r"[가-힣a-zA-Z0-9]+", lambda x: "*" * len(x.group()), rest)
        return region + " " + masked_rest
    # 패턴 없으면 첫 단어 유지 후 마스킹
    parts = s.split()
    if len(parts) > 1:
        return parts[0] + " " + " ".join("*" * len(p) for p in parts[1:])
    return s

# ── 컬럼 타입 자동 감지 ────────────────────────────────────────

def detect_column_type(series):
    sample = series.dropna().astype(str).head(20)
    phone_pat = re.compile(r"0\d{1,2}[-\s]?\d{3,4}[-\s]?\d{4}")
    email_pat = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")
    rrn_pat   = re.compile(r"\d{6}[-\s]?\d{7}")
    addr_pat  = re.compile(r"(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)")
    name_pat  = re.compile(r"^[가-힣]{2,4}$")

    scores = {"이름": 0, "전화번호": 0, "이메일": 0, "주민등록번호": 0, "주소": 0}
    for v in sample:
        if phone_pat.search(v):   scores["전화번호"] += 1
        if email_pat.search(v):   scores["이메일"] += 1
        if rrn_pat.search(v):     scores["주민등록번호"] += 1
        if addr_pat.search(v):    scores["주소"] += 1
        if name_pat.match(v):     scores["이름"] += 1

    best = max(scores, key=scores.get)
    return best if scores[best] >= max(2, len(sample) * 0.3) else "없음"

MASK_FUNCS = {
    "이름":        mask_name,
    "전화번호":    mask_phone,
    "이메일":      mask_email,
    "주민등록번호": mask_rrn,
    "주소":        mask_address,
}

# ── GUI ───────────────────────────────────────────────────────

class MaskingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("엑셀 개인정보 마스킹 도구")
        self.geometry("720x560")
        self.resizable(True, True)
        self.configure(bg="#f4f6f9")

        self.filepath = tk.StringVar()
        self.df = None
        self.col_vars = {}   # {col_name: StringVar (마스킹 타입)}
        self.col_checks = {} # {col_name: BooleanVar (선택 여부)}

        self._build_ui()

    # ── UI 빌드 ───────────────────────────────────────────────

    def _build_ui(self):
        # 헤더
        hdr = tk.Frame(self, bg="#2563eb", pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔒  엑셀 개인정보 마스킹 도구",
                 font=("맑은 고딕", 15, "bold"), fg="white", bg="#2563eb").pack()

        # 파일 선택
        frm = tk.LabelFrame(self, text=" 파일 선택 ", font=("맑은 고딕", 10, "bold"),
                             bg="#f4f6f9", padx=10, pady=8)
        frm.pack(fill="x", padx=18, pady=(14, 6))

        tk.Entry(frm, textvariable=self.filepath, font=("맑은 고딕", 10),
                 width=52, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(frm, text="파일 열기", command=self._open_file,
                  bg="#2563eb", fg="white", font=("맑은 고딕", 10),
                  relief="flat", padx=10, pady=4).pack(side="left")

        # 컬럼 설정
        self.col_frame_outer = tk.LabelFrame(self, text=" 마스킹할 컬럼 선택 ",
                                              font=("맑은 고딕", 10, "bold"),
                                              bg="#f4f6f9", padx=10, pady=8)
        self.col_frame_outer.pack(fill="both", expand=True, padx=18, pady=6)

        self.col_canvas = tk.Canvas(self.col_frame_outer, bg="#f4f6f9", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.col_frame_outer, orient="vertical",
                                  command=self.col_canvas.yview)
        self.col_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.col_canvas.pack(side="left", fill="both", expand=True)

        self.col_frame = tk.Frame(self.col_canvas, bg="#f4f6f9")
        self.col_canvas.create_window((0, 0), window=self.col_frame, anchor="nw")
        self.col_frame.bind("<Configure>",
            lambda e: self.col_canvas.configure(scrollregion=self.col_canvas.bbox("all")))

        self.placeholder = tk.Label(self.col_frame,
                                    text="파일을 열면 컬럼 목록이 표시됩니다.",
                                    font=("맑은 고딕", 10), fg="#94a3b8", bg="#f4f6f9")
        self.placeholder.pack(pady=30)

        # 실행 버튼 + 상태
        btm = tk.Frame(self, bg="#f4f6f9")
        btm.pack(fill="x", padx=18, pady=(4, 14))

        self.status_var = tk.StringVar(value="파일을 선택해 주세요.")
        tk.Label(btm, textvariable=self.status_var, font=("맑은 고딕", 9),
                 fg="#64748b", bg="#f4f6f9").pack(side="left")
        tk.Button(btm, text="  마스킹 실행 및 저장  ", command=self._run_masking,
                  bg="#16a34a", fg="white", font=("맑은 고딕", 11, "bold"),
                  relief="flat", padx=14, pady=6).pack(side="right")

    # ── 파일 열기 ─────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")])
        if not path:
            return
        try:
            self.df = pd.read_excel(path)
            self.filepath.set(path)
            self.status_var.set(f"✅ {len(self.df)}행 × {len(self.df.columns)}열 로드 완료")
            self._build_column_ui()
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽을 수 없습니다.\n{e}")

    # ── 컬럼 UI 빌드 ──────────────────────────────────────────

    def _build_column_ui(self):
        for w in self.col_frame.winfo_children():
            w.destroy()
        self.col_vars.clear()
        self.col_checks.clear()

        types = ["없음", "이름", "전화번호", "이메일", "주민등록번호", "주소"]

        # 헤더 행
        tk.Label(self.col_frame, text="선택", width=4, font=("맑은 고딕", 9, "bold"),
                 bg="#e2e8f0").grid(row=0, column=0, padx=4, pady=2, sticky="w")
        tk.Label(self.col_frame, text="컬럼명", width=22, font=("맑은 고딕", 9, "bold"),
                 bg="#e2e8f0", anchor="w").grid(row=0, column=1, padx=4, pady=2, sticky="w")
        tk.Label(self.col_frame, text="미리보기 (첫 3개)", width=30, font=("맑은 고딕", 9, "bold"),
                 bg="#e2e8f0", anchor="w").grid(row=0, column=2, padx=4, pady=2, sticky="w")
        tk.Label(self.col_frame, text="마스킹 유형", width=16, font=("맑은 고딕", 9, "bold"),
                 bg="#e2e8f0").grid(row=0, column=3, padx=4, pady=2, sticky="w")

        for i, col in enumerate(self.df.columns):
            detected = detect_column_type(self.df[col])
            check_var = tk.BooleanVar(value=(detected != "없음"))
            type_var  = tk.StringVar(value=detected)

            self.col_checks[col] = check_var
            self.col_vars[col]   = type_var

            bg = "#f0fdf4" if detected != "없음" else "#ffffff"
            row = i + 1

            tk.Checkbutton(self.col_frame, variable=check_var, bg=bg).grid(
                row=row, column=0, padx=4, sticky="w")
            tk.Label(self.col_frame, text=col, width=22, anchor="w",
                     font=("맑은 고딕", 9), bg=bg).grid(
                row=row, column=1, padx=4, pady=1, sticky="w")

            preview = ", ".join(str(v) for v in self.df[col].dropna().head(3).tolist())
            preview = preview[:35] + "…" if len(preview) > 35 else preview
            tk.Label(self.col_frame, text=preview, width=30, anchor="w",
                     font=("맑은 고딕", 9), fg="#475569", bg=bg).grid(
                row=row, column=2, padx=4, pady=1, sticky="w")

            ttk.Combobox(self.col_frame, textvariable=type_var, values=types,
                         width=14, state="readonly").grid(
                row=row, column=3, padx=4, pady=1)

    # ── 마스킹 실행 ───────────────────────────────────────────

    def _run_masking(self):
        if self.df is None:
            messagebox.showwarning("알림", "먼저 파일을 선택해 주세요.")
            return

        selected = [(col, self.col_vars[col].get())
                    for col, chk in self.col_checks.items()
                    if chk.get() and self.col_vars[col].get() != "없음"]

        if not selected:
            messagebox.showwarning("알림", "마스킹할 컬럼을 하나 이상 선택해 주세요.")
            return

        # 저장 경로
        src = self.filepath.get()
        base, ext = os.path.splitext(src)
        default_save = base + "_마스킹" + ext
        save_path = filedialog.asksaveasfilename(
            title="저장 위치 선택",
            initialfile=os.path.basename(default_save),
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")])
        if not save_path:
            return

        self.status_var.set("⏳ 마스킹 처리 중…")
        self.update()

        def worker():
            try:
                result = self.df.copy()
                for col, mtype in selected:
                    fn = MASK_FUNCS.get(mtype)
                    if fn:
                        result[col] = result[col].apply(
                            lambda v: fn(v) if pd.notna(v) else v)

                result.to_excel(save_path, index=False)
                self.after(0, lambda: (
                    self.status_var.set(f"✅ 저장 완료 → {os.path.basename(save_path)}"),
                    messagebox.showinfo("완료",
                        f"마스킹이 완료되었습니다!\n\n저장 위치:\n{save_path}")
                ))
            except Exception as e:
                self.after(0, lambda: (
                    self.status_var.set("❌ 오류 발생"),
                    messagebox.showerror("오류", f"처리 중 오류가 발생했습니다.\n{e}")
                ))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = MaskingApp()
    app.mainloop()
