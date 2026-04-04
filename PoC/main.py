"""
main.py
HWP-LLM Wrapper - 메인 진입점 (Tkinter GUI)

사용법:
    python main.py

의존성:
    pip install -r requirements.txt
"""
import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from hwp_controller import HwpController
from llm_client import DEFAULT_MODELS, LlmClient

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".hwp_llm_config.json")

COLOR_BG = "#1e1e2e"
COLOR_SURFACE = "#2a2a3e"
COLOR_ACCENT = "#7c6af7"
COLOR_ACCENT_HOVER = "#9d8fff"
COLOR_TEXT = "#cdd6f4"
COLOR_TEXT_DIM = "#6c7086"
COLOR_SUCCESS = "#a6e3a1"
COLOR_ERROR = "#f38ba8"
COLOR_WARNING = "#fab387"
COLOR_BORDER = "#363653"

FONT_DEFAULT = ("Pretendard", 10) if sys.platform == "win32" else ("Segoe UI", 10)
FONT_MONO = ("D2Coding", 10) if sys.platform == "win32" else ("Consolas", 10)
FONT_TITLE = ("Pretendard", 16, "bold") if sys.platform == "win32" else ("Segoe UI", 16, "bold")
FONT_LABEL = ("Pretendard", 9) if sys.platform == "win32" else ("Segoe UI", 9)


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class HwpLlmApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP-LLM Wrapper")
        self.geometry("860x720")
        self.minsize(700, 580)
        self.configure(bg=COLOR_BG)
        self.resizable(True, True)

        self.cfg = load_config()
        self.hwp = HwpController()
        self.hwp_connected = False
        self.current_file = tk.StringVar(value="")
        self._doc_summary: str | None = None

        # 에이전틱 루프 동기화 큐
        self._tool_request_q: queue.Queue = queue.Queue()
        self._tool_result_q: queue.Queue = queue.Queue()
        self._agentic_running = False

        self._build_ui()
        self._apply_styles()
        self._load_saved_config()
        self._set_icon()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ──────────────────── UI 빌드 ────────────────────────────

    def _build_ui(self):
        header = tk.Frame(self, bg=COLOR_BG, pady=16)
        header.pack(fill=tk.X, padx=24)
        tk.Label(header, text="✦ HWP-LLM Wrapper",
                 font=FONT_TITLE, fg=COLOR_ACCENT, bg=COLOR_BG).pack(side=tk.LEFT)
        tk.Label(header, text="한글 문서를 AI로 편집",
                 font=FONT_LABEL, fg=COLOR_TEXT_DIM, bg=COLOR_BG).pack(
            side=tk.LEFT, padx=(12, 0), pady=(6, 0))

        tk.Frame(self, height=1, bg=COLOR_BORDER).pack(fill=tk.X, padx=24)

        content = tk.Frame(self, bg=COLOR_BG)
        content.pack(fill=tk.BOTH, expand=True, padx=24, pady=16)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        self._build_left_panel(content)
        self._build_right_panel(content)
        self._build_statusbar()

    def _build_left_panel(self, parent):
        left = tk.Frame(parent, bg=COLOR_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._section_label(left, "📄 HWP 파일")
        file_row = tk.Frame(left, bg=COLOR_BG)
        file_row.pack(fill=tk.X, pady=(4, 12))
        self.file_entry = self._entry(file_row, textvariable=self.current_file)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._button(file_row, "열기", self._browse_file, small=True).pack(
            side=tk.RIGHT, padx=(6, 0))

        file_btn_row = tk.Frame(left, bg=COLOR_BG)
        file_btn_row.pack(fill=tk.X, pady=(0, 12))
        self._button(file_btn_row, "한글에서 열기", self._open_in_hwp,
                     accent=True, small=True).pack(side=tk.LEFT)
        self._button(file_btn_row, "문서 구조 미리보기", self._preview_structure,
                     small=True).pack(side=tk.LEFT, padx=(6, 0))

        tk.Frame(left, height=1, bg=COLOR_BORDER).pack(fill=tk.X, pady=8)

        self._section_label(left, "🤖 LLM 설정")

        prov_row = tk.Frame(left, bg=COLOR_BG)
        prov_row.pack(fill=tk.X, pady=(4, 6))
        tk.Label(prov_row, text="공급자", width=10, anchor="w",
                 font=FONT_LABEL, fg=COLOR_TEXT_DIM, bg=COLOR_BG).pack(side=tk.LEFT)
        self.provider_var = tk.StringVar(value="openai")
        self.provider_cb = ttk.Combobox(
            prov_row, textvariable=self.provider_var,
            values=["openai", "gemini", "anthropic"],
            state="readonly", width=16)
        self.provider_cb.pack(side=tk.LEFT)
        self.provider_cb.bind("<<ComboboxSelected>>", self._on_provider_change)

        model_row = tk.Frame(left, bg=COLOR_BG)
        model_row.pack(fill=tk.X, pady=(0, 6))
        tk.Label(model_row, text="모델", width=10, anchor="w",
                 font=FONT_LABEL, fg=COLOR_TEXT_DIM, bg=COLOR_BG).pack(side=tk.LEFT)
        self.model_var = tk.StringVar(value=DEFAULT_MODELS["openai"])
        self._entry(model_row, textvariable=self.model_var, width=28).pack(side=tk.LEFT)

        key_row = tk.Frame(left, bg=COLOR_BG)
        key_row.pack(fill=tk.X, pady=(0, 12))
        tk.Label(key_row, text="API Key", width=10, anchor="w",
                 font=FONT_LABEL, fg=COLOR_TEXT_DIM, bg=COLOR_BG).pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar()
        self._entry(key_row, textvariable=self.api_key_var, show="•", width=28).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        self._small_link(key_row, "저장", self._save_key).pack(side=tk.RIGHT, padx=(4, 0))

        tk.Frame(left, height=1, bg=COLOR_BORDER).pack(fill=tk.X, pady=8)

        self._section_label(left, "💬 수정 요청")
        self.query_text = scrolledtext.ScrolledText(
            left, height=7, wrap=tk.WORD,
            bg=COLOR_SURFACE, fg=COLOR_TEXT, font=FONT_DEFAULT,
            insertbackground=COLOR_TEXT, relief=tk.FLAT, padx=10, pady=8,
            highlightthickness=1, highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_ACCENT)
        self.query_text.pack(fill=tk.BOTH, expand=True, pady=(4, 12))
        self.query_text.insert("1.0", "예: '제목을 「2026년 업무 보고서」로 변경해줘'")
        self.query_text.bind("<FocusIn>", self._clear_placeholder)

        self.run_btn = self._button(left, "⚡ AI로 문서 수정 실행", self._run, accent=True)
        self.run_btn.pack(fill=tk.X, ipady=6)

    def _build_right_panel(self, parent):
        right = tk.Frame(parent, bg=COLOR_BG)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self._section_label(right, "📋 처리 로그")
        self.log_text = scrolledtext.ScrolledText(
            right, wrap=tk.WORD, bg=COLOR_SURFACE, fg=COLOR_TEXT,
            font=FONT_MONO, insertbackground=COLOR_TEXT, relief=tk.FLAT,
            padx=10, pady=8, highlightthickness=1,
            highlightbackground=COLOR_BORDER, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        self.log_text.tag_configure("success", foreground=COLOR_SUCCESS)
        self.log_text.tag_configure("error", foreground=COLOR_ERROR)
        self.log_text.tag_configure("warning", foreground=COLOR_WARNING)
        self.log_text.tag_configure("info", foreground=COLOR_ACCENT)
        self.log_text.tag_configure("dim", foreground=COLOR_TEXT_DIM)

        self._button(right, "로그 지우기", self._clear_log, small=True).pack(anchor="e")

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=COLOR_SURFACE, pady=4)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="준비")
        tk.Label(bar, textvariable=self.status_var, font=FONT_LABEL,
                 fg=COLOR_TEXT_DIM, bg=COLOR_SURFACE, anchor="w", padx=16).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=120)
        self.progress.pack(side=tk.RIGHT, padx=12)

    # ──────────────────── 위젯 헬퍼 ─────────────────────────

    def _section_label(self, parent, text):
        tk.Label(parent, text=text, font=(*FONT_DEFAULT[:1], 10, "bold"),
                 fg=COLOR_TEXT, bg=COLOR_BG, anchor="w").pack(fill=tk.X)

    def _entry(self, parent, **kw):
        return tk.Entry(parent, bg=COLOR_SURFACE, fg=COLOR_TEXT, font=FONT_DEFAULT,
                        insertbackground=COLOR_TEXT, relief=tk.FLAT,
                        highlightthickness=1, highlightbackground=COLOR_BORDER,
                        highlightcolor=COLOR_ACCENT, **kw)

    def _button(self, parent, text, cmd, accent=False, small=False):
        bg = COLOR_ACCENT if accent else COLOR_SURFACE
        fg = "#ffffff" if accent else COLOR_TEXT
        font = (*FONT_DEFAULT[:1], 9) if small else (*FONT_DEFAULT[:1], 10, "bold")
        btn = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                        activebackground=COLOR_ACCENT_HOVER, activeforeground="#ffffff",
                        font=font, relief=tk.FLAT, cursor="hand2",
                        padx=12, pady=4 if small else 6, bd=0)
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLOR_ACCENT_HOVER if accent else COLOR_BORDER))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    def _small_link(self, parent, text, cmd):
        lbl = tk.Label(parent, text=text, font=FONT_LABEL, fg=COLOR_ACCENT,
                       bg=COLOR_BG, cursor="hand2", underline=True)
        lbl.bind("<Button-1>", lambda e: cmd())
        return lbl

    def _apply_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox", fieldbackground=COLOR_SURFACE,
                        background=COLOR_SURFACE, foreground=COLOR_TEXT,
                        selectbackground=COLOR_ACCENT, selectforeground="#fff")
        style.configure("TProgressbar", troughcolor=COLOR_SURFACE, background=COLOR_ACCENT)

    def _set_icon(self):
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

    # ──────────────────── 이벤트 핸들러 ─────────────────────

    def _clear_placeholder(self, event):
        if self.query_text.get("1.0", "end-1c") == "예: '제목을 「2026년 업무 보고서」로 변경해줘'":
            self.query_text.delete("1.0", tk.END)

    def _on_provider_change(self, event=None):
        provider = self.provider_var.get()
        self.model_var.set(DEFAULT_MODELS.get(provider, ""))
        saved_key = self.cfg.get("api_keys", {}).get(provider, "")
        self.api_key_var.set(saved_key)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="HWP 파일 선택",
            filetypes=[("한글 문서", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
        if path:
            self.current_file.set(path)
            self._doc_summary = None
            self._log(f"파일 선택: {path}", "dim")

    def _save_key(self):
        provider = self.provider_var.get()
        key = self.api_key_var.get().strip()
        if not self.cfg.get("api_keys"):
            self.cfg["api_keys"] = {}
        self.cfg["api_keys"][provider] = key
        self.cfg["last_provider"] = provider
        save_config(self.cfg)
        self._log(f"API Key 저장 완료 ({provider})", "success")

    def _load_saved_config(self):
        provider = self.cfg.get("last_provider", "openai")
        self.provider_var.set(provider)
        self.model_var.set(DEFAULT_MODELS.get(provider, ""))
        saved_key = self.cfg.get("api_keys", {}).get(provider, "")
        self.api_key_var.set(saved_key)

    def _open_in_hwp(self):
        path = self.current_file.get().strip()
        if not path:
            messagebox.showwarning("파일 없음", "HWP 파일을 먼저 선택하세요.")
            return
        self._set_status("한글 연결 중...")
        self.after(0, self._connect_and_open, path)

    def _connect_and_open(self, path):
        try:
            if not self.hwp_connected:
                self.hwp.connect(visible=True)
                self.hwp_connected = True
                self._log("한글 연결 성공 ✓", "success")
            self.hwp.open_file(path)
            self._log(f"파일 열기 성공: {os.path.basename(path)}", "success")
            self._set_status("파일 열림")
            self._doc_summary = None
        except Exception as exc:
            self._log(f"오류: {exc}", "error")
            self._set_status("연결 실패")
            import traceback
            self._log(traceback.format_exc(), "dim")

    def _preview_structure(self):
        if not self.hwp_connected:
            messagebox.showinfo("안내", "먼저 '한글에서 열기'를 실행하세요.")
            return
        self._set_status("문서 구조 추출 중...")
        self.after(0, self._do_preview)

    def _do_preview(self):
        try:
            structure = self.hwp.extract_document_structure()
            self._log("─── 추출 진단 ───", "info")
            for entry in structure.get("extraction_log", []):
                self._log(f"  {entry}", "dim")

            # 표 구조 미리보기
            try:
                doc_info = self.hwp.analyze_document_structure()
                self._log(f"페이지: {doc_info['pages']}  표: {doc_info['table_count']}개  필드: {len(doc_info['fields'])}개", "info")
                if doc_info["fields"]:
                    self._log(f"필드 목록: {', '.join(doc_info['fields'])}", "dim")
                for ti in range(min(doc_info["table_count"], 3)):
                    schema = self.hwp.get_table_schema(ti)
                    self._log(f"  표[{ti}]: {schema['rows']}행 × {schema['cols']}열", "dim")
                    if schema.get("headers"):
                        self._log(f"    헤더: {schema['headers']}", "dim")
            except Exception:
                pass

            summary = self.hwp.get_summary_text(_structure=structure)
            self._doc_summary = summary
            self._log(f"─── 문서 요약 ({len(summary)}자) ───", "info")
            self._log(summary[:3000], "dim")
            self._log("─────────────────────────", "info")
            self._set_status("추출 완료")
        except Exception as exc:
            self._log(f"추출 실패: {exc}", "error")
            self._set_status("오류")

    # ──────────────────── 에이전틱 실행 흐름 ────────────────

    def _run(self):
        file_path = self.current_file.get().strip()
        api_key = self.api_key_var.get().strip()
        query = self.query_text.get("1.0", tk.END).strip()

        if not file_path:
            messagebox.showwarning("입력 오류", "HWP 파일을 선택하세요.")
            return
        if not api_key:
            messagebox.showwarning("입력 오류", "API Key를 입력하세요.")
            return
        if not query or query.startswith("예:"):
            messagebox.showwarning("입력 오류", "수정 요청 내용을 입력하세요.")
            return
        if not self.hwp_connected:
            messagebox.showinfo("안내", "먼저 '한글에서 열기'를 실행해 한글에 연결하세요.")
            return

        self.run_btn.configure(state=tk.DISABLED)
        self.progress.start(10)
        # Step 1: 문서 구조 추출 (메인 스레드 — HWP COM)
        self.after(0, self._run_step1_extract, file_path, api_key, query)

    def _run_step1_extract(self, file_path: str, api_key: str, query: str):
        """Step 1: 문서 구조 추출 — 메인 스레드"""
        try:
            provider = self.provider_var.get()
            model = self.model_var.get().strip()
            self._log(f"\n⚡ 시작 — {provider.upper()} / {model}", "info")
            self._set_status("문서 분석 중...")

            # 초기 문서 요약 (캐시 사용)
            if self._doc_summary is None:
                self._log("문서 구조 추출 중...", "dim")
                structure = self.hwp.extract_document_structure()
                for entry in structure.get("extraction_log", []):
                    self._log(f"  {entry}", "dim")
                self._doc_summary = self.hwp.get_summary_text(_structure=structure)

            self._log(f"문서 요약: {len(self._doc_summary)}자", "dim")
            if len(self._doc_summary) < 100:
                self._log("⚠️ 문서 내용이 거의 없습니다. 파일이 제대로 열렸는지 확인하세요.", "warning")

            # Step 2: 에이전틱 루프 시작
            self._set_status("AI 에이전트 실행 중...")
            self._log("─── 에이전틱 Tool Use 루프 시작 ───", "info")
            self._agentic_running = True
            self.after(50, self._poll_tool_requests)  # 도구 요청 폴링 시작

            doc_summary = self._doc_summary
            threading.Thread(
                target=self._agentic_thread,
                args=(api_key, query, provider, model, doc_summary),
                daemon=True,
            ).start()

        except Exception as exc:
            self._log(f"\n❌ 추출 오류: {exc}", "error")
            import traceback
            self._log(traceback.format_exc(), "dim")
            self.after(0, self._reset_ui)

    def _poll_tool_requests(self):
        """
        메인 스레드에서 도구 요청 큐를 폴링합니다.
        에이전틱 루프(백그라운드)가 도구 호출을 요청하면 메인 스레드에서 실행 후 결과 반환.
        """
        try:
            tool_name, tool_args = self._tool_request_q.get_nowait()
            try:
                result = self.hwp.dispatch_tool(tool_name, tool_args)
            except Exception as e:
                result = f"ERROR: {e}"
            self._tool_result_q.put(result)
        except queue.Empty:
            pass

        if self._agentic_running:
            self.after(30, self._poll_tool_requests)

    def _agentic_thread(self, api_key: str, query: str,
                         provider: str, model: str, doc_summary: str):
        """
        백그라운드 스레드: LLM 에이전틱 루프 실행.
        도구 호출은 큐를 통해 메인 스레드에 위임합니다.
        """
        def tool_executor(tool_name: str, tool_args: dict) -> str:
            self._tool_request_q.put((tool_name, tool_args))
            return self._tool_result_q.get()  # 메인 스레드 처리 대기

        try:
            client = LlmClient(provider=provider, api_key=api_key, model=model)
            final_response = client.call_agentic(
                doc_summary=doc_summary,
                user_query=query,
                tool_executor=tool_executor,
                log_cb=self._log,
                max_turns=25,
            )
            self.after(0, self._on_agentic_done, final_response)
        except Exception as exc:
            self._log(f"\n❌ 에이전트 오류: {exc}", "error")
            import traceback
            self._log(traceback.format_exc(), "dim")
            self.after(0, self._reset_ui_after_agentic)

    def _on_agentic_done(self, final_response: str):
        """에이전틱 루프 완료 — 메인 스레드에서 저장 처리"""
        self._agentic_running = False
        self._log(f"\n─── AI 응답 ───", "info")
        self._log(final_response, "success")
        self._set_status("완료")
        self._reset_ui_after_agentic()
        self._doc_summary = None

    def _reset_ui_after_agentic(self):
        self._agentic_running = False
        self.run_btn.configure(state=tk.NORMAL)
        self.progress.stop()

    # ──────────────────── 레거시 호환 (사용 안 함) ──────────

    def _reset_ui(self):
        self.run_btn.configure(state=tk.NORMAL)
        self.progress.stop()

    def _on_close(self):
        self._agentic_running = False
        try:
            if self.hwp_connected:
                self.hwp.close(save=False)
        except Exception:
            pass
        self.destroy()

    # ──────────────────── 로그 헬퍼 ─────────────────────────

    def _log(self, msg: str, tag: str = ""):
        def _do():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        self.after(0, _do)

    def _clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_var.set(msg))


if __name__ == "__main__":
    app = HwpLlmApp()
    app.mainloop()
