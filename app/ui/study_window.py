import tkinter as tk
from tkinter import ttk, messagebox
from app.core.srs import apply_rating
from app.ui.copy_text import make_copyable_text
from app.services.speech import SpeechService
from app.services.ai_tutor import build_ai_tutor_text

class StudyWindow(tk.Toplevel):
    def __init__(self, parent, words, mode="study", title="Study"):
        super().__init__(parent)
        self.parent = parent
        self.words = list(words)
        self.mode = mode
        self.index = 0
        self.missed = []
        self.reloop_count = 0
        self.speech = SpeechService()

        self.title(title)
        self.geometry("760x620")
        self.minsize(640, 500)

        self.word_var = tk.StringVar()
        self.progress_var = tk.StringVar()

        self._build_ui()
        self._bind_keys()
        self._load_current()

    def _build_ui(self):
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Label(top, textvariable=self.progress_var).pack(anchor="w")
        ttk.Label(top, textvariable=self.word_var, font=("Segoe UI", 26, "bold")).pack(anchor="center", pady=(8, 8))

        controls = ttk.Frame(top)
        controls.pack(fill="x")

        ttk.Button(controls, text="🔊 단어 (W)", command=self.speak_word).pack(side="left", padx=4)
        ttk.Button(controls, text="🔊 예문 (E)", command=self.speak_example).pack(side="left", padx=4)
        ttk.Button(controls, text="🤖 AI Tutor (A)", command=self.show_ai_tutor).pack(side="right", padx=4)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.body = ttk.Frame(self.canvas, padding=12)

        self.body.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def on_wheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        bottom = ttk.Frame(self, padding=12)
        bottom.pack(fill="x")

        ttk.Button(bottom, text="😊 알아요 (1)", command=lambda: self.rate("know")).pack(side="left", expand=True, fill="x", padx=4)
        ttk.Button(bottom, text="🤔 애매해요 (2)", command=lambda: self.rate("maybe")).pack(side="left", expand=True, fill="x", padx=4)
        ttk.Button(bottom, text="😭 몰라요 (3)", command=lambda: self.rate("dont")).pack(side="left", expand=True, fill="x", padx=4)

    def _bind_keys(self):
        self.bind("1", lambda e: self.rate("know"))
        self.bind("2", lambda e: self.rate("maybe"))
        self.bind("3", lambda e: self.rate("dont"))
        self.bind("w", lambda e: self.speak_word())
        self.bind("W", lambda e: self.speak_word())
        self.bind("e", lambda e: self.speak_example())
        self.bind("E", lambda e: self.speak_example())
        self.bind("a", lambda e: self.show_ai_tutor())
        self.bind("A", lambda e: self.show_ai_tutor())
        self.bind("<Escape>", lambda e: self.destroy())

    def _clear_body(self):
        for child in self.body.winfo_children():
            child.destroy()

    def _current(self):
        return self.words[self.index]

    def _load_current(self):
        if self.index >= len(self.words):
            self._finish()
            return

        row = self._current()
        word_id, word, meaning, example, example_ko, memo, level, tags = row

        self.word_var.set(word)
        self.progress_var.set(f"{self.index + 1} / {len(self.words)}")
        self._clear_body()

        self._section("Example", example or "(예문 없음)", height=4)
        self._reveal_button("번역 보기 (T)", "Translation", example_ko or "(번역 없음)", 3, "t")
        self._reveal_button("뜻 보기 (Space)", "Meaning", meaning or "(뜻 없음)", 3, "space")

        if memo:
            self._section("Memo", memo, height=3)
        if level or tags:
            self._section("Level / Tags", f"{level or ''}  {tags or ''}".strip(), height=2)

    def _section(self, title, text, height=4):
        ttk.Label(self.body, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 4))
        box = make_copyable_text(self.body, text, height=height)
        box.pack(fill="x", pady=(0, 6))

    def _reveal_button(self, button_text, title, text, height, key):
        frame = ttk.Frame(self.body)
        frame.pack(fill="x", pady=(8, 0))

        revealed = {"value": False}

        def reveal(event=None):
            if revealed["value"]:
                return
            revealed["value"] = True
            btn.pack_forget()
            self._section(title, text, height=height)

        btn = ttk.Button(frame, text=button_text, command=reveal)
        btn.pack(anchor="w")

        if key == "t":
            self.bind("t", reveal)
            self.bind("T", reveal)
        elif key == "space":
            self.bind("<space>", reveal)

    def rate(self, result):
        if self.index >= len(self.words):
            return

        row = self._current()
        word_id = row[0]
        apply_rating(word_id, result, self.mode)

        if result in ("maybe", "dont"):
            self.missed.append(row)

        self.index += 1
        self._load_current()

    def _finish(self):
        if self.missed and self.reloop_count < 3:
            again = self.missed
            self.words = again
            self.missed = []
            self.index = 0
            self.reloop_count += 1
            messagebox.showinfo("재복습", f"애매하거나 모르는 단어 {len(again)}개를 다시 복습합니다.")
            self._load_current()
            return

        messagebox.showinfo("완료", "학습이 완료되었습니다.")
        try:
            self.parent.refresh_dashboard()
        except Exception:
            pass
        self.destroy()

    def speak_word(self):
        try:
            self.speech.speak(self._current()[1])
        except Exception as e:
            messagebox.showwarning("발음 기능", str(e))

    def speak_example(self):
        try:
            self.speech.speak(self._current()[3])
        except Exception as e:
            messagebox.showwarning("발음 기능", str(e))

    def show_ai_tutor(self):
        win = tk.Toplevel(self)
        win.title("AI Tutor")
        win.geometry("620x520")
        text = make_copyable_text(win, build_ai_tutor_text(self._current()), height=25)
        text.pack(fill="both", expand=True, padx=12, pady=12)
