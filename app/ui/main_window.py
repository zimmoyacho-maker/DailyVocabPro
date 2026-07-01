import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date

from app.database.db import initialize_database, get_setting
from app.database.excel_importer import import_words_from_excel
from app.core.word_repository import count_words, fetch_new_words, fetch_review_words
from app.core.srs import due_review_count
from app.core.statistics import today_stats
from app.ui.study_window import StudyWindow
from app.ui.search_panel import SearchPanel

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DailyVocab Pro")
        self.geometry("900x700")
        self.minsize(760, 600)

        initialize_database()
        self._build_ui()
        self.refresh_dashboard()

    def _build_ui(self):
        header = ttk.Frame(self, padding=16)
        header.pack(fill="x")

        ttk.Label(header, text="DailyVocab Pro", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(header, text="Learn English Smarter.", font=("Segoe UI", 11)).pack(anchor="w")

        self.status_var = tk.StringVar()
        ttk.Label(header, textvariable=self.status_var).pack(anchor="w", pady=(8, 0))

        actions = ttk.Frame(self, padding=(16, 4))
        actions.pack(fill="x")

        ttk.Button(actions, text="Excel 불러오기", command=self.import_excel).pack(side="left")
        ttk.Button(actions, text="오늘 새 단어", command=self.start_new).pack(side="left", padx=6)
        ttk.Button(actions, text="오늘 복습", command=self.start_review).pack(side="left", padx=6)

        self.search_panel = SearchPanel(self)
        self.search_panel.pack(fill="both", expand=True, padx=16, pady=16)

    def refresh_dashboard(self):
        stats = today_stats()
        total_words = count_words()
        due = due_review_count(date.today().isoformat())
        self.status_var.set(
            f"총 단어: {total_words} | 오늘 복습 예정: {due} | "
            f"오늘 학습: {stats['total']}개 | 정답률: {stats['accuracy']}% | {stats['stars']}"
        )

    def import_excel(self):
        try:
            n = import_words_from_excel()
            messagebox.showinfo("완료", f"{n}개 단어를 불러왔습니다.")
            self.refresh_dashboard()
        except Exception as e:
            messagebox.showerror("Excel 오류", str(e))

    def start_new(self):
        limit = int(get_setting("daily_new", "10"))
        rows = fetch_new_words(limit)
        if not rows:
            messagebox.showinfo("안내", "새 단어가 없습니다. Excel을 불러오거나 모든 단어가 이미 학습되었습니다.")
            return
        StudyWindow(self, rows, mode="new", title="New Words")

    def start_review(self):
        limit = int(get_setting("daily_review", "5"))
        rows = fetch_review_words(limit, date.today().isoformat())
        if not rows:
            messagebox.showinfo("안내", "오늘 복습할 단어가 없습니다.")
            return
        StudyWindow(self, rows, mode="review", title="Review")
