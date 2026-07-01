import tkinter as tk
from tkinter import ttk
from app.core.word_repository import search_words, fetch_words_by_ids
from app.ui.study_window import StudyWindow

class SearchPanel(ttk.LabelFrame):
    def __init__(self, parent):
        super().__init__(parent, text="Search")
        self.parent = parent
        self.rows = []

        self.query_var = tk.StringVar()
        self.query_var.trace_add("write", lambda *_: self.refresh())

        entry = ttk.Entry(self, textvariable=self.query_var)
        entry.pack(fill="x", padx=8, pady=8)

        self.listbox = tk.Listbox(self, height=10, selectmode="extended")
        self.listbox.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.listbox.bind("<Double-Button-1>", lambda e: self.study_selected())

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="선택 학습", command=self.study_selected).pack(side="left")
        ttk.Button(btns, text="검색 결과 전체 학습", command=self.study_all).pack(side="left", padx=6)

    def refresh(self):
        q = self.query_var.get().strip()
        self.listbox.delete(0, "end")
        self.rows = []
        if not q:
            return
        self.rows = search_words(q)
        for row in self.rows:
            _, word, meaning, example, example_ko, memo, level, tags = row
            self.listbox.insert("end", f"{word} | {meaning} | {level} | {tags}")

    def study_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        ids = [self.rows[i][0] for i in sel]
        rows = fetch_words_by_ids(ids)
        if rows:
            StudyWindow(self.parent, rows, mode="search", title="Search Study")

    def study_all(self):
        if self.rows:
            StudyWindow(self.parent, self.rows, mode="search", title="Search Study")
