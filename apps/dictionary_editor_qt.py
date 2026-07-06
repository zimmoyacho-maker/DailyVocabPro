#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DailyVocabPro Master Dictionary Editor v2.2 Productivity

New in v2.2
-----------
- Offline Assist panel
- Auto Memo draft
- Auto Tags recommendation
- Auto Level recommendation
- Apply All Assist
- Batch dashboard
- QA color status
- Better shortcuts

Run:
    pip install -r apps/requirements_dictionary_editor.txt
    python apps/dictionary_editor_qt.py
"""

from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSplitter,
    QStatusBar, QVBoxLayout, QWidget,
)

REQUIRED_COLUMNS = ["word", "meaning", "example"]
EDIT_COLUMNS = ["example_ko", "memo", "level", "tags"]
OPTIONAL_COLUMNS = ["qa_flag"]
ALL_COLUMNS = REQUIRED_COLUMNS + EDIT_COLUMNS + OPTIONAL_COLUMNS

LEVEL_OPTIONS = ["", "A1", "A2", "B1", "B2", "C1", "C2"]
TAG_OPTIONS = [
    "Daily", "Business", "Academic", "Education", "Travel",
    "Medical", "Legal", "Science", "Technology", "Conversation",
]

A1_WORDS = {
    "be", "do", "go", "come", "get", "make", "take", "see", "look", "want", "need",
    "like", "good", "bad", "big", "small", "new", "old", "hot", "cold", "day", "time",
    "year", "home", "house", "school", "book", "water", "food", "man", "woman", "child",
}
A2_WORDS = {
    "traffic", "wall", "calendar", "temperature", "damage", "floor", "roof", "rain",
    "purpose", "change", "escape", "ready", "poor", "arrest", "score", "plan", "record",
    "movie", "ocean", "inside", "order", "guest", "guide", "control", "save", "plant",
    "green", "flavor", "wave", "break", "crop", "fall", "hall", "drain", "dog", "school",
    "death", "cycle", "honor", "corner", "core", "conflict", "exceed", "policy", "prison",
    "bitter", "weapon", "cinema", "cap", "weak", "unit", "huge", "sea", "captain", "sand",
}

TAG_KEYWORDS = {
    "Business": ["company", "market", "client", "contract", "revenue", "business", "office", "policy", "committee", "price", "proposal", "sales"],
    "Academic": ["theory", "philosophy", "research", "mathematical", "equation", "vector", "dimension", "definition", "university", "student", "teacher"],
    "Education": ["school", "student", "teacher", "pupil", "college", "university", "class", "lesson", "learn"],
    "Science": ["plant", "temperature", "biology", "geology", "disease", "velocity", "microscope", "fusion", "species", "chemical"],
    "Medical": ["doctor", "medicine", "medication", "disease", "health", "hospital", "drug", "cancer", "patient", "pain"],
    "Legal": ["court", "law", "police", "crime", "murder", "arrest", "custody", "defendant", "judge", "legal"],
    "Travel": ["airport", "hotel", "restaurant", "ship", "train", "bus", "station", "route", "embassy", "tour", "museum", "harbor"],
    "Technology": ["software", "computer", "server", "protocol", "network", "program", "http", "windows", "device", "data", "file"],
    "Daily": ["house", "room", "family", "food", "coffee", "clothes", "baby", "dog", "cat", "friend", "home", "desk", "kitchen"],
}

PATTERN_MEMOS = [
    (r"\bbe obsessed with\b", "be obsessed with = ~에 푹 빠져 있다 / 집착하다. 회화와 시험에서 자주 쓰이는 표현입니다."),
    (r"\bbe based on\b", "be based on = ~에 근거하다. 주장이나 제안의 근거를 말할 때 자주 씁니다."),
    (r"\bbe famous for\b", "be famous for = ~로 유명하다. 이유나 특징을 나타낼 때 사용합니다."),
    (r"\bbe famous as\b", "be famous as = ~로서 유명하다. 역할이나 신분을 나타낼 때 사용합니다."),
    (r"\bconsist(?:s|ed|ing)? of\b", "consist of = ~로 이루어져 있다. 구성 요소를 설명할 때 자주 사용합니다."),
    (r"\bdepend(?:s|ed|ing)? (on|upon)\b", "depend on/upon = ~에 달려 있다 / 의존하다."),
    (r"\bbe charged with\b", "be charged with = ~혐의로 기소되다. 법률 문맥에서 자주 쓰입니다."),
    (r"\bprohibit(?:s|ed|ing)?\b", "prohibit = 금지하다. 공식 문서에서 자주 쓰이는 표현입니다."),
    (r"\bundergo(?:es|ing|ne|went)?\b", "undergo = 변화·과정·치료 등을 겪다. undergo changes처럼 자주 씁니다."),
    (r"\bresult(?:s|ed|ing)? in\b", "result in = ~라는 결과를 낳다."),
]

DARK_STYLE = """
QMainWindow, QWidget {
    background-color: #111827;
    color: #E5E7EB;
    font-family: Segoe UI;
    font-size: 10.5pt;
}
QFrame#Card {
    background-color: #1F2937;
    border: 1px solid #374151;
    border-radius: 10px;
}
QLabel#TitleWord {
    font-size: 30px;
    font-weight: 700;
    color: #F9FAFB;
}
QLabel#Meaning {
    font-size: 16px;
    color: #D1D5DB;
}
QLabel#SectionTitle {
    font-size: 12px;
    font-weight: 700;
    color: #93C5FD;
}
QPlainTextEdit, QLineEdit, QComboBox, QListWidget {
    background-color: #0B1220;
    color: #E5E7EB;
    border: 1px solid #374151;
    border-radius: 6px;
    padding: 6px;
}
QPushButton {
    background-color: #2563EB;
    color: white;
    border: 0;
    border-radius: 7px;
    padding: 8px 12px;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #1D4ED8;
}
QPushButton#Secondary {
    background-color: #374151;
}
QPushButton#Secondary:hover {
    background-color: #4B5563;
}
QPushButton#Assist {
    background-color: #7C3AED;
}
QPushButton#Assist:hover {
    background-color: #6D28D9;
}
QProgressBar {
    border: 1px solid #374151;
    border-radius: 6px;
    text-align: center;
    background-color: #0B1220;
}
QProgressBar::chunk {
    background-color: #2563EB;
    border-radius: 6px;
}
QStatusBar {
    background-color: #0B1220;
    color: #D1D5DB;
}
"""


def norm(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").replace("\u2003", " ").strip()


def recommend_level(word: str, example: str) -> str:
    w = word.lower().strip()
    e = example.lower()
    if w in A1_WORDS:
        return "A1"
    if w in A2_WORDS:
        return "A2"
    if any(k in e for k in ["equation", "vector", "theory", "philosophy", "protocol", "government", "political"]):
        return "B2"
    if len(w) >= 12 or len(example) >= 150:
        return "C1"
    if len(w) >= 9 or len(example) >= 90:
        return "B2"
    return "B1"


def recommend_tags(word: str, meaning: str, example: str) -> str:
    text = f"{word} {meaning} {example}".lower()
    tags = []
    for tag, keywords in TAG_KEYWORDS.items():
        if any(k in text for k in keywords):
            tags.append(tag)
    if not tags:
        tags.append("Daily")
    return ", ".join(tags[:2])


def recommend_memo(word: str, meaning: str, example: str) -> str:
    lower = example.lower()
    notes = []
    for pattern, memo in PATTERN_MEMOS:
        if re.search(pattern, lower):
            notes.append(memo)

    if not notes:
        notes.append(f"{word} = {meaning}. 예문 속 문맥과 함께 뜻을 기억하세요.")

    if " with " in lower and word.lower() in lower:
        notes.append("전치사 with와 함께 쓰이는지 확인해 보세요.")
    if " of " in lower and word.lower() in lower:
        notes.append("of와 함께 쓰일 때 의미가 확장되는지 확인해 보세요.")

    return "\n".join(dict.fromkeys(notes))


@dataclass
class RowData:
    row_index: int
    word: str
    meaning: str
    example: str
    example_ko: str
    memo: str
    level: str
    tags: str
    qa_flag: str


class DictionaryWorkbook:
    def __init__(self) -> None:
        self.path: Optional[Path] = None
        self.wb = None
        self.ws = None
        self.headers: Dict[str, int] = {}
        self.data_rows: List[int] = []
        self.current_pos = 0

    def load(self, path: Path) -> None:
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb["words"] if "words" in self.wb.sheetnames else self.wb.active
        self.headers = self._ensure_columns()
        self.data_rows = [r for r in range(2, self.ws.max_row + 1) if norm(self.ws.cell(r, self.headers["word"]).value)]
        self.current_pos = 0
        self._style_sheet()

    def _header_map(self) -> Dict[str, int]:
        headers = {}
        for col in range(1, self.ws.max_column + 1):
            name = norm(self.ws.cell(1, col).value)
            if name:
                headers[name] = col
        return headers

    def _ensure_columns(self) -> Dict[str, int]:
        headers = self._header_map()
        missing = [c for c in REQUIRED_COLUMNS if c not in headers]
        if missing:
            raise ValueError(f"필수 컬럼이 없습니다: {missing}")
        next_col = self.ws.max_column + 1
        for col_name in ALL_COLUMNS:
            if col_name not in headers:
                self.ws.cell(1, next_col).value = col_name
                headers[col_name] = next_col
                next_col += 1
        return headers

    def _style_sheet(self) -> None:
        fill = PatternFill("solid", fgColor="1F4E78")
        font = Font(bold=True, color="FFFFFF")
        for cell in self.ws[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        widths = {"word":18, "meaning":24, "example":54, "example_ko":54, "memo":56, "level":10, "tags":22, "qa_flag":28}
        for name, col in self.headers.items():
            self.ws.column_dimensions[get_column_letter(col)].width = widths.get(name, 18)
        self.ws.freeze_panes = "A2"

    def count(self) -> int:
        return len(self.data_rows)

    def get_by_pos(self, pos: int) -> RowData:
        row = self.data_rows[pos]
        h = self.headers
        return RowData(
            row_index=row,
            word=norm(self.ws.cell(row, h["word"]).value),
            meaning=norm(self.ws.cell(row, h["meaning"]).value),
            example=norm(self.ws.cell(row, h["example"]).value),
            example_ko=norm(self.ws.cell(row, h["example_ko"]).value),
            memo=norm(self.ws.cell(row, h["memo"]).value),
            level=norm(self.ws.cell(row, h["level"]).value),
            tags=norm(self.ws.cell(row, h["tags"]).value),
            qa_flag=norm(self.ws.cell(row, h["qa_flag"]).value),
        )

    def get_current(self) -> RowData:
        return self.get_by_pos(self.current_pos)

    def update_current(self, example_ko: str, memo: str, level: str, tags: str) -> None:
        row = self.data_rows[self.current_pos]
        h = self.headers
        self.ws.cell(row, h["example_ko"]).value = example_ko.strip()
        self.ws.cell(row, h["memo"]).value = memo.strip()
        self.ws.cell(row, h["level"]).value = level.strip()
        self.ws.cell(row, h["tags"]).value = tags.strip()
        self.ws.cell(row, h["qa_flag"]).value = self.qa_flag_for_row(row)

    def qa_flag_for_row(self, row: int) -> str:
        h = self.headers
        flags = []
        if not norm(self.ws.cell(row, h["example_ko"]).value):
            flags.append("missing_example_ko")
        if not norm(self.ws.cell(row, h["memo"]).value):
            flags.append("missing_memo")
        if not norm(self.ws.cell(row, h["level"]).value):
            flags.append("missing_level")
        if not norm(self.ws.cell(row, h["tags"]).value):
            flags.append("missing_tags")
        ko = norm(self.ws.cell(row, h["example_ko"]).value)
        if ko and len(ko) < 8:
            flags.append("short_translation")
        return ", ".join(flags)

    def is_complete_row(self, row: int) -> bool:
        return not self.qa_flag_for_row(row)

    def qa_summary_for_rows(self, rows: List[int]) -> Dict[str, int]:
        summary = {"total":0, "complete":0, "missing_example_ko":0, "missing_memo":0, "missing_level":0, "missing_tags":0, "short_translation":0}
        for row in rows:
            summary["total"] += 1
            flag = self.qa_flag_for_row(row)
            self.ws.cell(row, self.headers["qa_flag"]).value = flag
            if not flag:
                summary["complete"] += 1
            for key in summary:
                if key not in ("total", "complete") and key in flag:
                    summary[key] += 1
        return summary

    def goto_pos(self, pos: int) -> None:
        self.current_pos = max(0, min(pos, len(self.data_rows) - 1))

    def goto_row_number(self, row_number: int) -> None:
        if row_number in self.data_rows:
            self.current_pos = self.data_rows.index(row_number)
            return
        candidates = [r for r in self.data_rows if r >= row_number]
        if candidates:
            self.current_pos = self.data_rows.index(candidates[0])

    def pos_for_row(self, row_number: int) -> Optional[int]:
        return self.data_rows.index(row_number) if row_number in self.data_rows else None

    def backup(self) -> Path:
        if not self.path:
            raise RuntimeError("파일 경로가 없습니다.")
        backup_dir = self.path.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = backup_dir / f"{self.path.stem}_backup_{stamp}{self.path.suffix}"
        shutil.copy2(self.path, backup_path)
        return backup_path

    def append_log(self, note: str) -> None:
        if "Translation_Log" in self.wb.sheetnames:
            ws = self.wb["Translation_Log"]
        else:
            ws = self.wb.create_sheet("Translation_Log")
            ws.append(["timestamp", "tool", "note"])
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "dictionary_editor_qt_v2.2", note])

    def save(self) -> None:
        self._style_sheet()
        self.append_log("Saved from PySide6 Dictionary Editor v2.2")
        self.wb.save(self.path)


class DictionaryEditorWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DailyVocab Master Dictionary Editor v2.2 Productivity")
        self.resize(1520, 880)

        self.model = DictionaryWorkbook()
        self.visible_rows: List[int] = []
        self.dirty = False

        self._build_actions()
        self._build_ui()
        self.setStyleSheet(DARK_STYLE)
        self._set_status("Excel 파일을 열어 주세요.")

    def _build_actions(self) -> None:
        save_action = QAction("Save", self)
        save_action.setShortcut(QKeySequence.Save)
        save_action.triggered.connect(self.save_file)
        save_next = QAction("Save Next", self)
        save_next.setShortcut("Ctrl+Return")
        save_next.triggered.connect(self.save_and_next)
        self.addAction(save_action)
        self.addAction(save_next)

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 10)

        toolbar = QHBoxLayout()
        self.open_btn = QPushButton("Excel 열기")
        self.open_btn.clicked.connect(self.open_file)
        self.save_btn = QPushButton("저장")
        self.save_btn.clicked.connect(self.save_file)
        self.qa_btn = QPushButton("QA 검사")
        self.qa_btn.setObjectName("Secondary")
        self.qa_btn.clicked.connect(self.run_qa)
        toolbar.addWidget(self.open_btn)
        toolbar.addWidget(self.save_btn)
        toolbar.addWidget(self.qa_btn)
        toolbar.addSpacing(12)
        toolbar.addWidget(QLabel("행 이동"))
        self.goto_input = QLineEdit()
        self.goto_input.setFixedWidth(90)
        self.goto_btn = QPushButton("Go")
        self.goto_btn.clicked.connect(self.goto_row)
        toolbar.addWidget(self.goto_input)
        toolbar.addWidget(self.goto_btn)
        self.save_state_label = QLabel("● No file")
        toolbar.addWidget(self.save_state_label)
        root.addLayout(toolbar)

        top = QFrame()
        top.setObjectName("Card")
        top_layout = QGridLayout(top)
        self.row_label = QLabel("Row -")
        self.word_label = QLabel("No word")
        self.word_label.setObjectName("TitleWord")
        self.meaning_label = QLabel("")
        self.meaning_label.setObjectName("Meaning")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress_label = QLabel("0 / 0")
        top_layout.addWidget(self.row_label, 0, 0)
        top_layout.addWidget(self.progress_label, 0, 1, alignment=Qt.AlignRight)
        top_layout.addWidget(self.word_label, 1, 0)
        top_layout.addWidget(self.meaning_label, 2, 0)
        top_layout.addWidget(self.progress, 1, 1, 2, 1)
        root.addWidget(top)

        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, 1)

        left = QFrame()
        left.setObjectName("Card")
        left_layout = QVBoxLayout(left)
        left_layout.addWidget(self._section_label("Batch Range"))
        batch_row = QHBoxLayout()
        self.batch_start = QLineEdit("1002")
        self.batch_end = QLineEdit("1101")
        self.apply_batch_btn = QPushButton("적용")
        self.apply_batch_btn.clicked.connect(self.apply_batch_filter)
        batch_row.addWidget(self.batch_start)
        batch_row.addWidget(QLabel("~"))
        batch_row.addWidget(self.batch_end)
        batch_row.addWidget(self.apply_batch_btn)
        left_layout.addLayout(batch_row)
        self.incomplete_only = QCheckBox("미완료만 보기")
        self.incomplete_only.stateChanged.connect(self.apply_batch_filter)
        left_layout.addWidget(self.incomplete_only)
        self.word_list = QListWidget()
        self.word_list.itemClicked.connect(self.open_word_list_item)
        left_layout.addWidget(self.word_list, 1)
        self.dashboard = QPlainTextEdit()
        self.dashboard.setReadOnly(True)
        self.dashboard.setMaximumHeight(140)
        left_layout.addWidget(self.dashboard)

        center = QFrame()
        center.setObjectName("Card")
        center_layout = QVBoxLayout(center)
        self.example_text = self._editor(readonly=True)
        self.example_ko_text = self._editor()
        self.memo_text = self._editor()
        center_layout.addWidget(self._section_label("Example"))
        center_layout.addWidget(self.example_text, 2)
        center_layout.addWidget(self._section_label("Example_KO"))
        center_layout.addWidget(self.example_ko_text, 2)
        center_layout.addWidget(self._section_label("Memo"))
        center_layout.addWidget(self.memo_text, 2)
        nav = QHBoxLayout()
        self.prev_btn = QPushButton("← 이전")
        self.prev_btn.setObjectName("Secondary")
        self.prev_btn.clicked.connect(self.prev_row)
        self.save_next_btn = QPushButton("저장 후 다음 →")
        self.save_next_btn.clicked.connect(self.save_and_next)
        self.next_btn = QPushButton("다음 →")
        self.next_btn.setObjectName("Secondary")
        self.next_btn.clicked.connect(self.next_row)
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.save_next_btn)
        nav.addWidget(self.next_btn)
        center_layout.addLayout(nav)

        right = QFrame()
        right.setObjectName("Card")
        right_layout = QVBoxLayout(right)
        right_layout.addWidget(self._section_label("Level"))
        self.level_combo = QComboBox()
        self.level_combo.addItems(LEVEL_OPTIONS)
        self.level_combo.currentTextChanged.connect(self._mark_dirty)
        right_layout.addWidget(self.level_combo)
        right_layout.addWidget(self._section_label("Tags"))
        self.tags_input = QLineEdit()
        self.tags_input.textChanged.connect(self._mark_dirty)
        right_layout.addWidget(self.tags_input)
        right_layout.addWidget(self._section_label("Assist"))
        self.assist_memo_btn = QPushButton("Memo 생성")
        self.assist_memo_btn.setObjectName("Assist")
        self.assist_memo_btn.clicked.connect(self.assist_memo)
        self.assist_tags_btn = QPushButton("Tags 추천")
        self.assist_tags_btn.setObjectName("Assist")
        self.assist_tags_btn.clicked.connect(self.assist_tags)
        self.assist_level_btn = QPushButton("Level 추천")
        self.assist_level_btn.setObjectName("Assist")
        self.assist_level_btn.clicked.connect(self.assist_level)
        self.assist_all_btn = QPushButton("Assist All")
        self.assist_all_btn.setObjectName("Assist")
        self.assist_all_btn.clicked.connect(self.assist_all)
        right_layout.addWidget(self.assist_memo_btn)
        right_layout.addWidget(self.assist_tags_btn)
        right_layout.addWidget(self.assist_level_btn)
        right_layout.addWidget(self.assist_all_btn)
        quick = QGridLayout()
        for i, tag in enumerate(TAG_OPTIONS):
            btn = QPushButton(tag)
            btn.setObjectName("Secondary")
            btn.clicked.connect(lambda checked=False, t=tag: self.add_tag(t))
            quick.addWidget(btn, i // 2, i % 2)
        right_layout.addLayout(quick)
        right_layout.addWidget(self._section_label("QA Flag"))
        self.qa_flag_label = QLabel("OK")
        self.qa_flag_label.setWordWrap(True)
        right_layout.addWidget(self.qa_flag_label)
        right_layout.addWidget(self._section_label("QA Summary"))
        self.qa_summary_text = QPlainTextEdit()
        self.qa_summary_text.setReadOnly(True)
        right_layout.addWidget(self.qa_summary_text, 1)

        splitter.addWidget(left)
        splitter.addWidget(center)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 2)

        self.setCentralWidget(central)
        self.setStatusBar(QStatusBar())
        self.example_ko_text.textChanged.connect(self._mark_dirty)
        self.memo_text.textChanged.connect(self._mark_dirty)

    def _section_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("SectionTitle")
        return label

    def _editor(self, readonly: bool = False) -> QPlainTextEdit:
        editor = QPlainTextEdit()
        editor.setReadOnly(readonly)
        editor.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        return editor

    def _set_status(self, text: str) -> None:
        self.statusBar().showMessage(text)

    def _mark_dirty(self) -> None:
        if self.model.path:
            self.dirty = True
            self.save_state_label.setText("● Unsaved")

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 열기", "", "Excel Files (*.xlsx)")
        if not path:
            return
        self.model.load(Path(path))
        self.visible_rows = list(self.model.data_rows)
        self.load_current_to_ui()
        self.apply_batch_filter()
        self.save_state_label.setText("● Saved")
        self._set_status(f"Loaded: {Path(path).name}")

    def apply_batch_filter(self) -> None:
        if not self.model.data_rows:
            return
        try:
            start = int(self.batch_start.text())
            end = int(self.batch_end.text())
        except ValueError:
            QMessageBox.warning(self, "Batch Range", "시작/끝 행 번호를 숫자로 입력하세요.")
            return
        rows = [r for r in self.model.data_rows if start <= r <= end]
        if self.incomplete_only.isChecked():
            rows = [r for r in rows if not self.model.is_complete_row(r)]
        self.visible_rows = rows
        self.refresh_word_list()
        if rows:
            pos = self.model.pos_for_row(rows[0])
            if pos is not None:
                self.model.goto_pos(pos)
                self.load_current_to_ui()
        self._set_status(f"Batch 적용 완료: {len(rows)}개 표시")

    def refresh_word_list(self) -> None:
        self.word_list.clear()
        complete = 0
        for row_number in self.visible_rows:
            pos = self.model.pos_for_row(row_number)
            if pos is None:
                continue
            row = self.model.get_by_pos(pos)
            ok = self.model.is_complete_row(row_number)
            if ok:
                complete += 1
            status = "✓" if ok else "•"
            item = QListWidgetItem(f"{status} {row.row_index}  {row.word}  |  {row.meaning}")
            item.setData(Qt.UserRole, row.row_index)
            self.word_list.addItem(item)
        total = len(self.visible_rows)
        rate = int(complete / total * 100) if total else 0
        remain = total - complete
        self.dashboard.setPlainText(f"Batch Dashboard\n---------------\nVisible: {total}\nComplete: {complete}\nRemaining: {remain}\nRate: {rate}%")

    def load_current_to_ui(self) -> None:
        row = self.model.get_current()
        self.row_label.setText(f"Excel Row {row.row_index}")
        self.progress_label.setText(f"{self.model.current_pos + 1} / {self.model.count()}")
        self.progress.setValue(int((self.model.current_pos + 1) / self.model.count() * 100))
        self.word_label.setText(row.word)
        self.meaning_label.setText(row.meaning)
        self.example_text.setPlainText(row.example)
        self.example_ko_text.blockSignals(True)
        self.memo_text.blockSignals(True)
        self.example_ko_text.setPlainText(row.example_ko)
        self.memo_text.setPlainText(row.memo)
        self.example_ko_text.blockSignals(False)
        self.memo_text.blockSignals(False)
        self.level_combo.setCurrentText(row.level if row.level in LEVEL_OPTIONS else "")
        self.tags_input.setText(row.tags)
        flag = row.qa_flag or self.model.qa_flag_for_row(row.row_index) or "OK"
        self.qa_flag_label.setText(flag)
        self.qa_flag_label.setStyleSheet("color: #34D399;" if flag == "OK" else "color: #F87171;")
        self.goto_input.setText(str(row.row_index))
        self.dirty = False
        self.save_state_label.setText("● Saved")

    def save_current_from_ui(self) -> None:
        self.model.update_current(self.example_ko_text.toPlainText(), self.memo_text.toPlainText(), self.level_combo.currentText(), self.tags_input.text())
        row = self.model.get_current()
        flag = row.qa_flag or "OK"
        self.qa_flag_label.setText(flag)
        self.qa_flag_label.setStyleSheet("color: #34D399;" if flag == "OK" else "color: #F87171;")
        self.dirty = False
        self.save_state_label.setText("● Saved")
        self.refresh_word_list()

    def open_word_list_item(self, item: QListWidgetItem) -> None:
        if self.dirty:
            self.save_current_from_ui()
        self.model.goto_row_number(item.data(Qt.UserRole))
        self.load_current_to_ui()

    def prev_row(self) -> None:
        if self.dirty:
            self.save_current_from_ui()
        self.model.goto_pos(self.model.current_pos - 1)
        self.load_current_to_ui()

    def next_row(self) -> None:
        if self.dirty:
            self.save_current_from_ui()
        self.model.goto_pos(self.model.current_pos + 1)
        self.load_current_to_ui()

    def save_and_next(self) -> None:
        self.save_current_from_ui()
        current_row = self.model.get_current().row_index
        if self.visible_rows and current_row in self.visible_rows:
            idx = self.visible_rows.index(current_row)
            next_row = self.visible_rows[min(idx + 1, len(self.visible_rows) - 1)]
            self.model.goto_row_number(next_row)
        else:
            self.model.goto_pos(self.model.current_pos + 1)
        self.load_current_to_ui()

    def goto_row(self) -> None:
        if self.dirty:
            self.save_current_from_ui()
        try:
            row_number = int(self.goto_input.text())
        except ValueError:
            QMessageBox.warning(self, "행 이동", "행 번호를 숫자로 입력하세요.")
            return
        self.model.goto_row_number(row_number)
        self.load_current_to_ui()

    def add_tag(self, tag: str) -> None:
        existing = [x.strip() for x in self.tags_input.text().split(",") if x.strip()]
        if tag not in existing:
            existing.append(tag)
        self.tags_input.setText(", ".join(existing))

    def assist_memo(self) -> None:
        row = self.model.get_current()
        self.memo_text.setPlainText(recommend_memo(row.word, row.meaning, row.example))

    def assist_tags(self) -> None:
        row = self.model.get_current()
        self.tags_input.setText(recommend_tags(row.word, row.meaning, row.example))

    def assist_level(self) -> None:
        row = self.model.get_current()
        self.level_combo.setCurrentText(recommend_level(row.word, row.example))

    def assist_all(self) -> None:
        self.assist_memo()
        self.assist_tags()
        self.assist_level()

    def run_qa(self) -> None:
        self.save_current_from_ui()
        rows = self.visible_rows if self.visible_rows else self.model.data_rows
        summary = self.model.qa_summary_for_rows(rows)
        complete_rate = int(summary["complete"] / summary["total"] * 100) if summary["total"] else 0
        self.qa_summary_text.setPlainText(
            "QA Summary\n----------\n"
            f"Scope rows: {summary['total']}\n"
            f"Complete: {summary['complete']}\n"
            f"Missing example_ko: {summary['missing_example_ko']}\n"
            f"Missing memo: {summary['missing_memo']}\n"
            f"Missing level: {summary['missing_level']}\n"
            f"Missing tags: {summary['missing_tags']}\n"
            f"Short translation: {summary['short_translation']}\n"
            f"Complete rate: {complete_rate}%"
        )
        self.progress.setValue(complete_rate)
        self.refresh_word_list()
        self._set_status("QA scan complete")

    def save_file(self) -> None:
        if not self.model.path:
            QMessageBox.warning(self, "저장", "먼저 Excel 파일을 열어 주세요.")
            return
        self.save_current_from_ui()
        backup_path = self.model.backup()
        self.model.save()
        self.save_state_label.setText("● Saved")
        self._set_status(f"Saved. Backup: {backup_path.name}")


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("DailyVocab Master Dictionary Editor")
    win = DictionaryEditorWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
