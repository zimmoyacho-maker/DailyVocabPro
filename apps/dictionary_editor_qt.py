#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DailyVocabPro Master Dictionary Editor v2.0

Run:
    pip install -r apps/requirements_dictionary_editor.txt
    python apps/dictionary_editor_qt.py
"""
from __future__ import annotations

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
    QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
    QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QSplitter, QStatusBar, QVBoxLayout, QWidget,
)

REQUIRED_COLUMNS = ["word", "meaning", "example"]
EDIT_COLUMNS = ["example_ko", "memo", "level", "tags"]
OPTIONAL_COLUMNS = ["qa_flag"]
ALL_COLUMNS = REQUIRED_COLUMNS + EDIT_COLUMNS + OPTIONAL_COLUMNS

LEVEL_OPTIONS = ["", "A1", "A2", "B1", "B2", "C1", "C2"]
TAG_OPTIONS = ["Daily", "Business", "Academic", "Education", "Travel", "Medical", "Legal", "Science", "Technology", "Conversation"]

DARK_STYLE = """
QMainWindow, QWidget { background-color: #111827; color: #E5E7EB; font-family: Segoe UI; font-size: 10.5pt; }
QFrame#Card { background-color: #1F2937; border: 1px solid #374151; border-radius: 10px; }
QLabel#TitleWord { font-size: 30px; font-weight: 700; color: #F9FAFB; }
QLabel#Meaning { font-size: 16px; color: #D1D5DB; }
QLabel#SectionTitle { font-size: 12px; font-weight: 700; color: #93C5FD; }
QPlainTextEdit, QLineEdit, QComboBox, QListWidget { background-color: #0B1220; color: #E5E7EB; border: 1px solid #374151; border-radius: 6px; padding: 6px; }
QPushButton { background-color: #2563EB; color: white; border: 0; border-radius: 7px; padding: 8px 12px; font-weight: 600; }
QPushButton:hover { background-color: #1D4ED8; }
QPushButton#Secondary { background-color: #374151; }
QPushButton#Secondary:hover { background-color: #4B5563; }
QProgressBar { border: 1px solid #374151; border-radius: 6px; text-align: center; background-color: #0B1220; }
QProgressBar::chunk { background-color: #2563EB; border-radius: 6px; }
QStatusBar { background-color: #0B1220; color: #D1D5DB; }
"""


def norm(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").replace("\u2003", " ").strip()


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
        widths = {"word": 18, "meaning": 24, "example": 54, "example_ko": 54, "memo": 56, "level": 10, "tags": 22, "qa_flag": 28}
        for name, col in self.headers.items():
            self.ws.column_dimensions[get_column_letter(col)].width = widths.get(name, 18)
        self.ws.freeze_panes = "A2"

    def count(self) -> int:
        return len(self.data_rows)

    def get_current(self) -> RowData:
        return self.get_by_pos(self.current_pos)

    def get_by_pos(self, pos: int) -> RowData:
        row = self.data_rows[pos]
        h = self.headers
        return RowData(row, norm(self.ws.cell(row, h["word"]).value), norm(self.ws.cell(row, h["meaning"]).value), norm(self.ws.cell(row, h["example"]).value), norm(self.ws.cell(row, h["example_ko"]).value), norm(self.ws.cell(row, h["memo"]).value), norm(self.ws.cell(row, h["level"]).value), norm(self.ws.cell(row, h["tags"]).value), norm(self.ws.cell(row, h["qa_flag"]).value))

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
        if not norm(self.ws.cell(row, h["example_ko"]).value): flags.append("missing_example_ko")
        if not norm(self.ws.cell(row, h["memo"]).value): flags.append("missing_memo")
        if not norm(self.ws.cell(row, h["level"]).value): flags.append("missing_level")
        if not norm(self.ws.cell(row, h["tags"]).value): flags.append("missing_tags")
        ko = norm(self.ws.cell(row, h["example_ko"]).value)
        if ko and len(ko) < 8: flags.append("short_translation")
        return ", ".join(flags)

    def qa_summary(self) -> Dict[str, int]:
        summary = {"total": 0, "complete": 0, "missing_example_ko": 0, "missing_memo": 0, "missing_level": 0, "missing_tags": 0, "short_translation": 0}
        for row in self.data_rows:
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

    def search(self, query: str) -> List[int]:
        q = query.lower().strip()
        if not q: return []
        fields = ["word", "meaning", "example", "example_ko", "memo", "level", "tags"]
        results = []
        for pos, row in enumerate(self.data_rows):
            text = " ".join(norm(self.ws.cell(row, self.headers[f]).value) for f in fields).lower()
            if q in text:
                results.append(pos)
        return results

    def backup(self) -> Path:
        if not self.path: raise RuntimeError("파일 경로가 없습니다.")
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
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "dictionary_editor_qt", note])
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="0F766E")
            cell.font = Font(bold=True, color="FFFFFF")
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 90

    def save(self) -> None:
        if not self.path: raise RuntimeError("저장할 파일이 없습니다.")
        self._style_sheet()
        self.append_log("Saved from PySide6 Dictionary Editor")
        self.wb.save(self.path)

    def save_as(self, path: Path) -> None:
        self.path = path
        self.save()


class DictionaryEditorWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DailyVocab Master Dictionary Editor v2.0")
        self.resize(1280, 820)
        self.model = DictionaryWorkbook()
        self.search_results: List[int] = []
        self.search_index = 0
        self.dirty = False
        self._build_actions()
        self._build_ui()
        self.setStyleSheet(DARK_STYLE)
        self._set_status("Excel 파일을 열어 주세요.")

    def _build_actions(self) -> None:
        open_action = QAction("Open", self); open_action.setShortcut(QKeySequence.Open); open_action.triggered.connect(self.open_file)
        save_action = QAction("Save", self); save_action.setShortcut(QKeySequence.Save); save_action.triggered.connect(self.save_file)
        self.addAction(open_action); self.addAction(save_action)

    def _section_label(self, text: str) -> QLabel:
        label = QLabel(text); label.setObjectName("SectionTitle"); return label

    def _editor(self, readonly: bool = False) -> QPlainTextEdit:
        editor = QPlainTextEdit(); editor.setReadOnly(readonly); editor.setLineWrapMode(QPlainTextEdit.WidgetWidth); return editor

    def _build_ui(self) -> None:
        central = QWidget(); root = QVBoxLayout(central); root.setContentsMargins(14,14,14,10); root.setSpacing(10)
        toolbar = QHBoxLayout()
        self.open_btn = QPushButton("Excel 열기"); self.open_btn.clicked.connect(self.open_file)
        self.save_btn = QPushButton("저장"); self.save_btn.clicked.connect(self.save_file)
        self.save_as_btn = QPushButton("다른 이름으로 저장"); self.save_as_btn.clicked.connect(self.save_as)
        self.qa_btn = QPushButton("QA 검사"); self.qa_btn.setObjectName("Secondary"); self.qa_btn.clicked.connect(self.run_qa)
        for w in [self.open_btn, self.save_btn, self.save_as_btn, self.qa_btn]: toolbar.addWidget(w)
        toolbar.addSpacing(16); toolbar.addWidget(QLabel("행 이동"))
        self.goto_input = QLineEdit(); self.goto_input.setFixedWidth(90); self.goto_btn = QPushButton("Go"); self.goto_btn.clicked.connect(self.goto_row)
        toolbar.addWidget(self.goto_input); toolbar.addWidget(self.goto_btn)
        toolbar.addSpacing(16); toolbar.addWidget(QLabel("검색"))
        self.search_input = QLineEdit(); self.search_input.returnPressed.connect(self.search)
        self.search_btn = QPushButton("찾기"); self.search_btn.clicked.connect(self.search)
        self.search_next_btn = QPushButton("다음"); self.search_next_btn.clicked.connect(self.search_next)
        toolbar.addWidget(self.search_input, 1); toolbar.addWidget(self.search_btn); toolbar.addWidget(self.search_next_btn)
        self.save_state_label = QLabel("● No file"); toolbar.addWidget(self.save_state_label)
        root.addLayout(toolbar)

        top_card = QFrame(); top_card.setObjectName("Card"); top_layout = QGridLayout(top_card); top_layout.setContentsMargins(18,14,18,14)
        self.row_label = QLabel("Row -"); self.word_label = QLabel("No word"); self.word_label.setObjectName("TitleWord"); self.meaning_label = QLabel(""); self.meaning_label.setObjectName("Meaning")
        self.progress = QProgressBar(); self.progress.setRange(0,100); self.progress_label = QLabel("0 / 0")
        top_layout.addWidget(self.row_label,0,0); top_layout.addWidget(self.progress_label,0,1,alignment=Qt.AlignRight)
        top_layout.addWidget(self.word_label,1,0); top_layout.addWidget(self.meaning_label,2,0); top_layout.addWidget(self.progress,1,1,2,1)
        root.addWidget(top_card)

        splitter = QSplitter(Qt.Horizontal); root.addWidget(splitter,1)
        left = QFrame(); left.setObjectName("Card"); left_layout = QVBoxLayout(left); left_layout.setContentsMargins(16,16,16,16)
        self.example_text = self._editor(readonly=True); self.example_ko_text = self._editor(); self.memo_text = self._editor()
        left_layout.addWidget(self._section_label("Example")); left_layout.addWidget(self.example_text,2)
        left_layout.addWidget(self._section_label("Example_KO")); left_layout.addWidget(self.example_ko_text,2)
        left_layout.addWidget(self._section_label("Memo")); left_layout.addWidget(self.memo_text,2)
        nav = QHBoxLayout(); self.prev_btn = QPushButton("← 이전"); self.prev_btn.setObjectName("Secondary"); self.prev_btn.clicked.connect(self.prev_row)
        self.save_next_btn = QPushButton("저장 후 다음 →"); self.save_next_btn.clicked.connect(self.save_and_next)
        self.next_btn = QPushButton("다음 →"); self.next_btn.setObjectName("Secondary"); self.next_btn.clicked.connect(self.next_row)
        for w in [self.prev_btn, self.save_next_btn, self.next_btn]: nav.addWidget(w)
        left_layout.addLayout(nav)

        right = QFrame(); right.setObjectName("Card"); right_layout = QVBoxLayout(right); right_layout.setContentsMargins(16,16,16,16)
        right_layout.addWidget(self._section_label("Level")); self.level_combo = QComboBox(); self.level_combo.addItems(LEVEL_OPTIONS); self.level_combo.currentTextChanged.connect(self._mark_dirty); right_layout.addWidget(self.level_combo)
        right_layout.addWidget(self._section_label("Tags")); self.tags_input = QLineEdit(); self.tags_input.textChanged.connect(self._mark_dirty); right_layout.addWidget(self.tags_input)
        quick = QGridLayout()
        for i, tag in enumerate(TAG_OPTIONS):
            btn = QPushButton(tag); btn.setObjectName("Secondary"); btn.clicked.connect(lambda checked=False, t=tag: self.add_tag(t)); quick.addWidget(btn, i//2, i%2)
        right_layout.addLayout(quick)
        right_layout.addWidget(self._section_label("QA Flag")); self.qa_flag_label = QLabel("OK"); self.qa_flag_label.setWordWrap(True); right_layout.addWidget(self.qa_flag_label)
        right_layout.addWidget(self._section_label("QA Summary")); self.qa_summary_text = QPlainTextEdit(); self.qa_summary_text.setReadOnly(True); right_layout.addWidget(self.qa_summary_text,1)
        right_layout.addWidget(self._section_label("Search Results")); self.search_list = QListWidget(); self.search_list.itemDoubleClicked.connect(self.open_selected_search_result); right_layout.addWidget(self.search_list,1)
        splitter.addWidget(left); splitter.addWidget(right); splitter.setStretchFactor(0,3); splitter.setStretchFactor(1,2)
        self.setCentralWidget(central); self.setStatusBar(QStatusBar())
        self.example_ko_text.textChanged.connect(self._mark_dirty); self.memo_text.textChanged.connect(self._mark_dirty)

    def _set_status(self, text: str) -> None: self.statusBar().showMessage(text)
    def _mark_dirty(self) -> None:
        if self.model.path: self.dirty = True; self.save_state_label.setText("● Unsaved")

    def _confirm_save_current(self) -> bool:
        if not self.dirty: return True
        result = QMessageBox.question(self, "저장 확인", "현재 행에 저장되지 않은 변경이 있습니다. 저장할까요?", QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if result == QMessageBox.Cancel: return False
        if result == QMessageBox.Yes: self.save_current_from_ui()
        return True

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 열기", "", "Excel Files (*.xlsx)")
        if not path: return
        try:
            self.model.load(Path(path)); self.dirty = False; self.load_current_to_ui(); self._set_status(f"Loaded: {Path(path).name}"); self.save_state_label.setText("● Saved")
        except Exception as e: QMessageBox.critical(self, "파일 열기 오류", str(e))

    def load_current_to_ui(self) -> None:
        if not self.model.data_rows: return
        row = self.model.get_current()
        self.row_label.setText(f"Excel Row {row.row_index}"); self.progress_label.setText(f"{self.model.current_pos + 1} / {self.model.count()}"); self.progress.setValue(int((self.model.current_pos + 1)/self.model.count()*100))
        self.word_label.setText(row.word); self.meaning_label.setText(row.meaning); self.example_text.setPlainText(row.example)
        self.example_ko_text.blockSignals(True); self.memo_text.blockSignals(True); self.example_ko_text.setPlainText(row.example_ko); self.memo_text.setPlainText(row.memo); self.example_ko_text.blockSignals(False); self.memo_text.blockSignals(False)
        self.level_combo.blockSignals(True); self.level_combo.setCurrentText(row.level if row.level in LEVEL_OPTIONS else ""); self.level_combo.blockSignals(False)
        self.tags_input.blockSignals(True); self.tags_input.setText(row.tags); self.tags_input.blockSignals(False)
        self.qa_flag_label.setText(row.qa_flag or "OK"); self.goto_input.setText(str(row.row_index)); self.dirty = False; self.save_state_label.setText("● Saved")

    def save_current_from_ui(self) -> None:
        if not self.model.data_rows: return
        self.model.update_current(self.example_ko_text.toPlainText(), self.memo_text.toPlainText(), self.level_combo.currentText(), self.tags_input.text())
        row = self.model.get_current(); self.qa_flag_label.setText(row.qa_flag or "OK"); self.dirty = False; self.save_state_label.setText("● Saved")

    def prev_row(self) -> None:
        if self._confirm_save_current(): self.model.goto_pos(self.model.current_pos - 1); self.load_current_to_ui()
    def next_row(self) -> None:
        if self._confirm_save_current(): self.model.goto_pos(self.model.current_pos + 1); self.load_current_to_ui()
    def save_and_next(self) -> None:
        self.save_current_from_ui(); self.model.goto_pos(self.model.current_pos + 1); self.load_current_to_ui()
    def goto_row(self) -> None:
        if not self._confirm_save_current(): return
        try: row_number = int(self.goto_input.text())
        except ValueError: QMessageBox.warning(self, "행 이동", "행 번호를 숫자로 입력하세요."); return
        self.model.goto_row_number(row_number); self.load_current_to_ui()

    def search(self) -> None:
        query = self.search_input.text().strip()
        if not query: return
        self.search_results = self.model.search(query); self.search_index = 0; self.search_list.clear()
        for pos in self.search_results[:300]:
            row = self.model.get_by_pos(pos); self.search_list.addItem(f"{row.row_index} | {row.word} | {row.meaning}")
        if not self.search_results: QMessageBox.information(self, "검색", "검색 결과가 없습니다."); return
        if self._confirm_save_current(): self.model.goto_pos(self.search_results[0]); self.load_current_to_ui(); self._set_status(f"Search: {len(self.search_results)} result(s)")

    def search_next(self) -> None:
        if not self.search_results: self.search(); return
        if self._confirm_save_current(): self.search_index = (self.search_index + 1) % len(self.search_results); self.model.goto_pos(self.search_results[self.search_index]); self.load_current_to_ui()
    def open_selected_search_result(self) -> None:
        item = self.search_list.currentItem()
        if not item: return
        row_number = int(item.text().split("|")[0].strip())
        if self._confirm_save_current(): self.model.goto_row_number(row_number); self.load_current_to_ui()
    def add_tag(self, tag: str) -> None:
        existing = [x.strip() for x in self.tags_input.text().split(",") if x.strip()]
        if tag not in existing: existing.append(tag)
        self.tags_input.setText(", ".join(existing))

    def run_qa(self) -> None:
        if not self.model.data_rows: return
        self.save_current_from_ui(); summary = self.model.qa_summary(); complete_rate = int(summary["complete"]/summary["total"]*100) if summary["total"] else 0
        lines = ["QA Summary", "----------", f"Total: {summary['total']}", f"Complete: {summary['complete']}", f"Missing example_ko: {summary['missing_example_ko']}", f"Missing memo: {summary['missing_memo']}", f"Missing level: {summary['missing_level']}", f"Missing tags: {summary['missing_tags']}", f"Short translation: {summary['short_translation']}", f"Complete rate: {complete_rate}%"]
        self.qa_summary_text.setPlainText("\n".join(lines)); self.progress.setValue(complete_rate); row = self.model.get_current(); self.qa_flag_label.setText(row.qa_flag or "OK"); self._set_status("QA scan complete")

    def save_file(self) -> None:
        if not self.model.path: QMessageBox.warning(self, "저장", "먼저 Excel 파일을 열어 주세요."); return
        try:
            self.save_current_from_ui(); backup_path = self.model.backup(); self.model.save(); self.dirty = False; self.save_state_label.setText("● Saved"); self._set_status(f"Saved. Backup: {backup_path.name}"); QMessageBox.information(self, "저장 완료", f"저장되었습니다.\n백업: {backup_path.name}")
        except Exception as e: QMessageBox.critical(self, "저장 오류", str(e))
    def save_as(self) -> None:
        if not self.model.data_rows: return
        path, _ = QFileDialog.getSaveFileName(self, "다른 이름으로 저장", "", "Excel Files (*.xlsx)")
        if not path: return
        try: self.save_current_from_ui(); self.model.save_as(Path(path)); self.dirty = False; self.save_state_label.setText("● Saved"); self._set_status(f"Saved as: {Path(path).name}")
        except Exception as e: QMessageBox.critical(self, "저장 오류", str(e))


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("DailyVocab Master Dictionary Editor")
    win = DictionaryEditorWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
