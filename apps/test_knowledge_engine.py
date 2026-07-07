#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from engine.knowledge_engine import KnowledgeEngine

engine = KnowledgeEngine()

samples = [
    ("substance", "실질", "Some textile fabrics have little substance."),
    ("dispute", "이의를 제기하다", "Some residents disputed the proposal, saying it was based more on emotion than fact."),
    ("obsess", "집착하다", "Some people are obsessed with sports."),
    ("superior", "우수한", "This product is superior to the old one."),
]

for word, meaning, example in samples:
    s = engine.suggest_all(word, meaning, example)
    print("=" * 72)
    print("WORD:", word)
    print("EXAMPLE:", example)
    print("EXAMPLE_KO:", s.example_ko)
    print("MEMO:", s.memo)
    print("TAGS:", s.tags)
    print("LEVEL:", s.level)
    print("CONFIDENCE:", s.confidence)
    print("REASONS:", "; ".join(s.reasons))
