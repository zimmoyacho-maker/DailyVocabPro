#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from engine.knowledge_engine import KnowledgeEngine

engine = KnowledgeEngine()

samples = [
    ("achieve", "성취하다", "She achieved remarkable success through persistence."),
    ("achieve", "달성하다", "The company achieved remarkable growth."),
    ("provide", "제공하다", "The organization provides financial support."),
    ("prevent", "예방하다", "Regular exercise prevents disease."),
    ("improve", "개선하다", "The new system improves product quality."),
    ("reduce", "줄이다", "The policy reduced costs."),
    ("maintain", "유지하다", "We maintain high standards."),
    ("substance", "실질", "Some textile fabrics have little substance."),
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
