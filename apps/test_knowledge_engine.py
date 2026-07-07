from pathlib import Path
import sys
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
from engine.knowledge_engine import KnowledgeEngine
engine = KnowledgeEngine()
for word, meaning, example in [
    ("substance", "실질", "Some textile fabrics have little substance."),
    ("obsess", "집착하다", "Some people are obsessed with sports."),
    ("superior", "우수한", "This product is superior to the old one."),
]:
    s = engine.suggest_all(word, meaning, example)
    print("="*60)
    print(word, example)
    print(s.example_ko)
    print(s.memo)
    print(s.tags, s.level, s.confidence, s.reasons)
