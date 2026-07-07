#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_KNOWLEDGE_DIR = PROJECT_ROOT / "knowledge"

TAG_KEYWORDS = {
    "Business": ["company", "market", "client", "contract", "revenue", "business", "office", "policy", "proposal", "sales", "product", "customer"],
    "Academic": ["theory", "philosophy", "research", "equation", "vector", "definition", "university", "student", "teacher"],
    "Education": ["school", "student", "teacher", "pupil", "college", "class", "lesson", "learn"],
    "Science": ["plant", "temperature", "biology", "geology", "disease", "velocity", "chemical", "research"],
    "Medical": ["doctor", "medicine", "disease", "health", "hospital", "drug", "patient", "pain", "treatment"],
    "Legal": ["court", "law", "police", "crime", "murder", "arrest", "custody", "judge", "legal"],
    "Travel": ["airport", "hotel", "restaurant", "ship", "train", "bus", "station", "route", "tour"],
    "Technology": ["software", "computer", "server", "protocol", "network", "program", "device", "data", "file"],
    "Daily": ["house", "room", "family", "food", "coffee", "clothes", "baby", "dog", "cat", "friend", "home"],
}

A1 = {"be", "do", "go", "come", "get", "make", "take", "see", "look", "want", "need", "like", "good", "bad", "big", "small", "day", "time", "home", "school", "book", "food"}
A2 = {"traffic", "wall", "calendar", "temperature", "damage", "floor", "rain", "purpose", "change", "escape", "ready", "poor", "plan", "record", "movie", "guide", "control", "plant", "wave", "school", "policy", "prison", "weapon", "unit", "sea", "sand"}

def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default

def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

@dataclass
class Suggestion:
    example_ko: str
    memo: str
    tags: str
    level: str
    confidence: int
    reasons: List[str]

class KnowledgeEngine:
    def __init__(self, knowledge_dir: Optional[Path] = None) -> None:
        self.knowledge_dir = knowledge_dir or DEFAULT_KNOWLEDGE_DIR
        self.tm_path = self.knowledge_dir / "translation_memory.json"
        self.phrase_path = self.knowledge_dir / "phrase_dictionary.json"
        self.golden_path = self.knowledge_dir / "golden_collection.json"
        self.noun_path = self.knowledge_dir / "noun_phrases.json"
        self.verb_path = self.knowledge_dir / "verb_patterns.json"
        self.translation_memory = load_json(self.tm_path, {})
        self.phrase_dictionary = load_json(self.phrase_path, {})
        self.golden_collection = load_json(self.golden_path, {"phrases": [], "translations": [], "examples": [], "memos": []})
        self.noun_phrases = load_json(self.noun_path, {})
        self.verb_patterns = load_json(self.verb_path, {})

    def save(self) -> None:
        save_json(self.tm_path, self.translation_memory)
        save_json(self.phrase_path, self.phrase_dictionary)
        save_json(self.golden_path, self.golden_collection)
        save_json(self.noun_path, self.noun_phrases)
        save_json(self.verb_path, self.verb_patterns)

    def find_phrase_matches(self, example: str) -> List[Tuple[str, Dict[str, Any]]]:
        lower = example.lower()
        return [(p, info) for p, info in sorted(self.phrase_dictionary.items(), key=lambda x: len(x[0]), reverse=True) if p.lower() in lower]

    def find_memory_matches(self, example: str) -> List[Tuple[str, Dict[str, Any]]]:
        lower = example.lower()
        return [(e, info) for e, info in sorted(self.translation_memory.items(), key=lambda x: len(x[0]), reverse=True) if e.lower() in lower]

    def find_noun_phrase_matches(self, example: str) -> List[Tuple[str, str]]:
        lower = example.lower()
        return [(eng, ko) for eng, ko in sorted(self.noun_phrases.items(), key=lambda x: len(x[0]), reverse=True) if eng.lower() in lower]

    def recommend_level(self, word: str, example: str) -> str:
        w = word.lower().strip()
        e = example.lower()
        if w in A1: return "A1"
        if w in A2: return "A2"
        if any(k in e for k in ["equation", "vector", "theory", "philosophy", "protocol", "government", "political"]): return "B2"
        if len(w) >= 12 or len(example) >= 150: return "C1"
        if len(w) >= 9 or len(example) >= 90: return "B2"
        return "B1"

    def recommend_tags(self, word: str, meaning: str, example: str) -> str:
        text = f"{word} {meaning} {example}".lower()
        tags = [tag for tag, keys in TAG_KEYWORDS.items() if any(k in text for k in keys)]
        return ", ".join((tags or ["Daily"])[:2])

    def recommend_memo(self, word: str, meaning: str, example: str) -> str:
        notes = []
        for phrase, info in self.find_phrase_matches(example):
            notes.append(f"{phrase} = {info.get('ko', '')}. {info.get('note', '')}".strip())
        if word.lower() in self.verb_patterns:
            notes.append(f"{word}는 목적어와 함께 쓰여 '{meaning}'의 의미를 만드는 동사입니다.")
        if not notes:
            notes.append(f"{word} = {meaning}. 예문 속 문맥과 함께 뜻을 기억하세요.")
        return "\n".join(dict.fromkeys(notes))

    def apply_memory(self, draft: str) -> str:
        out = draft
        sources = {}
        for eng, info in self.translation_memory.items():
            sources[eng] = info.get("ko", "")
        for eng, ko in self.noun_phrases.items():
            sources[eng] = ko
        for eng, ko in sorted(sources.items(), key=lambda x: len(x[0]), reverse=True):
            if ko:
                out = re.sub(re.escape(eng), ko, out, flags=re.IGNORECASE)
        return out

    def normalize_subject(self, subject: str) -> str:
        s = subject.strip()
        simple = {
            "he": "그는", "she": "그녀는", "they": "그들은", "we": "우리는", "i": "나는",
            "the company": "그 회사는", "the government": "정부는", "the system": "그 시스템은",
            "the program": "그 프로그램은", "the policy": "그 정책은",
            "some people": "어떤 사람들은", "many people": "많은 사람들은",
        }
        return simple.get(s.lower(), s)

    def translate_object_fragment(self, fragment: str) -> str:
        f = fragment.strip().strip(".")
        f = self.apply_memory(f)
        f = re.sub(r"\b(a|an|the)\s+", "", f, flags=re.IGNORECASE)
        return f.strip()

    def base_verb(self, verb: str) -> str:
        v = verb.lower()
        mapping = {
            "achieved": "achieve", "achieves": "achieve",
            "provided": "provide", "provides": "provide",
            "allowed": "allow", "allows": "allow",
            "prevented": "prevent", "prevents": "prevent",
            "required": "require", "requires": "require",
            "improved": "improve", "improves": "improve",
            "reduced": "reduce", "reduces": "reduce",
            "increased": "increase", "increases": "increase",
            "maintained": "maintain", "maintains": "maintain",
            "supported": "support", "supports": "support",
            "developed": "develop", "develops": "develop",
        }
        return mapping.get(v, v)

    def polish_korean(self, text: str) -> str:
        replacements = {
            "결정을 만들": "결정을 내리",
            "주의를 지불": "주의를 기울",
            "책임을 가지": "책임을 지",
            "성공을 달성": "성공을 거두",
            "진보를 만들": "진전을 이루",
            "성장을 달성": "성장을 이루",
            "성과를 달성": "성과를 내",
            "서비스을": "서비스를",
            "목표을": "목표를",
            "성장를": "성장을",
            "지원를": "지원을",
        }
        out = text
        for wrong, right in replacements.items():
            out = out.replace(wrong, right)
        return out

    def verb_template_translation(self, example: str) -> Optional[Tuple[str, int, List[str]]]:
        lower = example.strip().lower().rstrip(".")
        pattern = r"^(.+?)\s+(achieved|achieves|achieve|provided|provides|provide|allowed|allows|allow|prevented|prevents|prevent|required|requires|require|improved|improves|improve|reduced|reduces|reduce|increased|increases|increase|maintained|maintains|maintain|supported|supports|support|developed|develops|develop)\s+(.+)$"
        m = re.match(pattern, lower, flags=re.IGNORECASE)
        if not m:
            return None
        raw_subject, raw_verb, raw_object = m.groups()
        base = self.base_verb(raw_verb)
        subject = self.normalize_subject(raw_subject)
        obj = self.translate_object_fragment(raw_object)
        patterns = self.verb_patterns.get(base)
        if not patterns:
            return None
        selected = None
        for key, template in patterns.items():
            if key == "default":
                continue
            if key in raw_object.lower() or key in obj:
                selected = template
                break
        if not selected:
            selected = patterns.get("default")
        if not selected:
            return None
        draft = selected.format(subject=subject, object=obj, target="사람들")
        draft = self.polish_korean(draft)
        return draft + "  [초안]", 78, [f"Verb pattern: {base}", f"Object: {obj}"]

    def recommend_example_ko(self, word: str, meaning: str, example: str) -> Tuple[str, int, List[str]]:
        lower = example.strip().lower()
        reasons: List[str] = []
        confidence = 35
        memory_matches = self.find_memory_matches(example)
        phrase_matches = self.find_phrase_matches(example)
        noun_matches = self.find_noun_phrase_matches(example)
        for eng, info in memory_matches[:5]:
            reasons.append(f"TM: {eng} → {info.get('ko', '')}")
        for phrase, info in phrase_matches[:5]:
            reasons.append(f"Phrase: {phrase} → {info.get('ko', '')}")
        for eng, ko in noun_matches[:5]:
            reasons.append(f"Noun phrase: {eng} → {ko}")
        confidence += min(20, len(memory_matches) * 6)
        confidence += min(25, len(phrase_matches) * 8)
        confidence += min(20, len(noun_matches) * 5)

        patterns = [
            (r"some (.+) have little substance\.?$", "일부 {0}은 실질이나 내용이 부족합니다."),
            (r"some residents disputed the proposal, saying it was based more on emotion than fact\.?$", "일부 주민들은 그 제안이 사실보다 감정에 치우쳐 있다며 이의를 제기했습니다."),
            (r"some people are only famous within their city\.?$", "어떤 사람들은 자신이 사는 도시 안에서만 유명합니다."),
            (r"some people are obsessed with (.+)\.?$", "어떤 사람들은 {0}에 지나치게 몰두합니다."),
            (r"(.+) is superior to (.+)\.?$", "{0}은 {1}보다 우수합니다."),
            (r"(.+) was superior to (.+)\.?$", "{0}은 {1}보다 우수했습니다."),
            (r"(.+) is based on (.+)\.?$", "{0}은 {1}에 근거합니다."),
            (r"(.+) was based on (.+)\.?$", "{0}은 {1}에 근거했습니다."),
            (r"(.+) consists of (.+)\.?$", "{0}은 {1}로 이루어져 있습니다."),
            (r"(.+) resulted in (.+)\.?$", "{0}은 {1}이라는 결과를 낳았습니다."),
            (r"according to (.+), (.+)\.?$", "{0}에 따르면, {1}."),
            (r"although (.+), (.+)\.?$", "비록 {0}, {1}."),
            (r"because (.+), (.+)\.?$", "{0} 때문에, {1}."),
            (r"instead of (.+), (.+)\.?$", "{0} 대신에, {1}."),
        ]
        for pattern, template in patterns:
            m = re.match(pattern, lower, flags=re.IGNORECASE)
            if m:
                draft = template.format(*[self.translate_object_fragment(g.strip()) for g in m.groups()])
                draft = self.polish_korean(self.apply_memory(draft))
                reasons.append("Sentence pattern matched")
                return draft + "  [초안]", min(96, confidence + 25), reasons

        verb_result = self.verb_template_translation(example)
        if verb_result:
            draft, c, r = verb_result
            reasons.extend(r)
            return draft, min(92, max(confidence, c)), reasons

        if noun_matches:
            main = noun_matches[0][1]
            draft = f"{main}와 관련된 내용입니다. 문맥에 맞게 자연스럽게 다듬어 주세요.  [초안]"
            return draft, min(70, confidence), reasons

        if len(example.split()) <= 5:
            return f"{meaning}의 예로 쓰인 표현입니다.  [초안]", min(confidence, 55), reasons or ["Short example fallback"]

        return f"{word}가 '{meaning}'의 뜻으로 쓰인 예문입니다. 문맥에 맞게 한국어로 다듬어 주세요.  [초안]", min(confidence, 45), reasons or ["Last fallback"]

    def suggest_all(self, word: str, meaning: str, example: str) -> Suggestion:
        ko, confidence, reasons = self.recommend_example_ko(word, meaning, example)
        return Suggestion(
            example_ko=ko,
            memo=self.recommend_memo(word, meaning, example),
            tags=self.recommend_tags(word, meaning, example),
            level=self.recommend_level(word, example),
            confidence=confidence,
            reasons=reasons,
        )

    def learn_translation(self, english: str, korean: str, confidence: int = 90) -> None:
        english = english.strip()
        korean = korean.strip()
        if not english or not korean:
            return
        item = self.translation_memory.get(english, {})
        item["ko"] = korean
        item["count"] = int(item.get("count", 0)) + 1
        item["confidence"] = max(int(item.get("confidence", 0) or 0), confidence)
        item["last_used"] = now_iso()
        self.translation_memory[english] = item
        self.save()

    def learn_phrase(self, phrase: str, korean: str, note: str = "", confidence: int = 95) -> None:
        phrase = phrase.strip()
        korean = korean.strip()
        if not phrase or not korean:
            return
        item = self.phrase_dictionary.get(phrase, {})
        item["ko"] = korean
        item["note"] = note or item.get("note", "")
        item["confidence"] = max(int(item.get("confidence", 0) or 0), confidence)
        item["last_used"] = now_iso()
        self.phrase_dictionary[phrase] = item
        self.save()

    def add_golden(self, category: str, english: str, korean: str, note: str = "") -> None:
        category = category if category in self.golden_collection else "translations"
        self.golden_collection.setdefault(category, []).append({
            "english": english,
            "korean": korean,
            "note": note,
            "created_at": now_iso(),
        })
        self.save()
