#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations
import json, re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_KNOWLEDGE_DIR = PROJECT_ROOT / "knowledge"

DEFAULT_PHRASES = {
    "be superior to": {"ko": "~보다 우수하다", "note": "superior는 than이 아니라 to와 함께 사용합니다.", "confidence": 98},
    "be based on": {"ko": "~에 근거하다", "note": "근거, 기준, 원인을 설명할 때 사용합니다.", "confidence": 98},
    "consist of": {"ko": "~로 이루어져 있다", "note": "구성 요소를 설명할 때 사용합니다.", "confidence": 98},
    "result in": {"ko": "~라는 결과를 낳다 / ~을 초래하다", "note": "결과를 나타내는 중요 표현입니다.", "confidence": 96},
    "according to": {"ko": "~에 따르면", "note": "출처나 기준을 제시할 때 사용합니다.", "confidence": 99},
    "rather than": {"ko": "~라기보다 / ~보다는", "note": "대조와 선택을 나타냅니다.", "confidence": 96},
    "as a result": {"ko": "그 결과", "note": "결과를 연결하는 표현입니다.", "confidence": 97},
    "be obsessed with": {"ko": "~에 지나치게 몰두하다 / ~에 집착하다", "note": "with와 함께 쓰이는 표현입니다.", "confidence": 98},
    "be charged with": {"ko": "~혐의로 기소되다", "note": "법률·뉴스 문맥에서 자주 쓰입니다.", "confidence": 97},
    "have little substance": {"ko": "실질이 부족하다", "note": "substance가 '실질, 내용'의 뜻으로 쓰인 표현입니다.", "confidence": 96}
}
DEFAULT_MEMORY = {
    "textile fabric": {"ko": "직물", "count": 1, "confidence": 90, "last_used": ""},
    "textile fabrics": {"ko": "직물", "count": 1, "confidence": 90, "last_used": ""}
}
TAG_KEYWORDS = {
    "Business": ["company","market","client","contract","revenue","business","office","policy","proposal","sales"],
    "Academic": ["theory","philosophy","research","equation","vector","definition","university","student","teacher"],
    "Education": ["school","student","teacher","pupil","college","class","lesson","learn"],
    "Science": ["plant","temperature","biology","geology","disease","velocity","chemical"],
    "Medical": ["doctor","medicine","disease","health","hospital","drug","patient","pain"],
    "Legal": ["court","law","police","crime","murder","arrest","custody","judge","legal"],
    "Travel": ["airport","hotel","restaurant","ship","train","bus","station","route","tour"],
    "Technology": ["software","computer","server","protocol","network","program","device","data","file"],
    "Daily": ["house","room","family","food","coffee","clothes","baby","dog","cat","friend","home"]
}
A1={"be","do","go","come","get","make","take","see","look","want","need","like","good","bad","big","small","day","time","home","school","book","food"}
A2={"traffic","wall","calendar","temperature","damage","floor","rain","purpose","change","escape","ready","poor","plan","record","movie","guide","control","plant","wave","school","policy","prison","weapon","unit","sea","sand"}

def now_iso(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def load_json(path: Path, default: Any):
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
        return default
    try: return json.loads(path.read_text(encoding="utf-8"))
    except Exception: return default
def save_json(path: Path, data: Any):
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
    def __init__(self, knowledge_dir: Optional[Path]=None):
        self.knowledge_dir = knowledge_dir or DEFAULT_KNOWLEDGE_DIR
        self.tm_path = self.knowledge_dir / "translation_memory.json"
        self.phrase_path = self.knowledge_dir / "phrase_dictionary.json"
        self.golden_path = self.knowledge_dir / "golden_collection.json"
        self.translation_memory = load_json(self.tm_path, DEFAULT_MEMORY)
        self.phrase_dictionary = load_json(self.phrase_path, DEFAULT_PHRASES)
        self.golden_collection = load_json(self.golden_path, {"phrases":[],"translations":[],"examples":[],"memos":[]})
    def save(self):
        save_json(self.tm_path, self.translation_memory); save_json(self.phrase_path, self.phrase_dictionary); save_json(self.golden_path, self.golden_collection)
    def find_phrase_matches(self, example):
        lower=example.lower()
        return [(p,i) for p,i in sorted(self.phrase_dictionary.items(), key=lambda x:len(x[0]), reverse=True) if p.lower() in lower]
    def find_memory_matches(self, example):
        lower=example.lower()
        return [(e,i) for e,i in sorted(self.translation_memory.items(), key=lambda x:len(x[0]), reverse=True) if e.lower() in lower]
    def recommend_level(self, word, example):
        w=word.lower().strip(); e=example.lower()
        if w in A1: return "A1"
        if w in A2: return "A2"
        if any(k in e for k in ["equation","vector","theory","philosophy","protocol","government","political"]): return "B2"
        if len(w)>=12 or len(example)>=150: return "C1"
        if len(w)>=9 or len(example)>=90: return "B2"
        return "B1"
    def recommend_tags(self, word, meaning, example):
        text=f"{word} {meaning} {example}".lower()
        tags=[t for t, keys in TAG_KEYWORDS.items() if any(k in text for k in keys)]
        return ", ".join((tags or ["Daily"])[:2])
    def recommend_memo(self, word, meaning, example):
        notes=[f"{p} = {i.get('ko','')}. {i.get('note','')}".strip() for p,i in self.find_phrase_matches(example)]
        if not notes: notes.append(f"{word} = {meaning}. 예문 속 문맥과 함께 뜻을 기억하세요.")
        return "\n".join(dict.fromkeys(notes))
    def apply_memory(self, draft):
        out=draft
        for eng, info in sorted(self.translation_memory.items(), key=lambda x:len(x[0]), reverse=True):
            if info.get("ko"): out = re.sub(re.escape(eng), info["ko"], out, flags=re.IGNORECASE)
        return out
    def recommend_example_ko(self, word, meaning, example):
        lower=example.strip().lower(); reasons=[]; confidence=30
        mem=self.find_memory_matches(example); phr=self.find_phrase_matches(example)
        for e,i in mem[:5]: reasons.append(f"TM: {e} → {i.get('ko','')}")
        for p,i in phr[:5]: reasons.append(f"Phrase: {p} → {i.get('ko','')}")
        confidence += min(25, len(mem)*7) + min(30, len(phr)*10)
        patterns=[
            (r"some (.+) have little substance\.?$","일부 {0}은 실질이나 내용이 부족합니다."),
            (r"some residents disputed the proposal, saying it was based more on emotion than fact\.?$","일부 주민들은 그 제안이 사실보다 감정에 치우쳐 있다며 이의를 제기했습니다."),
            (r"some people are only famous within their city\.?$","어떤 사람들은 자신이 사는 도시 안에서만 유명합니다."),
            (r"some people are obsessed with (.+)\.?$","어떤 사람들은 {0}에 지나치게 몰두합니다."),
            (r"(.+) is superior to (.+)\.?$","{0}은 {1}보다 우수합니다."),
            (r"(.+) was superior to (.+)\.?$","{0}은 {1}보다 우수했습니다."),
            (r"(.+) is based on (.+)\.?$","{0}은 {1}에 근거합니다."),
            (r"(.+) consists of (.+)\.?$","{0}은 {1}로 이루어져 있습니다."),
            (r"(.+) resulted in (.+)\.?$","{0}은 {1}이라는 결과를 낳았습니다."),
            (r"according to (.+), (.+)\.?$","{0}에 따르면, {1}.")
        ]
        for pat, temp in patterns:
            m=re.match(pat, lower, flags=re.IGNORECASE)
            if m:
                draft=self.apply_memory(temp.format(*[g.strip() for g in m.groups()]))
                reasons.append("Pattern matched")
                return draft+"  [초안]", min(95, confidence+25), reasons
        if len(example.split())<=5:
            return f"{meaning}의 예로 쓰인 표현입니다.  [초안]", min(confidence,55), reasons or ["Short example fallback"]
        return f"이 예문에서 '{word}'는 '{meaning}'의 의미로 쓰였습니다. 자연스러운 한국어 문장으로 다듬어 주세요.  [초안]", min(confidence,50), reasons or ["General fallback"]
    def suggest_all(self, word, meaning, example):
        ko, conf, reasons = self.recommend_example_ko(word, meaning, example)
        return Suggestion(ko, self.recommend_memo(word, meaning, example), self.recommend_tags(word, meaning, example), self.recommend_level(word, example), conf, reasons)
    def learn_translation(self, english, korean, confidence=90):
        english=english.strip(); korean=korean.strip()
        if not english or not korean: return
        item=self.translation_memory.get(english,{})
        item["ko"]=korean; item["count"]=int(item.get("count",0))+1; item["confidence"]=max(int(item.get("confidence",0) or 0), confidence); item["last_used"]=now_iso()
        self.translation_memory[english]=item; self.save()
    def learn_phrase(self, phrase, korean, note="", confidence=95):
        phrase=phrase.strip(); korean=korean.strip()
        if not phrase or not korean: return
        item=self.phrase_dictionary.get(phrase,{})
        item["ko"]=korean; item["note"]=note or item.get("note",""); item["confidence"]=max(int(item.get("confidence",0) or 0), confidence); item["last_used"]=now_iso()
        self.phrase_dictionary[phrase]=item; self.save()

    def add_golden(self, category, english, korean, note=""):
        category = category if category in self.golden_collection else "translations"
        self.golden_collection.setdefault(category, []).append({
            "english": english,
            "korean": korean,
            "note": note,
            "created_at": now_iso()
        })
        self.save()
