def build_ai_tutor_text(word_row):
    word_id, word, meaning, example, example_ko, memo, level, tags = word_row

    lines = [
        f"🤖 AI Tutor: {word}",
        "",
        "1. 핵심 뜻",
        f"- {meaning or '뜻 정보가 없습니다.'}",
        "",
        "2. 예문",
        f"- {example or '예문이 없습니다.'}",
        "",
        "3. 자연스러운 해석",
        f"- {example_ko or '아직 example_ko가 비어 있습니다.'}",
        "",
        "4. 기억 팁",
    ]

    if memo:
        lines.append(f"- {memo}")
    else:
        lines.append(f"- '{word}'를 예문 속 문맥과 함께 기억해 보세요.")

    lines += [
        "",
        "5. 학습 정보",
        f"- Level: {level or '미지정'}",
        f"- Tags: {tags or '미지정'}",
        "",
        "6. 미니 퀴즈",
        f"- 다음 뜻에 해당하는 영어 단어는? → {meaning or '(뜻 없음)'}",
        "",
        "답을 떠올린 뒤 단어를 다시 확인해 보세요.",
    ]
    return "\n".join(lines)
