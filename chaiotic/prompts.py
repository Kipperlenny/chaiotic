def get_grammar_prompt(text, force_corrections=False):
    """Generate prompt for grammar checking."""
    base_prompt = (
        "Du bist ein professioneller deutscher Korrektor. Analysiere den folgenden Text "
        "auf Grammatik-, Rechtschreib- und Stilfehler. Antworte AUSSCHLIESSLICH im JSON-Format:\n"
        "{\n"
        '  "corrections": [\n'
        "    {\n"
        '      "original": "Text mit Fehler",\n'
        '      "corrected": "Korrigierter Text",\n'
        '      "explanation": "Erklärung der Korrektur"\n'
        "    }\n"
        "  ],\n"
        '  "corrected_full_text": "Der vollständige korrigierte Text"\n'
        "}\n\n"
        "Falls keine Korrekturen nötig sind, gib ein leeres corrections-Array zurück und den "
        "unveränderten Text als corrected_full_text. Hier ist der Text:\n\n"
    )

    if force_corrections:
        base_prompt = (
            "Du bist ein strenger deutscher Korrektor. Suche aktiv nach möglichen "
            "Verbesserungen im folgenden Text, auch bei Stil und Wortwahl. "
            "Antworte AUSSCHLIESSLICH im JSON-Format mit mindestens einer Korrektur:\n"
            "{\n"
            '  "corrections": [\n'
            "    {\n"
            '      "original": "Text mit Fehler",\n'
            '      "corrected": "Korrigierter Text",\n'
            '      "explanation": "Erklärung der Korrektur"\n'
            "    }\n"
            "  ],\n"
            '  "corrected_full_text": "Der vollständige korrigierte Text"\n'
            "}\n\n"
            "Hier ist der Text:\n\n"
        )

    return base_prompt + text