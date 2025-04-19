def get_grammar_prompt(text, force_corrections=False):
    """Generate a prompt for grammar checking."""
    
    prompt_prefix = """Du bist ein professioneller deutscher Sprachredakteur mit Fokus auf Grammatik- und Rechtschreibkorrekturen. Analysiere den bereitgestellten deutschen Text und identifiziere präzise Grammatik-, Rechtschreib- oder Stilfehler.

Formatiere deine Antwort als JSON-Objekt mit folgender Struktur:
{
  "corrections": [
    {
      "original": "Originaltext mit Fehler",
      "corrected": "Korrigierter Text",
      "explanation": "Grund für die Korrektur"
    }
  ],
  "corrected_full_text": "Der gesamte Text mit allen angewendeten Korrekturen."
}

Der Originaltext und der korrigierte Text sollten genau die fehlerhaften bzw. korrigierten Abschnitte enthalten. Achte auf folgendes:
1. "original" sollte den genauen Text enthalten, der korrigiert werden muss
2. "corrected" sollte die korrigierte Version genau dieses Textes enthalten
3. Nutze präzise, kleine Korrekturen statt ganzer Sätze zu ersetzen
4. Achte besonders auf Rechtschreibfehler in einzelnen Wörtern
5. Die Antwort MUSS ein valides JSON-Objekt sein

"""

    if force_corrections:
        prompt_prefix += """WICHTIG: Finde und korrigiere MINDESTENS ein paar kleine Fehler, selbst wenn der Text fast perfekt erscheint. Achte auf feine Nuancen in der Wortwahl, Zeichensetzung oder Grammatik, die verbessert werden könnten."""
    
    # Add special handling for ODT files
    prompt_prefix += """
Dieser Text stammt aus einer ODT-Datei. Bei der Korrektur ist es wichtig, dass deine gefundenen "original"-Abschnitte exakt mit dem Text übereinstimmen, da die Korrekturen mit Tracked Changes in der ODT-Datei angewendet werden.
"""
    
    prompt = f"{prompt_prefix}\n\nBitte analysiere den folgenden Text:\n\n{text}"
    return prompt