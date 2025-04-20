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

def get_grammar_system_prompt():
    """Get the system prompt for grammar checking in structured content analysis.
    
    Returns:
        String containing the system prompt
    """
    return """Du bist ein professioneller deutscher Sprachredakteur mit Fokus auf Grammatik- und Rechtschreibkorrekturen.
Analysiere den bereitgestellten deutschen Text und identifiziere Grammatik-, Rechtschreib- oder Stilfehler.
Deine Aufgabe ist, alle gefundenen Fehler zu korrigieren und zu erklären.

Gib deine Antwort im folgenden JSON-Format zurück:

```json
{
  "corrections": [
    {
      "original": "Text mit Fehler",
      "corrected": "Korrigierter Text",
      "explanation": "Erklärung der Korrektur"
    }
  ]
}
```

Falls keine Fehler im Text vorhanden sind, gib eine leere corrections-Liste zurück.
Achte besonders auf komplexe oder verschachtelte Textstrukturen und mögliche Formatierungsprobleme.
Analysiere auch die XML-Struktur, wenn verfügbar, um den Kontext des Textes besser zu verstehen.
"""

def get_grammar_user_prompt():
    """Get the user prompt for grammar checking in structured content analysis.
    
    Returns:
        String containing the user prompt template
    """
    return """Bitte analysiere den folgenden Text auf Grammatik-, Rechtschreib- und Stilfehler:

{content}

{xml_structure}

Korrigiere alle Fehler und erkläre die Änderungen.
"""

def process_content_batch(batch, client, model, system_prompt, user_prompt):
    """Process a batch of content items for grammar checking.
    
    Args:
        batch: List of content items to check
        client: OpenAI client
        model: Model to use
        system_prompt: System prompt
        user_prompt: User prompt template
        
    Returns:
        List of corrections
    """
    import json
    import hashlib
    import os
    from .config import DISABLE_CACHE
    
    corrections = []
    
    # Prepare batch content
    batch_content = ""
    xml_structure = ""
    item_map = {}
    
    # Debug info
    print(f"Preparing batch of {len(batch)} items for API request")
    
    for item in batch:
        item_id = item.get('id', '')
        item_content = item.get('content', '')
        
        # Add item to batch
        batch_content += f"\n--- ITEM {item_id} ({item.get('type', '')}) ---\n{item_content}\n\n"
        
        # Track item ID mapping for processing responses
        item_map[item_id] = item
        
        # If we have XML content, add it to the XML structure section
        if 'xml_content' in item:
            xml_structure += f"\n--- XML FOR ITEM {item_id} ---\n{item.get('xml_content', '')}\n"
    
    # Create final prompt
    xml_section = f"\nXML STRUCTURE:\n{xml_structure}" if xml_structure else ""
    formatted_user_prompt = user_prompt.format(content=batch_content, xml_structure=xml_section)
    
    # Cache hash based on content
    request_hash = hashlib.md5((system_prompt + formatted_user_prompt + model).encode('utf-8')).hexdigest()
    
    # Check for cached response
    request_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'api_requests')
    os.makedirs(request_dir, exist_ok=True)
    request_filename = os.path.join(request_dir, f"{request_hash}.json")
    
    # Try to use cached response if available
    if not DISABLE_CACHE and os.path.exists(request_filename):
        try:
            with open(request_filename, 'r', encoding='utf-8') as f:
                response = json.load(f)
                print(f"Using cached response for request hash '{request_hash}'")
        except Exception as e:
            print(f"Error reading cached response: {e}")
            response = call_openai_api_with_retry(client, model, system_prompt, formatted_user_prompt)
    else:
        # Call OpenAI API
        response = call_openai_api_with_retry(client, model, system_prompt, formatted_user_prompt)
    
    # Cache the response if needed
    if not DISABLE_CACHE and not os.path.exists(request_filename) and response:
        try:
            with open(request_filename, 'w', encoding='utf-8') as f:
                json.dump(response, f, ensure_ascii=False, indent=2)
            print(f"Response for request hash '{request_hash}' cached successfully.")
        except Exception as e:
            print(f"Error caching response: {e}")
    
    # Parse the API response
    try:
        # Extract corrections from JSON response
        parsed_corrections = extract_corrections_from_response(response)
        if parsed_corrections:
            corrections.extend(parsed_corrections)
    except Exception as e:
        print(f"Error parsing API response: {e}")
    
    return corrections

def call_openai_api_with_retry(client, model, system_prompt, user_prompt, max_retries=3, retry_delay=5):
    """Call OpenAI API with retry logic.
    
    Args:
        client: OpenAI client
        model: Model to use
        system_prompt: System prompt
        user_prompt: User prompt
        max_retries: Maximum number of retries
        retry_delay: Delay between retries in seconds
        
    Returns:
        API response text
    """
    import time
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    for attempt in range(max_retries):
        try:
            completion = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.1,
                max_tokens=4000,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )
            return completion.choices[0].message.content
        except Exception as e:
            print(f"API call failed (attempt {attempt+1}/{max_retries}): {e}")
            if attempt < max_retries - 1:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
    
    print("All API call attempts failed")
    return None

def extract_corrections_from_response(response):
    """Extract corrections from API response.
    
    Args:
        response: API response text
        
    Returns:
        List of corrections
    """
    import json
    import re
    
    if not response:
        return []
    
    # Try to extract JSON from the response
    try:
        # Check if response is already JSON
        parsed = json.loads(response)
        if isinstance(parsed, dict) and "corrections" in parsed:
            return parsed["corrections"]
    except json.JSONDecodeError:
        # Try to extract JSON from markdown code blocks
        json_blocks = re.findall(r'```(?:json)?\s*(.*?)```', response, re.DOTALL)
        for block in json_blocks:
            try:
                parsed = json.loads(block)
                if isinstance(parsed, dict) and "corrections" in parsed:
                    return parsed["corrections"]
            except json.JSONDecodeError:
                continue
        
        # Try to extract with regex as a last resort
        corrections_match = re.search(r'"corrections"\s*:\s*(\[.*?\])', response, re.DOTALL)
        if corrections_match:
            try:
                return json.loads(corrections_match.group(1))
            except json.JSONDecodeError:
                pass
    
    print("Failed to extract corrections from response")
    return []

def init_openai_client():
    """Initialize and return the OpenAI client."""
    try:
        from openai import OpenAI
        import os
        
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            print("Error: No OpenAI API key found")
            return None
            
        return OpenAI(api_key=api_key)
    except ImportError:
        print("Error: OpenAI package not installed")
        return None
    except Exception as e:
        print(f"Error initializing OpenAI client: {e}")
        return None

def load_config():
    """Load configuration."""
    try:
        import os
        import json
        
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        
        # Default config
        return {
            "ai": {
                "model": os.environ.get("GPT4O_MINI_MODEL", "gpt-4o-mini")
            }
        }
    except Exception as e:
        print(f"Error loading config: {e}")
        return {"ai": {"model": "gpt-4o-mini"}}

# Define default model
DEFAULT_MODEL = "gpt-4o-mini"