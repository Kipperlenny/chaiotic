"""AI interface module for interacting with language models."""

import json
import os
import hashlib
from typing import List, Dict, Any, Union

# Lazy initialization of OpenAI client
_openai_client = None

def _get_openai_client():
    """Lazily initialize and return the OpenAI client."""
    global _openai_client
    if _openai_client is None:
        api_key = os.environ.get("OPENAI_API_KEY", "")
        if api_key:
            try:
                from openai import OpenAI
                _openai_client = OpenAI(api_key=api_key)
            except ImportError:
                print("Warning: OpenAI package not installed. Cannot initialize client.")
        else:
            print("Warning: No OpenAI API key found in environment variables.")
    return _openai_client

def check_grammar_with_ai(content: str) -> List[Dict[str, Any]]:
    """
    Check grammar and spelling using AI service.
    
    Args:
        content: Content to check
        
    Returns:
        List of corrections
    """
    from chaiotic.config import GPT4O_MINI_MODEL
    
    # Get client (will be None if API key is not available)
    client = _get_openai_client()
    
    if not client:
        print("Warning: OpenAI client not available. Using mock data.")
        return get_mock_corrections()
    
    # Call OpenAI API
    try:
        # Prepare prompt
        system_message = (
            "Du bist ein professioneller deutscher Sprachredakteur mit Fokus auf Grammatik- und Rechtschreibkorrekturen. "
            "Analysiere den bereitgestellten deutschen Text und identifiziere Grammatik-, Rechtschreib- oder Stilfehler. "
            "Für jeden Fehler gib den Originaltext, die korrigierte Version und eine kurze Erklärung der Änderung auf Deutsch an. "
            "Achte darauf, nur echte Fehler zu korrigieren, nicht stilistische Präferenzen."
            "\n\nAntworte AUSSCHLIESSLICH im folgenden JSON-Format als Liste von Korrekturen:"
            "\n[\n  {\n    \"original\": \"Originaltext mit Fehler\",\n    \"corrected\": \"Korrigierter Text\","
            "\n    \"explanation\": \"Grund für die Korrektur\"\n  }\n]"
            "\n\nHalte die Liste leer, wenn keine Fehler gefunden wurden: []"
        )
        
        messages = [
            {"role": "system", "content": system_message},
            {"role": "user", "content": content}
        ]
        
        # Check cache
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'cache')
        os.makedirs(cache_dir, exist_ok=True)
        
        cache_key = hashlib.md5((json.dumps(messages) + GPT4O_MINI_MODEL).encode('utf-8')).hexdigest()
        cache_file = os.path.join(cache_dir, f"{cache_key}.json")
        
        # Try to load from cache first
        try:
            if os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cached_response = json.load(f)
                print("Using cached response for grammar check")
                return cached_response
        except Exception as e:
            print(f"Error reading from cache: {e}")
        
        # Call API if not in cache
        response = client.chat.completions.create(
            model=GPT4O_MINI_MODEL,
            messages=messages,
            temperature=0,
            max_tokens=4000
        )
        
        # Extract response
        response_text = response.choices[0].message.content
        
        # Debug print
        print(f"API Response: {response_text[:200]}...")
        
        # Process response
        corrections = _parse_corrections_response(response_text)
        
        # Save to cache
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(corrections, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error writing to cache: {e}")
        
        return corrections
    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        # Return mock data as fallback
        return get_mock_corrections()

def _parse_corrections_response(response_text: str) -> List[Dict[str, Any]]:
    """
    Parse the response from the AI to extract corrections.
    
    Args:
        response_text: Response from the AI
        
    Returns:
        List of corrections
    """
    # Clean up response
    response_text = response_text.strip()
    
    # Remove markdown code blocks if present
    if response_text.startswith('```') and response_text.endswith('```'):
        response_text = response_text[response_text.find('\n')+1:response_text.rfind('```')].strip()
    elif response_text.startswith('```json') and '```' in response_text[7:]:
        response_text = response_text[7:response_text.rfind('```')].strip()
    
    # Try to parse as JSON
    try:
        corrections = json.loads(response_text)
        
        # Validate that corrections is a list
        if isinstance(corrections, list):
            # Filter out invalid corrections and ensure required fields
            valid_corrections = []
            for correction in corrections:
                if isinstance(correction, dict) and 'original' in correction and 'corrected' in correction:
                    # Ensure explanation field exists
                    if 'explanation' not in correction:
                        correction['explanation'] = ""
                    valid_corrections.append(correction)
            return valid_corrections
        elif isinstance(corrections, dict) and 'corrections' in corrections:
            # Sometimes the API returns a dict with a corrections key
            corrections_list = corrections['corrections']
            if isinstance(corrections_list, list):
                # Filter and validate
                valid_corrections = []
                for correction in corrections_list:
                    if isinstance(correction, dict) and 'original' in correction and 'corrected' in correction:
                        if 'explanation' not in correction:
                            correction['explanation'] = ""
                        valid_corrections.append(correction)
                return valid_corrections
        
        # If we got here, the format is not what we expected
        print(f"Unexpected format for corrections: {type(corrections)}")
        print(f"Sample of corrections: {str(corrections)[:200]}")
        return []
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        print(f"Response text: {response_text[:200]}...")
        # Try to find JSON block surrounded by non-JSON text
        try:
            # Look for array of objects pattern
            import re
            json_match = re.search(r'\[\s*{.*}\s*\]', response_text, re.DOTALL)
            if json_match:
                try:
                    corrections = json.loads(json_match.group(0))
                    if isinstance(corrections, list):
                        valid_corrections = []
                        for correction in corrections:
                            if isinstance(correction, dict) and 'original' in correction and 'corrected' in correction:
                                if 'explanation' not in correction:
                                    correction['explanation'] = ""
                                valid_corrections.append(correction)
                        return valid_corrections
                except:
                    pass
        except:
            pass
        return []

def get_mock_corrections():
    """Get mock corrections for testing."""
    return [
        {
            "original": "Ich habe ein Fehler gemacht.",
            "corrected": "Ich habe einen Fehler gemacht.",
            "explanation": "Der Artikel für das maskuline Substantiv 'Fehler' im Akkusativ ist 'einen', nicht 'ein'."
        },
        {
            "original": "Die Katze ist auf dem Tisch gesprungen.",
            "corrected": "Die Katze ist auf den Tisch gesprungen.",
            "explanation": "Bei einer Bewegungsrichtung wird der Akkusativ (auf den Tisch) verwendet, nicht der Dativ."
        }
    ]

# ...existing code...