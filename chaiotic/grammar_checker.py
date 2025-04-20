"""Grammar checking module for documents."""

import json
import re
import time
import os
import hashlib
import difflib
import shutil
from typing import List, Dict, Any, Optional, Union

from chaiotic.text_utils import split_text_into_chunks

from .config import DISABLE_CACHE, GPT4O_MINI_MODEL
from .ai_interface import check_grammar_with_ai
from .prompts import init_openai_client, load_config, DEFAULT_MODEL, process_content_batch, extract_corrections_from_response

# Initialize OpenAI client only if API key is available
client = None
try:
    from openai import OpenAI
    api_key = os.environ.get("OPENAI_API_KEY")
    if api_key:
        client = OpenAI(api_key=api_key)
    else:
        print("Warning: No OpenAI API key found in environment variables.")
        print("Set the OPENAI_API_KEY environment variable to use OpenAI services.")
        print("Using mock data for grammar checking.")
except ImportError:
    print("Warning: OpenAI package not installed. Using mock data for grammar checking.")

def check_grammar(text, structured_content=None, use_structured=False, checkpoint_handler=None):
    """Check grammar and suggest corrections.
    
    Args:
        text: Text content to check
        structured_content: Optional structured content (paragraphs with IDs)
        use_structured: Whether to use structured approach
        checkpoint_handler: Optional checkpoint handler for saving state
        
    Returns:
        List of suggested corrections
    """
    # Ensure text is not None and has content
    if not text or not isinstance(text, str) or len(text.strip()) == 0:
        print("Error: Empty or invalid text provided to check_grammar")
        return []
        
    print(f"Checking grammar for text ({len(text)} characters)")
    
    # Check if API key is configured
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("API key not found. Please set the OPENAI_API_KEY environment variable.")
        return []
    
    # Try to restore from checkpoint
    if checkpoint_handler:
        checkpoint_data = checkpoint_handler.load_checkpoint()
        if checkpoint_data and 'corrections' in checkpoint_data:
            print(f"Resumed from checkpoint with {len(checkpoint_data['corrections'])} corrections")
            return checkpoint_data['corrections']
    
    # Determine which approach to use
    if use_structured and structured_content:
        print("Using structured approach for grammar checking")
        corrections = check_grammar_structured(text, structured_content, checkpoint_handler)
    else:
        print("Using standard approach for grammar checking")
        corrections = check_grammar_standard(text, checkpoint_handler)
    
    # Debug output
    print(f"Grammar check completed with {len(corrections) if corrections else 0} corrections")
    
    # Add file type info
    if corrections and isinstance(corrections, list):
        for correction in corrections:
            if isinstance(correction, dict) and 'metadata' not in correction:
                correction['metadata'] = {'file_type': 'odt'}

    return corrections

def check_grammar_standard(text, checkpoint_handler=None):
    """Standard approach for grammar checking - process all text in one go."""
    from .prompts import get_grammar_prompt
    
    try:
        # Use smaller chunks if text is large
        if len(text) > 10000:
            print(f"Text is {len(text)} characters, processing in chunks")
            chunks = split_into_chunks(text, 5000)
            all_corrections = []
            
            for i, chunk in enumerate(chunks):
                print(f"Processing chunk {i+1}/{len(chunks)} ({len(chunk)} characters)")
                prompt = get_grammar_prompt(chunk)
                response = call_openai_api(prompt)
                
                # Parse the response
                chunk_corrections = parse_corrections(response, chunk)
                if chunk_corrections:
                    all_corrections.extend(chunk_corrections)
                    
                    # Save checkpoint after each chunk
                    if checkpoint_handler:
                        checkpoint_handler.save_checkpoint(
                            total_elements=len(chunks), 
                            processed_elements=i+1, 
                            corrections=all_corrections
                        )
                
            # Force one more check if we didn't get any corrections
            if not all_corrections:
                print("No corrections found in chunks, trying again with full text...")
                prompt = get_grammar_prompt(text, True)
                response = call_openai_api(prompt, temperature=0.2)
                all_corrections = parse_corrections(response, text)
            
            return all_corrections
        else:
            # Process text in one go for smaller documents
            print("Processing full text as single chunk")
            prompt = get_grammar_prompt(text)
            response = call_openai_api(prompt)
            
            corrections = parse_corrections(response, text)
            
            # Try again with different settings if we didn't get any corrections
            if not corrections:
                print("No corrections found, trying with different prompt...")
                prompt = get_grammar_prompt(text, force_corrections=True)
                response = call_openai_api(prompt, temperature=0.2)
                corrections = parse_corrections(response, text)
            
            # Save checkpoint
            if checkpoint_handler and corrections:
                checkpoint_handler.save_checkpoint(
                    total_elements=1, 
                    processed_elements=1, 
                    corrections=corrections
                )
                
            return corrections
    except Exception as e:
        print(f"Error in check_grammar_standard: {e}")
        import traceback
        traceback.print_exc()
        return []

def call_openai_api(prompt, model=GPT4O_MINI_MODEL, temperature=0, retries=3, delay=5):
    """Call the OpenAI API with caching."""
    from time import sleep
    
    # Create a hash based on the prompt content and model
    request_hash = hashlib.md5((prompt + model).encode('utf-8')).hexdigest()
    
    # Ensure api_requests directory exists
    request_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'api_requests')
    os.makedirs(request_dir, exist_ok=True)
    request_filename = os.path.join(request_dir, f"{request_hash}.json")
    
    # Check if we already have a cached response
    if not DISABLE_CACHE and os.path.exists(request_filename):
        try:
            with open(request_filename, 'r', encoding='utf-8') as f:
                response = json.load(f)
            print(f"Using cached response for request hash '{request_hash}'")
            return response
        except Exception as e:
            print(f"Error reading cached response: {e}")
    
    # Make API call if no cache or cache reading failed
    for attempt in range(retries):
        try:
            messages = [
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ]
            
            # Define appropriate max_tokens
            max_tokens = 4096
            
            completion = client.chat.completions.create(
                model=model,
                messages=messages,
                stream=False,
                temperature=temperature,
                max_tokens=max_tokens,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )
            response = completion.choices[0].message.content
            
            # Cache the response
            if not DISABLE_CACHE:
                try:
                    with open(request_filename, 'w', encoding='utf-8') as f:
                        json.dump(response, f, indent=4, ensure_ascii=False)
                    print(f"Response for request hash '{request_hash}' cached successfully.")
                except Exception as cache_err:
                    print(f"Warning: Failed to cache response: {cache_err}")
                
            return response
        except Exception as e:
            print(f"Error calling OpenAI API (attempt {attempt + 1}/{retries}): {e}")
            if attempt < retries - 1:
                sleep(delay)
            else:
                # Use the ai_interface module as fallback
                print("Using ai_interface as fallback for OpenAI API call")
                return fallback_grammar_check(prompt)

def fallback_grammar_check(prompt):
    """Use the existing ai_interface module as a fallback for grammar checking."""
    # Extract the actual content to check from the prompt
    content_match = re.search(r'Bitte analysiere den folgenden Text:\s*\n\n(.*)', prompt, re.DOTALL)
    if content_match:
        content = content_match.group(1)
    else:
        content = prompt
    
    # Use the existing AI interface
    corrections = check_grammar_with_ai(content)
    
    # Format the response to match what's expected by parse_corrections
    response = {
        "corrections": corrections,
        "corrected_full_text": content  # This will be updated later
    }
    
    # Convert to JSON string
    return json.dumps(response, ensure_ascii=False, indent=2)

def display_corrections(corrections):
    """Display corrections in a readable format."""
    return display_corrections_table(corrections)

def grammar_check(content):
    """Check grammar and spelling using GPT-4o-mini."""
    # Split content into appropriate chunks
    MAX_TOKENS_PER_CHUNK = 100000  # Leave room for system message and output
    chunks = split_text_into_chunks(content, MAX_TOKENS_PER_CHUNK, model=GPT4O_MINI_MODEL)
    
    if len(chunks) == 1:
        # If only one chunk, process normally
        return process_grammar_chunk(chunks[0])
    else:
        # For multiple chunks, process each separately and combine results
        print(f"Text is large, splitting into {len(chunks)} chunks for processing...")
        all_corrections = []
        corrected_content = ""
        
        for i, chunk in enumerate(chunks):
            print(f"Processing chunk {i+1}/{len(chunks)}...")
            chunk_result = process_grammar_chunk(chunk)
            
            if chunk_result and "corrections" in chunk_result:
                all_corrections.extend(chunk_result["corrections"])
                corrected_content += chunk_result["corrected_full_text"] + "\n"
        
        # Return combined results
        return {
            "corrections": all_corrections,
            "corrected_full_text": corrected_content.strip()
        }

def process_grammar_chunk(content_chunk):
    """Process a single chunk for grammar checking."""
    return _call_grammar_api(
        content_chunk,
        full_json_response=True,
        element_type=None
    )

def check_grammar_of_paragraph(content, element_type):
    """Check grammar of a single paragraph or element."""
    # Skip if content is empty
    if not content or len(content.strip()) < 3:
        return {
            "corrected": content,
            "has_corrections": False
        }
    
    return _call_grammar_api(
        content,
        full_json_response=False,
        element_type=element_type
    )

def _call_grammar_api(content, full_json_response=True, element_type=None):
    """Unified helper function to call OpenAI API for grammar checking.
    
    Args:
        content: The text content to check
        full_json_response: Whether to request the full corrections array (True) or simplified format (False)
        element_type: The type of element being checked (paragraph, heading, etc.)
    """
    # Skip if content is empty
    if not content or len(content.strip()) < 3:
        if full_json_response:
            return {
                "corrections": [],
                "corrected_full_text": content
            }
        else:
            return {
                "corrected": content,
                "has_corrections": False
            }
    
    if full_json_response:
        # Create prompt for full document correction with corrections array
        system_content = (
            "Du bist ein professioneller deutscher Sprachredakteur mit Fokus auf Grammatik- und Rechtschreibkorrekturen. "
            "Analysiere den bereitgestellten deutschen Text und identifiziere Grammatik-, Rechtschreib- oder Stilfehler. "
            "Für jeden Fehler gib den Originaltext, die korrigierte Version und eine kurze Erklärung der Änderung auf Deutsch an. "
            "Achte darauf, alle Formatierungen, einschließlich Absatzumbrüche, Überschriften und Dokumentstruktur zu erhalten. "
            "Formatiere deine Antwort als JSON-Objekt mit folgender Struktur:\n"
            "{\n"
            "  \"corrections\": [\n"
            "    {\n"
            "      \"original\": \"Originaltext mit Fehler\",\n"
            "      \"corrected\": \"Korrigierter Text\",\n"
            "      \"explanation\": \"Grund für die Korrektur\"\n"
            "    }\n"
            "  ],\n"
            "  \"corrected_full_text\": \"Der gesamte Text mit allen angewendeten Korrekturen.\"\n"
            "}\n"
            "Wenn es keine zu korrigierenden Fehler gibt, gib ein leeres corrections-Array zurück, aber schließe den Originaltext als corrected_full_text ein."
        )
    else:
        # Create optimized prompt for single paragraph correction
        system_content = (
            "Du bist ein professioneller deutscher Sprachkorrektor. Prüfe den folgenden Text auf "
            "Grammatik-, Rechtschreib- und Stilfehler. Beantworte AUSSCHLIESSLICH im folgenden JSON-Format:\n"
            "{\n"
            "  \"corrected\": \"Der korrigierte Text\",\n"
            "  \"has_corrections\": true/false,\n"
            "  \"explanation\": \"Erklärung der Änderungen (nur wenn Korrekturen vorgenommen wurden)\"\n"
            "}\n\n"
            "Wenn keine Fehler gefunden wurden, setze has_corrections auf false und explanation kann weggelassen werden. "
        )
        
        # Add element type context if provided
        if element_type:
            system_content += f"Beachte, dass es sich um ein Element vom Typ '{element_type}' handelt."
    
    messages = [
        {
            "role": "system",
            "content": system_content
        },
        {
            "role": "user",
            "content": content
        }
    ]
    
    try:
        # Call OpenAI API
        response = call_openai_api(messages, model=GPT4O_MINI_MODEL)
        
        if full_json_response:
            # Parse the full JSON result with corrections array
            correction_data = parse_json_response(response)
            if correction_data is None:
                return None
            return correction_data
        else:
            # Parse the simplified JSON format
            try:
                result = json.loads(response)
                # Ensure the response has the required fields
                if "corrected" not in result:
                    result["corrected"] = content
                if "has_corrections" not in result:
                    # Infer has_corrections by comparing content
                    result["has_corrections"] = (result["corrected"] != content)
                return result
            except json.JSONDecodeError:
                # Try to extract using regex if not valid JSON
                corrected_match = re.search(r'"corrected"\s*:\s*"(.*?)"(?=\s*[,}])', response, re.DOTALL)
                has_corrections_match = re.search(r'"has_corrections"\s*:\s*(true|false)', response, re.IGNORECASE)
                explanation_match = re.search(r'"explanation"\s*:\s*"(.*?)"(?=\s*[,}])', response, re.DOTALL)
                
                if corrected_match:
                    corrected = corrected_match.group(1).replace('\\"', '"')
                    has_corrections = (has_corrections_match and has_corrections_match.group(1).lower() == 'true')
                    explanation = explanation_match.group(1).replace('\\"', '"') if explanation_match else ""
                    
                    return {
                        "corrected": corrected,
                        "has_corrections": has_corrections,
                        "explanation": explanation
                    }
                else:
                    # Fallback: return original content
                    return {
                        "corrected": content,
                        "has_corrections": False
                    }
    except Exception as e:
        print(f"Error checking grammar: {e}")
        if full_json_response:
            return None
        else:
            return {
                "corrected": content,
                "has_corrections": False
            }

def process_structured_grammar_check(structured_content):
    """Process grammar checking paragraph by paragraph for structured content."""
    if not structured_content or not isinstance(structured_content, (list, dict)):
        print("No valid structured content to process.")
        return []
        
    print(f"Processing {len(structured_content)} document elements...")
    
    all_corrections = []
    
    # Create a map to store original and corrected content by ID
    corrections_map = {}
    
    # Create checkpoint directory and filename
    checkpoint_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'checkpoints')
    os.makedirs(checkpoint_dir, exist_ok=True)
    checkpoint_file = os.path.join(checkpoint_dir, f"checkpoint_{int(time.time())}.json")
    
    # Process each paragraph separately
    print(f"Processing {len(structured_content)} document elements...")
    
    request_count = 0
    
    for i, element in enumerate(structured_content):
        element_id = element.get('id', f'item{i}')
        content = element.get('content', '').strip()
        
        # Skip empty or very short content
        if not content or len(content) < 3:
            corrections_map[element_id] = {
                "original": content,
                "corrected": content,
                "has_corrections": False
            }
            continue
            
        print(f"Checking element {i+1}/{len(structured_content)}: {element['type']} (ID: {element_id})")
        
        # Only send non-empty paragraphs to OpenAI
        result = check_grammar_of_paragraph(content, element['type'])
        request_count += 1
        
        if result and 'corrected' in result:
            # Store the result
            corrections_map[element_id] = {
                "original": content,
                "corrected": result['corrected'],
                "has_corrections": result['has_corrections'],
                "explanation": result.get('explanation', '')
            }
            
            # Add to corrections list if there were actual changes
            if result['has_corrections']:
                all_corrections.append({
                    "id": element_id,
                    "type": element['type'],
                    "original": content,
                    "corrected": result['corrected'],
                    "explanation": result.get('explanation', '')
                })
        else:
            # If API call failed, keep original content
            corrections_map[element_id] = {
                "original": content,
                "corrected": content,
                "has_corrections": False
            }
        
        # Save checkpoint every 10 API requests
        if request_count % 10 == 0:
            # Create interim results for checkpoint
            checkpoint_data = {
                "progress": {
                    "processed_elements": i + 1,
                    "total_elements": len(structured_content),
                    "timestamp": time.time()
                },
                "corrections_so_far": all_corrections,
                "corrections_map": corrections_map
            }
            
            try:
                with open(checkpoint_file, 'w', encoding='utf-8') as f:
                    json.dump(checkpoint_data, f, indent=4, ensure_ascii=False)
                print(f"Checkpoint saved after processing {i+1}/{len(structured_content)} elements.")
            except Exception as e:
                print(f"Warning: Failed to save checkpoint: {e}")
    
    # Rebuild full text with corrections
    corrected_full_text = []
    for i, element in enumerate(structured_content):
        element_id = element.get('id', f'item{i}')
        if element_id in corrections_map:
            corrected_full_text.append(corrections_map[element_id]['corrected'])
        else:
            # Fallback to original content if not in map
            corrected_full_text.append(element.get('content', ''))
    
    return {
        "corrections": all_corrections,
        "corrected_full_text": '\n' .join(corrected_full_text)
    }

def parse_json_response(response):
    """Parse JSON response from OpenAI API with error handling."""
    # Sanitize the response before parsing as JSON
    try:
        from sanitize import sanitize_response
        response = sanitize_response(response)
    except ImportError:
        # If sanitize module is not available, use a basic cleanup
        response = response.strip()
        if response.startswith('```json'):
            response = response.split('```json', 1)[1]
        if response.endswith('```'):
            response = response.rsplit('```', 1)[0]
    
    try:
        # Try to parse the JSON directly
        data = json.loads(response)
        # Ensure corrected_full_text is present, even if empty
        if "corrected_full_text" not in data:
            data["corrected_full_text"] = data.get("text", "")
        return data
    except json.JSONDecodeError as e:
        print(f"JSON decode error: {e}")
        
        # Try to extract text within JSON blocks
        try:
            # Fix common JSON syntax issues
            fixed_response = re.sub(r'(?<=":[^{]*?)(?<!\\)"(?![,}\]])', r'\"', response)
            fixed_response = re.sub(r'(?<=":[^{]*?)(?<!\\)"(?=.*?:)', r'\"', fixed_response)
            
            # Fix potential issues with improperly terminated JSON
            if not fixed_response.strip().endswith('}'):
                fixed_response = fixed_response.rsplit('}', 1)[0] + '}'
                
            data = json.loads(fixed_response)
            if "corrected_full_text" not in data:
                data["corrected_full_text"] = data.get("text", "")
            return data
        except json.JSONDecodeError:
            pass
        
        # Extract content with regex if JSON parsing fails completely
        corrections = []
        corrected_full_text = ""
        
        # Try to extract the corrections array
        corrections_match = re.search(r'"corrections"\s*:\s*(\[.*?\])', response, re.DOTALL)
        if corrections_match:
            try:
                corrections = json.loads(corrections_match.group(1))
            except:
                pass
        
        # Try to extract the corrected text
        text_match = re.search(r'"corrected_full_text"\s*:\s*"(.*?)"(?=\s*[,}])', response, re.DOTALL)
        if text_match:
            corrected_full_text = text_match.group(1).replace('\\"', '"')
        
        # If we didn't find corrected_full_text, look for "text" field
        if not corrected_full_text:
            text_match = re.search(r'"text"\s*:\s*"(.*?)"(?=\s*[,}])', response, re.DOTALL)
            if text_match:
                corrected_full_text = text_match.group(1).replace('\\"', '"')
        
        # If we have corrections but no full text, try to extract text outside the JSON
        if not corrected_full_text and re.search(r'```json', response, re.IGNORECASE):
            # Extract text outside the JSON blocks
            no_json = re.sub(r'```json.*?```', '', response, flags=re.DOTALL|re.IGNORECASE)
            no_json = re.sub(r'[{}\[\]"\\]', '', no_json)
            if no_json.strip():
                corrected_full_text = no_json.strip()
        
        return {
            "corrections": corrections, 
            "corrected_full_text": corrected_full_text
        }

def display_corrections_table(corrections):
    """Display corrections in a readable table format in the terminal.
    
    Args:
        corrections: Correction data (dict, list, or individual correction)
    """
    # Extract corrections list from various possible formats
    corrections_list = []
    
    if isinstance(corrections, dict):
        # Handle dictionary with 'corrections' key
        if 'corrections' in corrections and isinstance(corrections['corrections'], list):
            corrections_list = corrections['corrections']
        # Handle single correction dictionary
        elif 'original' in corrections and 'corrected' in corrections:
            corrections_list = [corrections]
    elif isinstance(corrections, list):
        # Handle list of corrections
        corrections_list = corrections
    
    # Check if we have any corrections to display
    if not corrections_list or len(corrections_list) == 0:
        print("Keine Korrekturen notwendig.")
        return
    
    print("\n" + "="*80)
    print("GRAMMATIK- UND RECHTSCHREIBKORREKTUREN")
    print("="*80)
    
    valid_correction_count = 0
    for i, correction in enumerate(corrections_list, 1):
        if isinstance(correction, dict) and 'original' in correction and 'corrected' in correction:
            original = correction.get('original', 'N/A')
            corrected = correction.get('corrected', 'N/A')
            explanation = correction.get('explanation', 'Keine Erklärung verfügbar')
            
            print(f"\n{i}. ORIGINAL: {original}")
            print(f"   KORRIGIERT: {corrected}")
            print(f"   BEGRÜNDUNG: {explanation}")
            print("-"*80)
            valid_correction_count += 1
    
    if valid_correction_count == 0:
        print("\nKeine validen Korrekturen gefunden.")
    else:
        print(f"\n{valid_correction_count} Korrekturen gefunden.")
    
    # Don't display warnings about invalid formats
    # This removes the "WARNUNG: Ungültiges Korrekturformat" messages

def check_grammar_structured(text, structured_content=None, checkpoint_handler=None):
    """Check grammar using structured content analysis.
    
    Args:
        text: Full text to check
        structured_content: Optional structured content from document
        checkpoint_handler: Optional checkpoint handler for progress tracking
        
    Returns:
        List of corrections
    """
    corrections = []
    
    try:
        if structured_content:
            # Debug output to see what structured content we have
            print(f"\nProcessing {len(structured_content)} structural elements for grammar checking")
            
            # Use the analyze_document_with_openai function which has the debug output
            from .prompts import get_grammar_system_prompt, get_grammar_user_prompt
            
            # Get prompts
            system_prompt = get_grammar_system_prompt()
            user_prompt = get_grammar_user_prompt()
            
            # Call the analyze function with debugging output
            corrections = analyze_document_with_openai(
                structured_content, 
                checkpoint_handler=checkpoint_handler,
                system_prompt=system_prompt,
                user_prompt=user_prompt
            )
        else:
            # Fall back to standard approach for non-structured content
            result = _check_text_chunk(text)
            # Extract corrections list from the result dictionary
            if isinstance(result, dict) and 'corrections' in result:
                corrections = result['corrections']
            else:
                corrections = result if isinstance(result, list) else []
        
        if checkpoint_handler:
            checkpoint_handler.update_progress(1.0)
            
    except Exception as e:
        print(f"Error in grammar checking: {e}")
        import traceback
        traceback.print_exc()
        if checkpoint_handler:
            checkpoint_handler.update_progress(1.0)
    
    return corrections

def check_grammar_full_text(content: str) -> List[Dict[str, Any]]:
    """
    Check grammar and spelling in full text content.
    
    Args:
        content: The text content to check
        
    Returns:
        List of corrections
    """
    from chaiotic.ai_interface import check_grammar_with_ai
    
    # Check the full content
    print("Checking full text content...")
    corrections = check_grammar_with_ai(content)
    
    # Validate that we received a list of corrections
    if not isinstance(corrections, list):
        print(f"Warning: Expected list of corrections, got {type(corrections)}")
        if isinstance(corrections, str):
            # If we got a string, it might be an error message
            print(f"Error from API: {corrections}")
            return []
        # Convert to list if possible
        try:
            # Try to convert string to list via JSON if it's a string
            if isinstance(corrections, str):
                corrections = json.loads(corrections)
            return corrections if isinstance(corrections, list) else []
        except:
            print("Could not convert corrections to list")
            return []
    
    # Create a new list to store valid corrections
    valid_corrections = []
    
    # Validate each correction is a proper dictionary
    for correction in corrections:
        if not isinstance(correction, dict):
            print(f"Warning: Invalid correction format (not a dict): {correction}")
            continue
            
        if 'original' not in correction or 'corrected' not in correction:
            print(f"Warning: Correction missing required fields: {correction}")
            continue
            
        valid_corrections.append(correction)
    
    return valid_corrections

class CheckpointHandler:
    """Handles checkpoint saving and cleanup for grammar checking process."""
    
    def __init__(self, base_dir='/home/lenny/chaiotic/checkpoints', max_checkpoints=5, keep_last=True):
        """
        Initialize the checkpoint handler.
        
        Args:
            base_dir: Directory to store checkpoints
            max_checkpoints: Maximum number of checkpoints to keep
            keep_last: Whether to keep the last checkpoint after success
        """
        self.base_dir = base_dir
        self.max_checkpoints = max_checkpoints
        self.keep_last = keep_last
        self.checkpoints = []
        self.current_progress = 0.0
        
        # Create directory if it doesn't exist
        os.makedirs(base_dir, exist_ok=True)
        
        # Load existing checkpoints
        self._find_existing_checkpoints()

    def update_progress(self, progress_value):
        """Update the current progress and save checkpoint if needed.
        
        Args:
            progress_value: Float between 0 and 1 indicating progress
        """
        self.current_progress = max(0.0, min(1.0, float(progress_value)))
        
        # Save checkpoint at certain progress intervals (e.g., every 10%)
        progress_percent = int(self.current_progress * 100)
        if progress_percent % 10 == 0:  # Save at every 10% progress
            self.save_checkpoint(
                total_elements=100,
                processed_elements=progress_percent
            )
    
    def _find_existing_checkpoints(self):
        """Find existing checkpoint files and store them in memory."""
        self.checkpoints = []
        if os.path.exists(self.base_dir):
            for filename in os.listdir(self.base_dir):
                if filename.startswith('checkpoint_') and filename.endswith('.json'):
                    filepath = os.path.join(self.base_dir, filename)
                    try:
                        timestamp = int(filename.replace('checkpoint_', '').replace('.json', ''))
                        self.checkpoints.append((filepath, timestamp))
                    except (ValueError, TypeError):
                        # Skip files with invalid format
                        continue
                        
        # Sort by timestamp (newest first)
        self.checkpoints.sort(key=lambda x: x[1], reverse=True)
    
    def load_checkpoint(self):
        """
        Load the latest checkpoint.
        
        Returns:
            Dictionary containing checkpoint data or None if no checkpoint exists
        """
        return self.get_latest_checkpoint()

    def checkpoint_exists(self):
        """Check if there are any checkpoints available.
        
        Returns:
            True if at least one checkpoint exists, False otherwise
        """
        return len(self.checkpoints) > 0
        
    def save_checkpoint(self, corrections=None, total_elements=None, processed_elements=None):
        """
        Save a checkpoint with current progress.
        
        Args:
            corrections: Corrections found so far
            total_elements: Total number of elements to process
            processed_elements: Number of elements processed so far
        """
        # Handle different parameter formats for backward compatibility
        if corrections is None:
            corrections = []
            
        # Only save checkpoint every 10% or at least 5 elements processed
        if total_elements and processed_elements:
            checkpoint_interval = max(1, int(total_elements * 0.1))
            if processed_elements == 0 or processed_elements == total_elements or processed_elements % checkpoint_interval == 0:
                pass  # Continue with saving
            else:
                return  # Skip this checkpoint
        
        timestamp = int(time.time())
        filename = f"checkpoint_{timestamp}.json"
        filepath = os.path.join(self.base_dir, filename)
        
        # Prepare data
        corrections_map = {}
        for correction in corrections:
            if isinstance(correction, dict) and 'id' in correction:
                para_id = correction['id']
                corrections_map[para_id] = {
                    'original': correction.get('original', ''),
                    'corrected': correction.get('corrected', ''),
                    'has_corrections': True,
                    'explanation': correction.get('explanation', '')
                }
        
        # Create checkpoint data
        checkpoint_data = {
            'progress': {
                'processed_elements': processed_elements if processed_elements is not None else 0,
                'total_elements': total_elements if total_elements is not None else 0,
                'timestamp': timestamp
            },
            'corrections': corrections,
            'corrections_map': corrections_map
        }
        
        # Write checkpoint file
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(checkpoint_data, f, ensure_ascii=False, indent=4)
        
        # Add to checkpoints list
        self.checkpoints.append((filepath, timestamp))
        
        # Maintain max checkpoints
        self._clean_old_checkpoints()
        
        print(f"Saved checkpoint: {filename} ({checkpoint_data['progress']['processed_elements']}/{checkpoint_data['progress']['total_elements']} elements)")
    
    def save_corrections(self, corrections):
        """Save a checkpoint with the current corrections.
        
        Args:
            corrections: List of correction dictionaries
        """
        # Calculate progress based on number of corrections
        total_elements = 100  # Default value
        processed_elements = len(corrections)
        
        # Save checkpoint
        self.save_checkpoint(
            corrections=corrections,
            total_elements=total_elements,
            processed_elements=processed_elements
        )
    
    def load_corrections(self):
        """Load corrections from the latest checkpoint.
        
        Returns:
            List of correction dictionaries
        """
        checkpoint = self.get_latest_checkpoint()
        if checkpoint and 'corrections' in checkpoint:
            return checkpoint['corrections']
        return []
    
    def _clean_old_checkpoints(self):
        """Remove old checkpoints if we have too many."""
        if len(self.checkpoints) > self.max_checkpoints:
            # Sort by timestamp (newest first)
            self.checkpoints.sort(key=lambda x: x[1], reverse=True)
            
            # Delete older checkpoints
            for filepath, _ in self.checkpoints[self.max_checkpoints:]:
                try:
                    os.remove(filepath)
                    print(f"Removed old checkpoint: {os.path.basename(filepath)}")
                except OSError:
                    # Failed to delete
                    pass
            
            # Update list
            self.checkpoints = self.checkpoints[:self.max_checkpoints]
    
    def clean_up_on_success(self):
        """Clean up all checkpoints after successful completion."""
        if not self.keep_last:
            # Remove all checkpoints
            for filepath, _ in self.checkpoints:
                try:
                    os.remove(filepath)
                    print(f"Cleaned up checkpoint: {os.path.basename(filepath)}")
                except OSError:
                    pass
            self.checkpoints = []
        else:
            # Keep only the latest checkpoint
            if len(self.checkpoints) > 0:
                latest = self.checkpoints[0]
                for filepath, _ in self.checkpoints[1:]:
                    try:
                        os.remove(filepath)
                        print(f"Removed intermediate checkpoint: {os.path.basename(filepath)}")
                    except OSError:
                        pass
                self.checkpoints = [latest]
    
    def get_latest_checkpoint(self):
        """Get the latest checkpoint if any exists."""
        if not self.checkpoints:
            return None
        
        filepath, _ = self.checkpoints[0]
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return None
    
    def purge_all_checkpoints(self):
        """Purge all checkpoints from the directory."""
        for filepath, _ in self.checkpoints:
            try:
                os.remove(filepath)
                print(f"Purged checkpoint: {os.path.basename(filepath)}")
            except OSError:
                pass
        
        self.checkpoints = []

def parse_corrections(response, original_text):
    """Parse corrections from API response.
    
    Args:
        response: JSON response from OpenAI API
        original_text: Original text that was checked
        
    Returns:
        Dict with corrections list and corrected_full_text
    """
    try:
        # Parse and validate response
        if isinstance(response, str):
            try:
                parsed = json.loads(response)
            except json.JSONDecodeError:
                print("Invalid JSON response received")
                return {"corrections": [], "corrected_full_text": original_text}
        else:
            parsed = response
        
        # Ensure we have a proper dictionary
        if not isinstance(parsed, dict):
            print(f"Expected dict response, got {type(parsed)}")
            return {"corrections": [], "corrected_full_text": original_text}
        
        # Extract and validate corrections
        corrections = parsed.get('corrections', [])
        if not isinstance(corrections, list):
            print(f"Expected corrections to be list, got {type(corrections)}")
            corrections = []
        
        # Validate each correction
        valid_corrections = []
        for i, corr in enumerate(corrections):
            if not isinstance(corr, dict):
                print(f"Invalid correction {i+1}: not a dictionary")
                continue
                
            if 'original' not in corr or 'corrected' not in corr:
                print(f"Invalid correction {i+1}: missing required fields")
                continue
            
            # Ensure all fields are strings
            corr['original'] = str(corr['original'])
            corr['corrected'] = str(corr['corrected'])
            if 'explanation' in corr:
                corr['explanation'] = str(corr['explanation'])
                
            # Only include corrections where original text exists in document
            if corr['original'] in original_text:
                valid_corrections.append(corr)
            else:
                print(f"Warning: Original text '{corr['original']}' not found in document")
        
        # Get or generate corrected full text
        corrected_full_text = parsed.get('corrected_full_text', original_text)
        if not corrected_full_text or not isinstance(corrected_full_text, str):
            corrected_full_text = original_text
            # Apply corrections to generate full text
            for corr in valid_corrections:
                corrected_full_text = corrected_full_text.replace(
                    corr['original'], 
                    corr['corrected']
                )
        
        return {
            "corrections": valid_corrections,
            "corrected_full_text": corrected_full_text
        }
        
    except Exception as e:
        print(f"Error parsing corrections: {str(e)}")
        return {"corrections": [], "corrected_full_text": original_text}

def split_into_chunks(text, chunk_size):
    """Split text into chunks of approximately equal size.
    
    Args:
        text: Text to split
        chunk_size: Approximate size of each chunk in characters
        
    Returns:
        List of text chunks
    """
    # If text is smaller than chunk_size, return it as is
    if len(text) <= chunk_size:
        return [text]
    
    # Split text into paragraphs
    paragraphs = text.split('\n\n')
    
    chunks = []
    current_chunk = []
    current_size = 0
    
    for paragraph in paragraphs:
        # Skip empty paragraphs
        if not paragraph.strip():
            continue
            
        paragraph_size = len(paragraph)
        
        # If adding this paragraph would exceed chunk_size, start a new chunk
        if current_size + paragraph_size > chunk_size and current_chunk:
            chunks.append('\n\n'.join(current_chunk))
            current_chunk = [paragraph]
            current_size = paragraph_size
        else:
            current_chunk.append(paragraph)
            current_size += paragraph_size
    
    # Add the last chunk if it's not empty
    if current_chunk:
        chunks.append('\n\n'.join(current_chunk))
    
    return chunks

def _check_text_chunk(text):
    """Check grammar of a single text chunk.
    
    Args:
        text: Text chunk to check
        
    Returns:
        Dict containing corrections and corrected_full_text
    """
    from .prompts import get_grammar_prompt
    
    # Skip empty chunks
    if not text or not text.strip():
        return {
            'corrections': [],
            'corrected_full_text': text
        }
    
    try:
        # Create prompt for this chunk
        prompt = get_grammar_prompt(text)
        
        # Get corrections from AI
        response = call_openai_api(prompt)
        
        # Parse corrections from response
        parsed = parse_corrections(response, text)
        if not isinstance(parsed, dict):
            # Convert legacy format to dict
            return {
                'corrections': parsed if isinstance(parsed, list) else [],
                'corrected_full_text': text
            }
        return parsed
        
    except Exception as e:
        print(f"Error checking text chunk: {e}")
        return {
            'corrections': [],
            'corrected_full_text': text
        }

def analyze_document_with_openai(structured_content, checkpoint_handler=None, system_prompt="", user_prompt=""):
    """Analyze the document with OpenAI API for grammar and spelling corrections.
    
    Args:
        structured_content: Structured content from the document
        checkpoint_handler: Optional checkpoint handler to resume from previous state
        system_prompt: Optional system prompt to use
        user_prompt: Optional user prompt to use
        
    Returns:
        List of correction dictionaries
    """
    client = init_openai_client()
    
    if not client:
        return []
    
    if not structured_content:
        print("No content to analyze")
        return []
    
    # Optionally skip content analysis if checkpoint already exists
    if checkpoint_handler and checkpoint_handler.checkpoint_exists():
        print("Found checkpoint, resuming from saved state")
        return checkpoint_handler.load_corrections()
    
    corrections = []
    
    # Get language model from config
    config = load_config()
    model = config.get("ai", {}).get("model", DEFAULT_MODEL)
    
    # Define batching parameters
    max_tokens = 4000  # Max tokens per batch to avoid token limits
    current_batch = []
    current_token_count = 0
    batch_count = 0
    
    # Debug information for API requests
    print("\n===== OPENAI API REQUEST DEBUG =====")
    print(f"Using model: {model}")
    print(f"Total content items: {len(structured_content)}")
    
    for item in structured_content:
        # Estimate token count (roughly 4 chars per token)
        item_content = item.get('content', '')
        item_tokens = len(item_content) // 4 + 1
        
        # Debug info about sending content to API
        print(f"\nItem ID: {item.get('id')}, Type: {item.get('type')}, Estimated tokens: {item_tokens}")
        print(f"Content preview: {item_content[:50]}{'...' if len(item_content) > 50 else ''}")
        
        # Check if we have XML content to include
        xml_content = item.get('xml_content', '')
        if xml_content:
            xml_preview = xml_content[:100].replace('\n', ' ')
            print(f"Including XML structure: {xml_preview}{'...' if len(xml_content) > 100 else ''}")
            # Include token estimate for XML
            xml_tokens = len(xml_content) // 4 + 1
            print(f"XML estimated tokens: {xml_tokens}")
            item_tokens += xml_tokens
        
        # If adding this item would exceed the token limit, process the current batch
        if current_token_count + item_tokens > max_tokens and current_batch:
            batch_count += 1
            print(f"\nProcessing batch {batch_count} with {len(current_batch)} items and ~{current_token_count} tokens")
            batch_corrections = process_content_batch(current_batch, client, model, system_prompt, user_prompt)
            corrections.extend(batch_corrections)
            
            # Save checkpoint after each batch
            if checkpoint_handler:
                checkpoint_handler.save_corrections(corrections)
                print(f"Saved checkpoint after batch {batch_count}")
            
            # Reset for next batch
            current_batch = []
            current_token_count = 0
        
        # Add the item to the current batch
        current_batch.append(item)
        current_token_count += item_tokens
    
    # Process any remaining items
    if current_batch:
        batch_count += 1
        print(f"\nProcessing final batch {batch_count} with {len(current_batch)} items and ~{current_token_count} tokens")
        batch_corrections = process_content_batch(current_batch, client, model, system_prompt, user_prompt)
        corrections.extend(batch_corrections)
    
    print("=" * 60)
    
    # Save final checkpoint
    if checkpoint_handler:
        checkpoint_handler.save_corrections(corrections)
        print("Saved final checkpoint")
    
    return corrections