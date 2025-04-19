"""Grammar checking module for documents."""

import json
import re
import time
import os
import hashlib
import difflib
import shutil
from typing import List, Dict, Any, Optional, Union

from .config import DISABLE_CACHE, GPT4O_MINI_MODEL

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

def check_grammar(content, structured_content=None, use_structured=False, model=None, force_refresh=False, use_mock_data=False, cache_key=None, checkpoint_handler=None):
    """
    Check grammar and spelling in the content.
    
    Args:
        content: The text content to check
        structured_content: Optional structured content for paragraph-by-paragraph processing
        use_structured: Whether to use structured processing
        model: Optional model to use
        force_refresh: Whether to force refresh cached results
        use_mock_data: Whether to use mock data
        cache_key: Optional cache key
        checkpoint_handler: Optional checkpoint handler for managing checkpoints
        
    Returns:
        List of corrections
    """
    from chaiotic.config import load_config
    config = load_config()
    
    # Check if we have mock data for testing
    mock_enabled = use_mock_data
    if not mock_enabled:
        try:
            mock_enabled = config.get('use_mock_data', False)
        except AttributeError:
            # Handle the case where Config object doesn't have get method
            mock_enabled = hasattr(config, 'use_mock_data') and config.use_mock_data
    
    if mock_enabled:
        print("Using mock data for grammar checking")
        return get_mock_corrections()
    
    if use_structured and structured_content:
        print("Using structured content for grammar checking")
        return check_grammar_structured(structured_content, checkpoint_handler)
    else:
        print("Using full text for grammar checking")
        return check_grammar_full_text(content)

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

def call_openai_api(messages, model=GPT4O_MINI_MODEL, retries=3, delay=5):
    """Call the OpenAI API with caching."""
    from time import sleep
    
    # Create a hash based on the full message content and model
    # This ensures we properly cache based on the document content
    message_content = json.dumps(messages, sort_keys=True, ensure_ascii=False)
    request_hash = hashlib.md5((message_content + model).encode('utf-8')).hexdigest()
    
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
            # Continue to make a new API call if cache read fails
    
    # Define appropriate max_tokens based on model
    if model == GPT4O_MINI_MODEL:
        max_tokens = 4096  # Reasonable limit for paragraph-by-paragraph checks
    else:
        max_tokens = 4096  # Default for other models
    
    # Make API call if no cache or cache reading failed
    for attempt in range(retries):
        try:
            completion = client.chat.completions.create(
                model=model,
                messages=messages,
                stream=False,
                temperature=0,
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
                raise

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
    """Display corrections in a readable table format in the terminal."""
    # Handle different formats of corrections input
    if isinstance(corrections, list):
        # If corrections is already a list, use it directly
        corrections_list = corrections
    elif isinstance(corrections, dict) and 'corrections' in corrections:
        # If corrections is a dict with a 'corrections' key, use that
        corrections_list = corrections.get('corrections', [])
    else:
        # Default to empty list if unknown format
        corrections_list = []
    
    # Check if we have any corrections to display
    if not corrections_list or len(corrections_list) == 0:
        print("Keine Korrekturen notwendig.")
        return
    
    print("\n" + "="*80)
    print("GRAMMATIK- UND RECHTSCHREIBKORREKTUREN")
    print("="*80)
    
    for i, correction in enumerate(corrections_list, 1):
        if isinstance(correction, dict):
            original = correction.get('original', 'N/A')
            corrected = correction.get('corrected', 'N/A')
            explanation = correction.get('explanation', 'Keine Erklärung verfügbar')
            print(f"\n{i}. ORIGINAL: {original}")
            print(f"   KORRIGIERT: {corrected}")
            print(f"   BEGRÜNDUNG: {explanation}")
            print("-"*80)
        else:
            print(f"\n{i}. WARNUNG: Ungültiges Korrekturformat: {correction}")
            print("-"*80)

def check_grammar_structured(structured_content, checkpoint_handler=None) -> List[Dict[str, Any]]:
    """
    Check grammar and spelling in structured content (paragraph by paragraph).
    
    Args:
        structured_content: The structured content to check
        checkpoint_handler: Optional checkpoint handler
        
    Returns:
        List of corrections
    """
    if not structured_content or not isinstance(structured_content, (list, dict)):
        print("No valid structured content to process.")
        return []
        
    print(f"Processing {len(structured_content)} document elements...")
    
    all_corrections = []
    
    # Create a map to store original and corrected content by ID
    corrections_map = {}
    
    # Process each paragraph separately
    print(f"Processing {len(structured_content)} document elements...")
    
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
        
        # Optional checkpoint saving - only if handler provided
        if checkpoint_handler:
            checkpoint_handler.save_checkpoint(total_elements=len(structured_content), 
                                              processed_elements=i+1, 
                                              corrections=all_corrections)
    
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
        
        # Create directory if it doesn't exist
        os.makedirs(base_dir, exist_ok=True)
        
        # Load existing checkpoints
        self._find_existing_checkpoints()
    
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
    
    def save_checkpoint(self, total_elements, processed_elements, corrections):
        """
        Save a checkpoint with current progress.
        
        Args:
            total_elements: Total number of elements to process
            processed_elements: Number of elements processed so far
            corrections: Corrections found so far
        """
        # Only save checkpoint every 10% or at least 5 elements processed
        checkpoint_interval = max(1, int(total_elements * 0.1))
        if processed_elements == 0 or processed_elements == total_elements or processed_elements % checkpoint_interval == 0:
            timestamp = int(time.time())
            filename = f"checkpoint_{timestamp}.json"
            filepath = os.path.join(self.base_dir, filename)
            
            # Prepare data
            corrections_map = {}
            for correction in corrections:
                if 'id' in correction:
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
                    'processed_elements': processed_elements,
                    'total_elements': total_elements,
                    'timestamp': timestamp
                },
                'corrections_so_far': corrections,
                'corrections_map': corrections_map
            }
            
            # Write checkpoint file
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(checkpoint_data, f, ensure_ascii=False, indent=4)
            
            # Add to checkpoints list
            self.checkpoints.append((filepath, timestamp))
            
            # Maintain max checkpoints
            self._clean_old_checkpoints()
            
            print(f"Saved checkpoint: {filename} ({processed_elements}/{total_elements} elements)")
    
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