"""Utility functions for document processing."""

import re
import os
import logging
import json

# Try to import optional dependencies
NLTK_AVAILABLE = False
try:
    import nltk
    NLTK_AVAILABLE = True
    
    # Try to download NLTK data if needed
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        try:
            nltk.download('punkt', quiet=True)
        except:
            print("Warning: Could not download NLTK data for text splitting")
except ImportError:
    print("Warning: NLTK not installed. Using fallback text splitting method.")
    # NLTK is not required, we'll use regex-based splitting as fallback

def split_text_into_chunks(text, max_tokens_per_chunk=80000, overlap_tokens=100, model=None):
    """Split text into chunks while respecting sentence boundaries.
    
    Args:
        text: The text to split
        max_tokens_per_chunk: Maximum tokens per chunk
        overlap_tokens: Number of tokens to overlap between chunks
        model: Model name to use for token estimation
        
    Returns:
        List of text chunks
    """
    if not text:
        return []
        
    # Rough estimate of tokens per character for German language
    # More precise would be to use the tokenizer, but this works well enough for chunking
    tokens_per_char = 0.25  # Approximate for most European languages
    
    # Get character limits from token limits
    max_chars = int(max_tokens_per_chunk / tokens_per_char)
    overlap_chars = int(overlap_tokens / tokens_per_char)
    
    # If text is small enough, return as is
    if len(text) <= max_chars:
        return [text]
    
    # Split text into sentences
    if NLTK_AVAILABLE:
        try:
            sentences = nltk.sent_tokenize(text)
        except Exception as e:
            print(f"Error tokenizing text with NLTK: {e}")
            # Fallback to regex-based sentence splitting
            sentences = re.split(r'(?<=[.!?])\s+', text)
    else:
        # Fallback method using regex for sentence splitting when NLTK is not available
        sentences = re.split(r'(?<=[.!?])\s+', text)
    
    chunks = []
    current_chunk = []
    current_length = 0
    
    for sentence in sentences:
        sentence_length = len(sentence)
        
        # If a single sentence exceeds the max length, we need to split it further
        if sentence_length > max_chars:
            if current_chunk:
                chunks.append(''.join(current_chunk))
                current_chunk = []
                current_length = 0
            
            # Split the long sentence by paragraphs or just by character count if needed
            sentence_parts = sentence.split('\n')
            if max(len(part) for part in sentence_parts) > max_chars:
                # Some paragraph is still too long, do crude character-based splitting
                start = 0
                while start < len(sentence):
                    end = start + max_chars
                    if end >= len(sentence):
                        chunks.append(sentence[start:])
                    else:
                        # Try to end at a word boundary
                        while end > start and not sentence[end].isspace():
                            end -= 1
                        if end == start:  # If no space found, just cut at max_chars
                            end = start + max_chars
                        chunks.append(sentence[start:end])
                    start = end
            else:
                current_part = []
                part_length = 0
                for part in sentence_parts:
                    part_size = len(part) + 1  # +1 for the newline
                    if part_length + part_size > max_chars:
                        chunks.append('\n'.join(current_part))
                        current_part = [part]
                        part_length = part_size
                    else:
                        current_part.append(part)
                        part_length += part_size
                if current_part:
                    chunks.append('\n'.join(current_part))
        else:
            # If adding this sentence would exceed the limit, save chunk and start a new one
            if current_length + sentence_length > max_chars:
                chunks.append(''.join(current_chunk))
                
                # Start new chunk with some overlap from the end of the previous chunk
                overlap_start = max(0, len(''.join(current_chunk)) - overlap_chars)
                overlap_text = ''.join(current_chunk)[overlap_start:]
                
                current_chunk = [overlap_text, sentence]
                current_length = len(overlap_text) + sentence_length
            else:
                current_chunk.append(sentence)
                current_length += sentence_length
    
    # Add the last chunk if not empty
    if current_chunk:
        chunks.append(''.join(current_chunk))
    
    return chunks

def preprocess_content(content):
    """Preprocess document content for grammar checking.
    
    This function performs basic cleaning and normalization
    of the text content before sending it for grammar checking.
    
    Args:
        content: The text content to preprocess
        
    Returns:
        Preprocessed text content
    """
    if not content:
        return content
        
    # Replace multiple newlines with a maximum of two
    content = re.sub(r'\n{3,}', '\n\n', content)
    
    # Strip whitespace from each line
    lines = [line.strip() for line in content.split('\n')]
    content = '\n'.join(lines)
    
    # Remove multiple spaces
    content = re.sub(r' {2,}', ' ', content)
    
    # Ensure proper spacing after punctuation
    content = re.sub(r'([.!?])([A-ZÄÖÜ])', r'\1 \2', content)
    
    return content

def sanitize_response(response):
    """Clean up JSON responses from OpenAI to make them valid JSON.
    
    Args:
        response: The API response to sanitize
        
    Returns:
        Sanitized response string
    """
    # First handle markdown code blocks
    if '```' in response:
        # Extract content between triple backticks
        match = re.search(r'```(?:json)?(.*?)```', response, re.DOTALL)
        if match:
            response = match.group(1).strip()
    
    # Basic cleanup
    response = response.strip()
    
    # If response starts or ends with JSON delimiters, clean it up
    if response.startswith('{') and not response.endswith('}'):
        last_brace = response.rfind('}')
        if last_brace > 0:
            response = response[:last_brace+1]
    
    if response.startswith('[') and not response.endswith(']'):
        last_bracket = response.rfind(']')
        if last_bracket > 0:
            response = response[:last_bracket+1]
    
    # Attempt to fix common JSON syntax errors
    # Fix unescaped quotes in strings
    response = re.sub(r'(?<=":[^{[]*?)(?<!\\)"(?=[^"]*?")', r'\"', response)
    
    # Fix trailing commas in objects
    response = re.sub(r',\s*}', '}', response)
    
    # Fix trailing commas in arrays
    response = re.sub(r',\s*]', ']', response)
    
    # Validate JSON by trying to parse it
    try:
        json.loads(response)
    except json.JSONDecodeError as e:
        # If still not valid, try to narrow down the valid JSON portion
        if response.startswith('{'):
            # For objects, try to find the last valid closing brace
            stack = []
            valid_end = -1
            for i, char in enumerate(response):
                if char == '{':
                    stack.append(i)
                elif char == '}' and stack:
                    stack.pop()
                    if not stack:  # If all braces are closed
                        valid_end = i
            
            if valid_end >= 0:
                response = response[:valid_end+1]
    
    return response

def load_cached_response(request_hash, cache_dir=None):
    """Load a cached API response from disk.
    
    Args:
        request_hash: Hash of the request to load
        cache_dir: Directory containing cached responses
        
    Returns:
        Cached response or None if not found/invalid
    """
    if not cache_dir:
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'api_requests')
    
    cache_file = os.path.join(cache_dir, f"{request_hash}.json")
    
    if not os.path.exists(cache_file):
        return None
        
    try:
        with open(cache_file, 'r', encoding='utf-8') as f:
            cached_response = json.load(f)
        return cached_response
    except Exception as e:
        print(f"Error loading cached response: {e}")
        return None

def save_cached_response(request_hash, response, cache_dir=None):
    """Save an API response to the cache.
    
    Args:
        request_hash: Hash of the request to save
        response: Response to cache
        cache_dir: Directory to save cached responses
        
    Returns:
        True if saved successfully, False otherwise
    """
    if not cache_dir:
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'api_requests')
    
    os.makedirs(cache_dir, exist_ok=True)
    cache_file = os.path.join(cache_dir, f"{request_hash}.json")
    
    try:
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(response, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving cached response: {e}")
        return False

def format_duration(seconds):
    """Format a duration in seconds to a readable string.
    
    Args:
        seconds: Duration in seconds
        
    Returns:
        Formatted string like "2m 30s" or "1h 15m 30s"
    """
    minutes, seconds = divmod(int(seconds), 60)
    hours, minutes = divmod(minutes, 60)
    
    parts = []
    if hours > 0:
        parts.append(f"{hours}h")
    if minutes > 0 or hours > 0:
        parts.append(f"{minutes}m")
    parts.append(f"{seconds}s")
    
    return " ".join(parts)

def save_document(file_path, corrections, original_doc=None, is_docx=None):
    """Save corrected content to document files.
    
    Args:
        file_path: Original document file path.
        corrections: Correction data from the grammar checker.
        original_doc: Original document object to preserve formatting.
        is_docx: Boolean indicating if the file is a DOCX (True) or ODT (False).
        
    Returns:
        Tuple of (json_file_path, text_file_path, doc_file_path)
    """
    from datetime import datetime
    import json
    import os

    # Get the directory, filename, and extension from the original path
    dir_name = os.path.dirname(file_path) or '.'
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1]
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Paths for output files
    json_file = os.path.join(dir_name, f"{base_name}_corrected_{timestamp}.json")
    text_file = os.path.join(dir_name, f"{base_name}_corrected_{timestamp}.txt")
    doc_file = os.path.join(dir_name, f"{base_name}_corrected_{timestamp}{ext}")
    
    try:
        # Save JSON file with all corrections
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(corrections, f, indent=4, ensure_ascii=False)
        
        # Extract corrected text
        if 'corrected_full_text' in corrections:
            # Save text file with corrected text
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(corrections['corrected_full_text'])
            
            # Save corrected document with original formatting
            save_document_content(corrections['corrected_full_text'], doc_file, original_doc, is_docx)
        else:
            print("No corrected text found in corrections.")
            text_file = None
            doc_file = None
            
        return json_file, text_file, doc_file
    except Exception as e:
        print(f"Error saving document: {e}")
        return None, None, None

def save_document_content(content, file_path, original_doc=None, is_docx=None):
    """Save content to a document file with appropriate formatting.
    
    Args:
        content: Text content to write.
        file_path: Path where to save the document.
        original_doc: Original document object to preserve formatting.
        is_docx: Boolean indicating if the file is a DOCX (True) or ODT (False).
        
    Returns:
        Tuple of (json_path, text_path, doc_path) or None on error.
    """
    try:
        # Import here to avoid circular imports
        from .document_handler import save_correction_outputs
        
        # Create a correction entry for the full text
        corrections = [{
            'original': '',  # Empty since this is a full text replacement
            'corrected': content,
            'type': 'full_text'
        }]
        
        return save_correction_outputs(file_path, corrections, original_doc, is_docx)
    except Exception as e:
        print(f"Error saving document content: {e}")
        return None, None, None