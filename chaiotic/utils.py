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