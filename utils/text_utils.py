"""Text utility functions used across multiple modules."""

import re
from typing import List, Dict, Any, Optional
from difflib import SequenceMatcher

class FuzzyMatcher:
    """Class for finding text with fuzzy matching to handle minor differences."""
    
    def __init__(self, threshold=0.8):
        """Initialize the matcher with a similarity threshold."""
        self.threshold = threshold
    
    def find_best_match(self, needle: str, haystack: str) -> tuple:
        """Find the best match position for needle in haystack.
        
        Args:
            needle: Text to find
            haystack: Text to search in
            
        Returns:
            Tuple of (start_pos, end_pos, confidence)
        """
        if not needle or not haystack:
            return (-1, -1, 0.0)
            
        # For very short needles, use direct text search
        if len(needle) < 10:
            pos = haystack.find(needle)
            if pos >= 0:
                return (pos, pos + len(needle), 1.0)
        
        # For longer text, use sequence matcher
        matcher = SequenceMatcher(None, haystack, needle)
        match = matcher.find_longest_match(0, len(haystack), 0, len(needle))
        
        # Calculate match quality
        if match.size == 0:
            return (-1, -1, 0.0)
            
        confidence = match.size / len(needle)
        
        # If confidence is below threshold, try word-by-word matching
        if confidence < self.threshold and len(needle.split()) > 1:
            words = needle.split()
            # Try to find first and last words
            first_word = words[0]
            last_word = words[-1]
            
            first_pos = haystack.find(first_word)
            last_pos = haystack.find(last_word)
            
            if first_pos >= 0 and last_pos >= 0 and first_pos < last_pos:
                # We found both words in correct order
                start_pos = first_pos
                end_pos = last_pos + len(last_word)
                
                # Calculate new confidence based on ratio of found text to needle
                found_text = haystack[start_pos:end_pos]
                confidence = SequenceMatcher(None, found_text, needle).ratio()
                
                if confidence >= self.threshold:
                    return (start_pos, end_pos, confidence)
        
        # If we have a good enough match, return positions
        if confidence >= self.threshold:
            start_pos = match.b
            end_pos = match.b + match.size
            return (start_pos, end_pos, confidence)
            
        return (-1, -1, 0.0)
    
    def replace_with_context(self, original: str, search_text: str, replacement: str) -> str:
        """Replace text with fuzzy matching, considering context.
        
        Args:
            original: Original full text
            search_text: Text to find and replace
            replacement: Replacement text
            
        Returns:
            Text with replacement applied
        """
        if not search_text or search_text == replacement:
            return original
            
        # Find position with fuzzy matching
        start_pos, end_pos, confidence = self.find_best_match(search_text, original)
        
        if start_pos >= 0 and confidence >= self.threshold:
            # We found a good match, replace it
            result = original[:start_pos] + replacement + original[end_pos:]
            return result
            
        # If no good match, return original
        return original

def generate_full_text_from_corrections(corrections: List[Dict[str, Any]], 
                                      original_text: Optional[str] = None) -> str:
    """Generate full corrected text by applying corrections to original text.
    
    Args:
        corrections: List of correction dictionaries
        original_text: Original text to apply corrections to (if available)
        
    Returns:
        Full text with all corrections applied
    """
    if not corrections:
        return original_text or ""
        
    # If we have original text, apply corrections to it
    if original_text:
        result = original_text
        matcher = FuzzyMatcher(threshold=0.7)  # Lower threshold for better matches
        
        # Sort corrections by position in original text (if available)
        # This prevents issues when applying multiple overlapping corrections
        sorted_corrections = []
        for corr in corrections:
            if not isinstance(corr, dict):
                continue
                
            original = corr.get('original', '')
            if not original:
                continue
                
            # Find position in text
            pos = original_text.find(original)
            sorted_corrections.append((pos, corr))
            
        # Sort by position, with -1 (not found) at the end
        sorted_corrections.sort(key=lambda x: float('inf') if x[0] == -1 else x[0])
        
        # Apply corrections in order
        for _, corr in sorted_corrections:
            original = corr.get('original', '')
            corrected = corr.get('corrected', '')
            
            if original and corrected and original != corrected:
                result = matcher.replace_with_context(result, original, corrected)
                
        return result
    
    # Without original text, try to reconstruct from corrections
    # This is less reliable but can work for simple cases
    full_text = ""
    last_correction = None
    
    for corr in corrections:
        if not isinstance(corr, dict):
            continue
            
        # If this correction has full_text or corrected_full_text, use it
        if 'corrected_full_text' in corr:
            return corr['corrected_full_text']
        if 'full_text' in corr:
            return corr['full_text']
            
        # Otherwise, add the corrected text
        corrected = corr.get('corrected', '')
        if corrected:
            # Add separator if needed
            if last_correction and not (last_correction.endswith('\n') or corrected.startswith('\n')):
                full_text += '\n\n'
            full_text += corrected
            last_correction = corrected
    
    return full_text

def sanitize_response(response):
    """Clean up JSON responses from API to make them valid JSON.
    
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
    
    return response

def split_text_into_chunks(text, max_tokens_per_chunk=8000, overlap_tokens=150, model=None):
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
        
    # Rough estimate of tokens per character (for European languages)
    tokens_per_char = 0.25
    
    # Get character limits from token limits
    max_chars = int(max_tokens_per_chunk / tokens_per_char)
    overlap_chars = int(overlap_tokens / tokens_per_char)
    
    # If text is small enough, return as is
    if len(text) <= max_chars:
        return [text]
    
    # Split text into sentences using regex
    sentences = re.split(r'(?<=[.!?])\s+', text)
    
    chunks = []
    current_chunk = []
    current_length = 0
    
    for sentence in sentences:
        sentence_length = len(sentence)
        
        # If a single sentence exceeds the max length, split it further
        if sentence_length > max_chars:
            # First add the current chunk if not empty
            if current_chunk:
                chunks.append(''.join(current_chunk))
                current_chunk = []
                current_length = 0
            
            # Split the long sentence by paragraphs or character count
            sentence_parts = sentence.split('\n')
            if max(len(part) for part in sentence_parts) > max_chars:
                # Some paragraph is still too long, do character-based splitting
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
                # Process paragraph by paragraph
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
                
                # Start new chunk with overlap from previous chunk
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