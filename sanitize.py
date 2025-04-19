"""
Sanitize module - Provides utilities for sanitizing responses from LLMs like OpenAI
to ensure they're properly formatted JSON before parsing.
"""

import re

def sanitize_response(response_str):
    """
    Clean up a response string to extract valid JSON.
    
    Args:
        response_str (str): The raw response string from an LLM
        
    Returns:
        str: A sanitized JSON string
    """
    # If it's not a string, return it as is
    if not isinstance(response_str, str):
        return response_str
        
    # Extract JSON from markdown code blocks if present
    json_pattern = r'```(?:json)?\s*([\s\S]*?)\s*```'
    matches = re.findall(json_pattern, response_str)
    if matches:
        response_str = matches[0]
    
    # Remove any leading/trailing whitespace
    response_str = response_str.strip()
    
    # Ensure the response is wrapped in curly braces if it looks like JSON
    if not (response_str.startswith('{') and response_str.endswith('}')):
        # If it doesn't look like JSON at all, wrap it in a simple structure
        if not ('{' in response_str and '}' in response_str):
            return '{"text": ' + json.dumps(response_str) + '}'
    
    # Handle escaped quotes and other common issues
    response_str = response_str.replace('\\"', '"')
    
    # Find and fix unquoted keys (a common LLM mistake)
    def quote_keys(match):
        return f'"{match.group(1)}":'
    
    response_str = re.sub(r'([a-zA-Z_][a-zA-Z0-9_]*)\s*:', quote_keys, response_str)
    
    return response_str

# Add import inside the module to avoid circular imports in main.py
import json