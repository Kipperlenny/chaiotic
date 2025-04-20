"""Language processing module for grammar checking and text analysis."""

import re
import json
import os
import logging
from datetime import datetime

# Import text utilities
from .text_utils import (
    sanitize_response,
    split_text_into_chunks,
    generate_full_text_from_corrections
)

# Try to import OpenAI
try:
    import openai
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# Setup logging
logger = logging.getLogger(__name__)

class GrammarProcessor:
    """Main class for processing grammar issues in text."""
    
    def __init__(self, config):
        """Initialize grammar processor with configuration.
        
        Args:
            config: Configuration dictionary with API keys and settings
        """
        self.config = config
        self.openai_client = None
        
        self._initialize_openai()
    
    def _initialize_openai(self):
        """Initialize the OpenAI client."""
        # Initialize OpenAI if available
        if OPENAI_AVAILABLE and self.config.get('openai_api_key'):
            try:
                openai.api_key = self.config['openai_api_key']
                self.openai_client = OpenAI(api_key=self.config['openai_api_key'])
                logger.info("OpenAI client initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize OpenAI client: {e}")
                self.openai_client = None
    
    def check_grammar(self, text, tool='openai'):
        """Check grammar issues in the provided text.
        
        Args:
            text: The text to check
            tool: The tool to use (only 'openai' supported)
            
        Returns:
            Dictionary with correction information
        """
        if not text.strip():
            return {"error": "Empty text provided"}
        
        if tool != 'openai' or not self.openai_client:
            return {"error": "Only OpenAI grammar checking is supported"}
        
        return self._check_with_openai(text)
    
    def _check_with_openai(self, text):
        """Check grammar using OpenAI.
        
        Args:
            text: The text to check
            
        Returns:
            Dictionary with correction information
        """
        try:
            # Split text into chunks if it's large
            if len(text) > 10000:
                chunks = split_text_into_chunks(text, max_tokens_per_chunk=8000)
                all_corrections = self._process_text_chunks(chunks)
                return all_corrections
            
            # Prepare the system message
            system_message = """You are a professional editor and grammar checker. 
            Analyze the text for grammar, spelling, punctuation, and style issues. 
            Provide corrections and explanations for each issue found.
            
            Output must be in the following JSON format:
            {
                "corrections": [
                    {
                        "original": "text with error",
                        "corrected": "corrected text",
                        "explanation": "explanation of the correction",
                        "type": "grammar|spelling|punctuation|style"
                    }
                ],
                "corrected_full_text": "Full corrected text with all issues fixed"
            }
            """
            
            # Make API call
            model = self.config.get('openai_model', 'gpt-4o-mini')
            response = self.openai_client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": text}
                ],
                temperature=0.3,
                response_format={"type": "json_object"}
            )
            
            # Parse response
            response_text = response.choices[0].message.content
            result = json.loads(sanitize_response(response_text))
            
            # Ensure expected format
            if "corrections" not in result:
                result["corrections"] = []
                
            if "corrected_full_text" not in result:
                result["corrected_full_text"] = text
                
            return result
            
        except Exception as e:
            logger.error(f"Error checking grammar with OpenAI: {e}")
            return {
                "original_text": text,
                "error": f"OpenAI API error: {str(e)}",
                "corrections": [],
                "corrected_full_text": text
            }
    
    def _process_text_chunks(self, chunks):
        """Process large text by checking grammar in chunks.
        
        Args:
            chunks: List of text chunks to process
            
        Returns:
            Combined results from all chunks
        """
        all_corrections = []
        all_text = ""
        
        for i, chunk in enumerate(chunks):
            logger.info(f"Processing chunk {i+1}/{len(chunks)}")
            
            result = self._check_with_openai(chunk)
            
            if "corrections" in result:
                all_corrections.extend(result["corrections"])
            
            if "corrected_full_text" in result:
                all_text += result["corrected_full_text"] + " "
        
        return {
            "corrections": all_corrections,
            "corrected_full_text": all_text.strip()
        }
    
    def summarize_text(self, text, max_length=200):
        """Generate a summary of the provided text.
        
        Args:
            text: The text to summarize
            max_length: Maximum length of the summary
            
        Returns:
            String containing the summary
        """
        if not text.strip():
            return "No text to summarize"
            
        # Use OpenAI if available, otherwise return a simple summary
        if self.openai_client:
            try:
                model = self.config.get('openai_model', 'gpt-4o-mini')
                response = self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": f"Summarize the following text in {max_length} characters or less:"},
                        {"role": "user", "content": text}
                    ],
                    temperature=0.5,
                    max_tokens=max_length
                )
                
                return response.choices[0].message.content.strip()
            except Exception as e:
                logger.error(f"Error summarizing text with OpenAI: {e}")
                # Fall back to simple summary
        
        # Simple fallback summary (first few sentences)
        sentences = re.split(r'(?<=[.!?])\s+', text)
        simple_summary = " ".join(sentences[:3])
        
        if len(simple_summary) > max_length:
            simple_summary = simple_summary[:max_length - 3] + "..."
            
        return simple_summary