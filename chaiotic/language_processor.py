"""Language processing module for grammar checking and text analysis."""

import re
import json
import time
import os
from pathlib import Path
import logging

# Try to import optional dependencies
try:
    import openai
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

try:
    import language_tool_python
    LANGUAGE_TOOL_AVAILABLE = True
except ImportError:
    LANGUAGE_TOOL_AVAILABLE = False

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
        self.anthropic_client = None
        self.language_tool = None
        
        self._initialize_backends()
    
    def _initialize_backends(self):
        """Initialize the available grammar checking backends."""
        # Initialize OpenAI if available
        if OPENAI_AVAILABLE and self.config.get('openai_api_key'):
            try:
                openai.api_key = self.config['openai_api_key']
                self.openai_client = OpenAI(api_key=self.config['openai_api_key'])
                logger.info("OpenAI client initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize OpenAI client: {e}")
                self.openai_client = None
        
        # Initialize Anthropic if available
        if ANTHROPIC_AVAILABLE and self.config.get('anthropic_api_key'):
            try:
                self.anthropic_client = anthropic.Anthropic(api_key=self.config['anthropic_api_key'])
                logger.info("Anthropic client initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize Anthropic client: {e}")
                self.anthropic_client = None
        
        # Initialize LanguageTool if available
        if LANGUAGE_TOOL_AVAILABLE:
            try:
                language = self.config.get('language', 'en-US')
                self.language_tool = language_tool_python.LanguageTool(language)
                logger.info(f"LanguageTool initialized for language: {language}")
            except Exception as e:
                logger.error(f"Failed to initialize LanguageTool: {e}")
                self.language_tool = None
    
    def check_grammar(self, text, tool='openai'):
        """Check grammar issues in the provided text.
        
        Args:
            text: The text to check
            tool: The tool to use ('openai', 'anthropic', 'language_tool')
            
        Returns:
            Dictionary with correction information
        """
        if not text.strip():
            return {"error": "Empty text provided"}
        
        # Decide which backend to use
        if tool == 'openai' and self.openai_client:
            return self._check_with_openai(text)
        elif tool == 'anthropic' and self.anthropic_client:
            return self._check_with_anthropic(text)
        elif tool == 'language_tool' and self.language_tool:
            return self._check_with_language_tool(text)
        else:
            # Fall back to any available backend
            if self.openai_client:
                return self._check_with_openai(text)
            elif self.anthropic_client:
                return self._check_with_anthropic(text)
            elif self.language_tool:
                return self._check_with_language_tool(text)
            else:
                return {"error": "No grammar checking backend available"}
    
    def _check_with_openai(self, text):
        """Check grammar using OpenAI.
        
        Args:
            text: The text to check
            
        Returns:
            Dictionary with correction information
        """
        try:
            # Prepare the system message
            system_message = """You are a professional editor and grammar checker. 
            Analyze the text for grammar, spelling, punctuation, and style issues. 
            Provide corrections and explanations for each issue found.
            
            Output must be in the following JSON format:
            {
                "original_text": "Original text provided",
                "corrected_text": "Full corrected text",
                "changes": [
                    {
                        "original": "text with error",
                        "corrected": "corrected text",
                        "explanation": "explanation of the correction",
                        "type": "grammar|spelling|punctuation|style"
                    }
                ]
            }
            """
            
            # Make API call
            model = self.config.get('openai_model', 'gpt-4')
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
            result = json.loads(response_text)
            
            # Ensure expected format
            if "original_text" not in result:
                result["original_text"] = text
            
            if "changes" not in result:
                result["changes"] = []
                
            return result
            
        except Exception as e:
            logger.error(f"Error checking grammar with OpenAI: {e}")
            return {
                "original_text": text,
                "error": f"OpenAI API error: {str(e)}",
                "changes": []
            }
    
    def _check_with_anthropic(self, text):
        """Check grammar using Anthropic.
        
        Args:
            text: The text to check
            
        Returns:
            Dictionary with correction information
        """
        try:
            # Prepare the system message
            system_message = """You are a professional editor and grammar checker.
            Analyze the text for grammar, spelling, punctuation, and style issues.
            Provide corrections and explanations for each issue found.
            
            Your response must be in valid JSON format with the following structure:
            {
                "original_text": "Original text provided",
                "corrected_text": "Full corrected text",
                "changes": [
                    {
                        "original": "text with error",
                        "corrected": "corrected text",
                        "explanation": "explanation of the correction",
                        "type": "grammar|spelling|punctuation|style"
                    }
                ]
            }
            """
            
            # Make API call
            model = self.config.get('anthropic_model', 'claude-3-opus-20240229')
            response = self.anthropic_client.messages.create(
                model=model,
                max_tokens=4000,
                system=system_message,
                messages=[
                    {"role": "user", "content": text}
                ],
                temperature=0.3
            )
            
            # Extract JSON from response
            response_text = response.content[0].text
            
            # Find JSON in response
            json_match = re.search(r'```json(.*?)```', response_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(1).strip()
            else:
                # Try to find JSON without code blocks
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                else:
                    json_str = response_text
            
            # Parse response
            result = json.loads(json_str)
            
            # Ensure expected format
            if "original_text" not in result:
                result["original_text"] = text
            
            if "changes" not in result:
                result["changes"] = []
                
            return result
            
        except Exception as e:
            logger.error(f"Error checking grammar with Anthropic: {e}")
            return {
                "original_text": text,
                "error": f"Anthropic API error: {str(e)}",
                "changes": []
            }
    
    def _check_with_language_tool(self, text):
        """Check grammar using LanguageTool.
        
        Args:
            text: The text to check
            
        Returns:
            Dictionary with correction information
        """
        try:
            matches = self.language_tool.check(text)
            
            original_text = text
            corrected_text = text
            changes = []
            
            # Process each match
            for match in matches:
                if not match.replacements:
                    continue
                    
                original = text[match.offset:match.offset + match.errorLength]
                corrected = match.replacements[0]
                explanation = match.message
                match_type = match.ruleIssueType.lower() if match.ruleIssueType else "grammar"
                
                # Add to changes list
                changes.append({
                    "original": original,
                    "corrected": corrected,
                    "explanation": explanation,
                    "type": match_type
                })
                
                # Apply correction to the text
                corrected_text = corrected_text.replace(original, corrected, 1)
            
            return {
                "original_text": original_text,
                "corrected_text": corrected_text,
                "changes": changes
            }
            
        except Exception as e:
            logger.error(f"Error checking grammar with LanguageTool: {e}")
            return {
                "original_text": text,
                "error": f"LanguageTool error: {str(e)}",
                "changes": []
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
                model = self.config.get('openai_model', 'gpt-4')
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