"""Language processing module for grammar checking and text analysis."""

import re
import json
import time
import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime
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
    
    def apply_corrections_to_odt(self, odt_file_path, corrections, output_path=None):
        """Apply corrections to an ODT document with tracked changes.
        
        Args:
            odt_file_path: Path to the original ODT document
            corrections: List of corrections to apply
            output_path: Path to save the corrected document (if None, creates a new path)
            
        Returns:
            Path to the saved document
        """
        if not output_path:
            file_base, file_ext = os.path.splitext(odt_file_path)
            output_path = f"{file_base}_corrected{file_ext}"
        
        # Register XML namespaces
        ET.register_namespace('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0')
        ET.register_namespace('office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0')
        ET.register_namespace('style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0')
        ET.register_namespace('fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0')
        ET.register_namespace('dc', 'http://purl.org/dc/elements/1.1/')
        
        # Create a temporary directory for ODT manipulation
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extract ODT file (it's a ZIP file)
            try:
                with zipfile.ZipFile(odt_file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Parse content.xml
                content_file = os.path.join(temp_dir, 'content.xml')
                tree = ET.parse(content_file)
                root = tree.getroot()
                
                # ODT namespaces
                ns = {
                    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
                    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
                }
                
                # Enable tracked changes
                # Find office:text element
                office_text = root.find('.//office:text', ns)
                if office_text is None:
                    raise ValueError("Could not find office:text element in ODT file")
                
                # Create tracked-changes element if it doesn't exist
                tracked_changes = office_text.find('.//text:tracked-changes', ns)
                if tracked_changes is None:
                    tracked_changes = ET.SubElement(office_text, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}tracked-changes')
                
                # Get all paragraphs
                paragraphs = office_text.findall('.//text:p', ns)
                logger.info(f"Found {len(paragraphs)} paragraphs in ODT document")
                
                # Get current time for tracked changes
                now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                user = "Grammar Checker"
                
                # Generate unique change IDs
                change_id_counter = 1
                
                # Prepare a dictionary to track which paragraph matches which correction
                para_corrections = {}
                
                # First pass: identify which paragraphs contain the original text from corrections
                for i, para in enumerate(paragraphs):
                    para_text = ''.join(para.itertext())
                    para_id = f"p{i+1}"
                    
                    for j, correction in enumerate(corrections):
                        # Handle different correction formats
                        if isinstance(correction, dict):
                            if 'original' in correction:
                                original = correction.get('original', '')
                            elif 'error' in correction:
                                original = correction.get('error', '')
                            else:
                                continue
                                
                            if original in para_text:
                                if para_id not in para_corrections:
                                    para_corrections[para_id] = []
                                para_corrections[para_id].append(correction)
                
                logger.info(f"Matched corrections to {len(para_corrections)} paragraphs")
                
                # Second pass: apply corrections with tracked changes
                for i, para in enumerate(paragraphs):
                    para_id = f"p{i+1}"
                    
                    if para_id in para_corrections:
                        # This paragraph has corrections
                        para_corrections_list = para_corrections[para_id]
                        para_text = ''.join(para.itertext())
                        logger.info(f"Applying {len(para_corrections_list)} corrections to paragraph {para_id}")
                        
                        # Clear paragraph content (we'll rebuild it)
                        for child in list(para):
                            para.remove(child)
                        
                        # Apply each correction
                        remaining_text = para_text
                        
                        for correction in para_corrections_list:
                            # Handle different correction formats
                            if 'original' in correction:
                                original = correction.get('original', '')
                                corrected = correction.get('corrected', '')
                            elif 'error' in correction:
                                original = correction.get('error', '')
                                corrected = correction.get('correction', '')
                            else:
                                continue
                            
                            if original in remaining_text:
                                # Split at the correction point
                                before, original_match, after = remaining_text.partition(original)
                                
                                if before:
                                    # Add text before the correction
                                    span = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span')
                                    span.text = before
                                
                                # Create change ID for this correction
                                change_id = f"ct{change_id_counter}"
                                change_id_counter += 1
                                
                                # Add change entry to tracked-changes
                                change = ET.SubElement(tracked_changes, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}changed-region')
                                change.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}id', change_id)
                                
                                # Add deletion
                                deletion = ET.SubElement(change, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}deletion')
                                creator = ET.SubElement(deletion, '{urn:oasis:names:tc:opendocument:xmlns:dc:1.1}creator')
                                creator.text = user
                                date = ET.SubElement(deletion, '{urn:oasis:names:tc:opendocument:xmlns:dc:1.1}date')
                                date.text = now
                                deletion_text = ET.SubElement(deletion, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p')
                                deletion_span = ET.SubElement(deletion_text, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span')
                                deletion_span.text = original
                                
                                # Add insertion
                                insertion = ET.SubElement(change, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}insertion')
                                creator = ET.SubElement(insertion, '{urn:oasis:names:tc:opendocument:xmlns:dc:1.1}creator')
                                creator.text = user
                                date = ET.SubElement(insertion, '{urn:oasis:names:tc:opendocument:xmlns:dc:1.1}date')
                                date.text = now
                                
                                # Mark deletion in document with change-start
                                change_start = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}change-start')
                                change_start.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}change-id', change_id)
                                
                                # Add original text (will be shown as deleted)
                                deletion_span = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}deletion')
                                deletion_span.text = original
                                
                                # Mark end of deletion
                                change_end = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}change-end')
                                change_end.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}change-id', change_id)
                                
                                # Add corrected text (shown as inserted)
                                insertion_span = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}insertion')
                                insertion_span.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}change-id', change_id)
                                insertion_span.text = corrected
                                
                                # Update remaining text
                                remaining_text = after
                            else:
                                logger.warning(f"Could not find '{original}' in paragraph text")
                        
                        # Add any remaining text
                        if remaining_text:
                            span = ET.SubElement(para, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span')
                            span.text = remaining_text
                
                # Write back the modified content.xml
                tree.write(content_file, encoding='UTF-8', xml_declaration=True)
                
                # Create the new ODT file
                with zipfile.ZipFile(output_path, 'w') as zipf:
                    for root_dir, _, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root_dir, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            zipf.write(file_path, arcname)
                
                logger.info(f"Successfully saved ODT with tracked changes to {output_path}")
                return output_path
                
            except Exception as e:
                logger.error(f"Error applying corrections to ODT: {e}")
                import traceback
                traceback.print_exc()
                
                # Fallback to text file
                output_txt = os.path.splitext(output_path)[0] + ".txt"
                with open(output_txt, 'w', encoding='utf-8') as f:
                    f.write("CORRECTION SUGGESTIONS:\n\n")
                    for i, correction in enumerate(corrections, 1):
                        if isinstance(correction, dict):
                            if 'original' in correction:
                                original = correction.get('original', 'N/A')
                                corrected = correction.get('corrected', 'N/A')
                            elif 'error' in correction:
                                original = correction.get('error', 'N/A')
                                corrected = correction.get('correction', 'N/A')
                            else:
                                continue
                            explanation = correction.get('explanation', 'No explanation provided')
                            f.write(f"{i}. ORIGINAL: {original}\n")
                            f.write(f"   CORRECTED: {corrected}\n")
                            f.write(f"   EXPLANATION: {explanation}\n\n")
                
                logger.info(f"Saved corrections to text file: {output_txt}")
                return output_txt
    
    def apply_corrections_to_document(self, file_path, corrections, output_path=None):
        """Apply corrections to a document with track changes.
        
        Args:
            file_path: Path to the original document
            corrections: List of corrections to apply
            output_path: Path for the output file (generated if None)
            
        Returns:
            Path to the saved document
        """
        # Determine file type
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Default output path
        if not output_path:
            file_base = os.path.splitext(file_path)[0]
            output_path = f"{file_base}_corrected{file_ext}"
        
        # Apply corrections based on file type
        if file_ext == '.odt':
            return self.apply_corrections_to_odt(file_path, corrections, output_path)
        else:
            # Handle other file types or return an error
            logger.error(f"Unsupported file type for tracked changes: {file_ext}")
            return None