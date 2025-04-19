"""
Grammar and Logic Checker for document files
Package initialization
"""

__version__ = '0.1.0'

from .document_handler import read_document, save_document, create_sample_document, extract_structured_content
from .document_writer import save_correction_outputs
from .grammar_checker import check_grammar, display_corrections
from .utils import preprocess_content, split_text_into_chunks, sanitize_response
from .config import load_config

# Package exports
__all__ = [
    'read_document', 
    'save_document', 
    'create_sample_document',
    'extract_structured_content',
    'save_correction_outputs',
    'check_grammar', 
    'display_corrections',
    'preprocess_content',
    'load_config'
]