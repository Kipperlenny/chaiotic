#!/usr/bin/env python3
"""
Grammar and Logic Checker for document files
Main entry point for the application
"""

import argparse
import os
import sys

# Add the current directory to the path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import from our modules
from chaiotic.document_handler import read_document, save_document, create_sample_document, extract_structured_content
from chaiotic.grammar_checker import check_grammar, display_corrections
from chaiotic.utils import preprocess_content
from chaiotic.config import load_config

def main(args=None):
    """Main entry point for the application."""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Grammar and Logic Checker for document files')
    parser.add_argument('--nocache', action='store_true', help='Disable API request caching')
    parser.add_argument('--file', type=str, help='Path to the document file (DOCX or ODT)')
    parser.add_argument('--structured', action='store_true', 
                        help='Use structured paragraph-by-paragraph processing')
    cli_args = parser.parse_args(args)

    # Load configuration and set cache settings
    config = load_config()
    if cli_args.nocache:
        config.set_cache_enabled(False)
    
    # Get file path from command line args or prompt user
    file_path = cli_args.file
    if not file_path:
        file_path = input("Enter the path to the document file (DOCX or ODT): ")

    # If nothing is entered, use test.docx
    if not file_path:
        file_path = 'test.docx'
        print(f"No file path provided. Using default: {file_path}")
        
        # If default file doesn't exist, create a sample file
        if not os.path.exists(file_path):
            print(f"Default file {file_path} not found. Creating sample file...")
            create_sample_document(file_path)

    # Check file extension
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext == '.docx':
        is_docx = True
    elif file_ext == '.odt':
        is_docx = False
    else:
        print("Please provide a valid .docx or .odt file.")
        exit()
        
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        print("Please provide a valid file path.")
        exit()

    # Ask user what they want to do
    print("\nWhat would you like to do with this document?")
    print("1. Check grammar and spelling (GPT-4o-mini)")
    print("2. Check logic and get creative ideas (GPT-4.1) - Coming soon")
    choice = input("Enter your choice (1 or 2): ")

    if not choice:
        choice = "1"
        print("No choice provided. Defaulting to grammar and spelling check.")
    
    if choice == "1":
        # Use structured processing if specified by command line flag
        use_structured = cli_args.structured
        
        # Read the document content with explicit structured content extraction for structured mode
        if use_structured:
            print("Using structured paragraph-by-paragraph processing...")
            content, original_doc, structured_content = read_document(file_path)
            # Check if structured_content is actually a structured content array
            if not isinstance(structured_content, (list, dict)) or len(structured_content) == 0:
                print("No structured content found in document. Extracting structured content...")
                structured_content = extract_structured_content(file_path)
                print(f"Extracted {len(structured_content)} structured elements.")
        else:
            content, original_doc, structured_content = read_document(file_path)
        
        if not content:
            print(f"Failed to extract content from the {file_ext} file.")
            exit()
        
        # Show preview of the content
        preview_length = min(200, len(content))
        print(f"\nPreview of '{file_path}':")
        print("=" * 40)
        print(content[:preview_length] + ("..." if len(content) > preview_length else ""))
        if structured_content and isinstance(structured_content, (list, dict)):
            print(f"Document contains {len(structured_content)} structured elements.")
        elif structured_content:
            print("Document contains structured content.")
        print("=" * 40)
        
        # Preprocess the content
        content = preprocess_content(content)
        
        print(f"Checking grammar and spelling of German text in {file_path}...")
        # Check grammar and spelling
        corrections = check_grammar(content, structured_content, use_structured)
        
        # Display corrections in a readable format
        if corrections:
            display_corrections(corrections)
            
            # Ask if user wants to save corrections
            save_prompt = input("\nDo you want to save corrections to a new document? (y/n): ").lower()
            if save_prompt == 'y' or save_prompt == 'yes':
                # Use document_handler's save_document to preserve formatting better
                try:
                    # Ensure all corrections have the correct structure
                    valid_corrections = []
                    for correction in corrections:
                        # Check if correction is a string (error case)
                        if isinstance(correction, str):
                            print(f"Warning: Found string instead of correction object: {correction}")
                            continue
                        # Only include valid correction objects
                        if isinstance(correction, dict) and 'original' in correction and 'corrected' in correction:
                            valid_corrections.append(correction)
                        else:
                            print(f"Warning: Skipping invalid correction: {correction}")
                    
                    # Save the corrected document, preserving original formatting
                    doc_file = save_document(file_path, valid_corrections, original_doc, is_docx=is_docx)
                    
                    # Also save text and JSON versions for reference
                    from chaiotic.utils import save_document as utils_save_document
                    json_file, text_file, _ = utils_save_document(
                        file_path, valid_corrections, original_doc, is_docx=is_docx
                    )
                    
                    print(f"\nCorrected file (with preserved formatting) saved to: {doc_file}")
                    print(f"Corrections data saved to: {json_file}")
                    print(f"Plain text version saved to: {text_file}")
                except Exception as e:
                    print(f"Error saving document with preserved formatting: {str(e)}")
                    print("Falling back to basic text saving...")
                    try:
                        # Fallback to the utils version if document_handler version fails
                        from chaiotic.utils import save_document as utils_save_document
                        json_file, text_file, doc_file = utils_save_document(
                            file_path, valid_corrections, original_doc, is_docx=is_docx
                        )
                        if doc_file:
                            print(f"\nCorrections saved to: {json_file}")
                            print(f"Corrected text saved to: {text_file}")
                            print(f"Basic corrected file saved to: {doc_file}")
                        else:
                            print("Could not create corrected file.")
                    except Exception as e2:
                        print(f"Fallback saving also failed: {str(e2)}")
                        print("Could not save corrections to file.")
            else:
                print("Corrections not saved.")
        else:
            print("Failed to process corrections.")
    
    elif choice == "2":
        print(f"Checking logic and generating creative ideas using GPT-4.1...")
        print("This feature is coming soon. Currently not implemented.")
        # Future implementation
        # from chaiotic.logic_checker import check_logic
        # check_logic(content)
    
    else:
        print("Invalid choice. Exiting.")

if __name__ == "__main__":
    main()