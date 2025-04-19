# Chaiotic

A grammar and spelling checker for German documents that uses OpenAI's language models for high-quality corrections.

## Features

- **Grammar and Spelling Checking**: Uses GPT-4o-mini to provide high-quality corrections for German texts
- **Document Support**: Works with DOCX and ODT files
- **Structured Processing**: Option to process documents paragraph-by-paragraph for more detailed corrections
- **Format Preservation**: Maintains document formatting when saving corrections
- **API Request Caching**: Caches API requests to save time and money on repeated checks

## Setup

### Prerequisites

- Python 3.8 or higher
- OpenAI API Key

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/chaiotic.git
cd chaiotic
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set up your OpenAI API key:
```bash
echo "OPENAI_API_KEY=your-api-key-here" > .env
```

## Usage

### Basic Grammar Check

```bash
python main.py --file path/to/your/document.docx
```

### Structured Paragraph-by-Paragraph Analysis

```bash
python main.py --file path/to/your/document.docx --structured
```

### Without API Caching

```bash
python main.py --file path/to/your/document.docx --nocache
```

## How It Works

1. Chaiotic extracts text content from your document
2. For structured mode, it breaks the document into individual paragraphs and elements
3. The text is sent to OpenAI's GPT-4o-mini model for grammar and spell checking
4. Corrections are displayed in the terminal
5. Corrected content is saved in three formats:
   - JSON file with detailed corrections
   - Plain text file with corrected text
   - Corrected document file in the original format

## Planned Features

- **Logic Checking**: Analyze document logic and structure using GPT-4.1
- **Creative Suggestions**: Get creative improvement ideas for your text
- **Multiple Languages**: Support for languages beyond German
- **Custom Ruleset**: Define your own grammar and style rules
- **Web Interface**: User-friendly web interface for document processing

## Dependencies

- `python-docx`: For DOCX file handling
- `odfpy`: For ODT file handling
- `openai`: OpenAI API client
- `python-dotenv`: For environment variable management
- `nltk`: For text chunking (optional)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.