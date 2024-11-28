# Templatized
A Python tool for converting loan documents into templated versions while preserving formatting

# README.md
# Loan Document Templater

A Python tool for converting loan documents into templated versions while preserving exact formatting.

## Features

- Converts Word documents to templates with variable placeholders
- Preserves all document formatting including:
  - Fonts and sizes
  - Paragraph styling
  - Text effects
  - Spacing and indentation
- Automatically identifies variables using AI
- Generates reusable templates
- Maintains document integrity

## Installation


bash
pip install -r requirements.txt
unknown

Copy Code


## Usage

python
from src.processor import LoanDocumentProcessor

Initialize processor

processor = LoanDocumentProcessor(api_key="your-openai-api-key")

Process document

docx_path = "loan_agreement.docx"
output_dir = "processed_documents"
template_path, variables_path = processor.process_loan_document(docx_path, output_dir)
unknown

Copy Code


## Requirements

- Python 3.7+
- OpenAI API key
- Required packages listed in requirements.txt

## License

MIT License
