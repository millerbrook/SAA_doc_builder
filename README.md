# SAA_doc_builder
Build templated documents from a template and spreadsheet for the PRINT project

# SAA Document Builder

A document automation tool for generating structured Word documents from transcription and translation source files.

## Overview

SAA Document Builder automates the creation of standardized documents for the PRINT Project, converting data from spreadsheets into properly formatted Word documents with consistent metadata and content. The tool handles both transcriptions and translations, preserving complex document elements like footnotes and formatting.

## Features

- **Automated Document Generation**: Creates documents from template with proper metadata
- **Metadata Insertion**: Adds Digital ID, citation information, document details
- **Content Integration**: Imports text and footnotes from source documents
- **Language Code Mapping**: Converts language names to standard language codes
- **Customizable Formatting**: Intelligently formats fields with appropriate spacing
- **Makefile-Based Workflow**: Simple commands for processing data and generating documents

## Prerequisites

- Python 3.9+
- Microsoft Word (for viewing generated documents)
- Git (optional, for version control)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd SAA_doc_builder
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/Scripts/activate  # Windows with Git Bash
   # OR
   .\.venv\Scripts\activate  # Windows with CMD/PowerShell
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Directory Structure

```
SAA_doc_builder/
├── data/
│   ├── SAA-DBL-TranscriptionTemplate.docx  # Template document
│   └── transcriptions-translations/  # Source content files
├── temp/
│   └── processed_data.pkl  # Intermediate data
├── generated_documents/  # Output directory
├── 01_process_data.py  # Data processing script
├── 02_build_document_and_header.py  # Document generation script
├── Makefile  # Workflow automation
└── requirements.txt  # Project dependencies
```

## Usage

### Basic Workflow

1. Place your Excel data file in the data directory
2. Update the file path in `01_process_data.py` if necessary
3. Run the complete process:
   ```bash
   make all
   ```

### Step-by-Step Commands

- Process input data:
  ```bash
  make run
  ```

- Generate documents:
  ```bash
  make build_docs
  ```

- Clean temporary files:
  ```bash
  make clean
  ```

## Technical Details

- **Data Processing**: Reads Excel files and prepares data for document generation
- **Document Generation**: Creates Word documents based on templates
- **Metadata Handling**: Inserts metadata fields with proper formatting (Digital ID, Citation, etc.)
- **Content Integration**: Uses `docxcompose` to preserve footnotes when copying content
- **Error Handling**: Comprehensive error reporting for missing files or data

## Configuration

Edit the Makefile to customize:
- Input file paths
- Output directory
- Python executable path
- Other workflow parameters

## Troubleshooting

- **Missing Files**: Ensure all source documents exist in the data/transcriptions-translations directory
- **Footnote Issues**: Make sure docxcompose is properly installed
- **Path Errors**: Check for spaces or special characters in file paths

## License

GNU General License

## Contributors

[Brook Miller, Claude 3.7 Sonnet Thinking]