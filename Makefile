.PHONY: run setup clean help build_docs merge_updated all_with_merge

# Python interpreter to use
PYTHON = python
VENV = .venv
VENV_ACTIVATE = source $(VENV)/Scripts/activate

# Directories
TEMP_DIR = temp
OUTPUT_DIR = generated_documents

# Default target
.DEFAULT_GOAL := help

# Run all steps in sequence (original workflow)
all: run build_docs

# Run all steps including the merge/update step
all_with_merge: run build_docs merge_updated

# Step 1: Run the data loading script
run:
	@echo "Running data loading script..."
	@mkdir -p $(TEMP_DIR)
	@$(PYTHON) 01_load_excel_data.py

# Step 2: Build documents from template
build_docs: $(TEMP_DIR)/processed_data.pkl
	@echo "Building documents from template..."
	@mkdir -p $(OUTPUT_DIR)
	@$(PYTHON) 02_build_document_and_header.py

# Step 3: Merge updated footnotes (optional)
merge_updated:
	@echo "Merging updated footnotes from DBL-UpdatedFootnotes..."
	@$(PYTHON) 03_merge_good_format.py

# Set up virtual environment and install dependencies
setup:
	@echo "Setting up virtual environment..."
	@$(PYTHON) -m venv $(VENV)
	@echo "Installing dependencies..."
	@$(VENV_ACTIVATE) && pip install pandas openpyxl python-docx docxcompose

# Clean up temporary files
clean:
	@echo "Cleaning up..."
	@rm -rf __pycache__
	@rm -rf *.pyc
	@rm -rf $(TEMP_DIR)

# Deep clean - removes all generated files but keeps environment
deep-clean: clean
	@echo "Performing deep clean..."
	@rm -rf $(OUTPUT_DIR)
	@rm -rf generated_docs_updated_*

# Help information
help:
	@echo "Available commands:"
	@echo "  make run            - Step 1: Run the data loading script"
	@echo "  make build_docs     - Step 2: Build documents from template"
	@echo "  make merge_updated  - Step 3: Merge updated footnotes (optional)"
	@echo "  make all            - Run steps 1 and 2 in sequence"
	@echo "  make all_with_merge - Run all steps 1, 2, and 3 in sequence"
	@echo "  make setup          - Set up virtual environment and install dependencies"
	@echo "  make clean          - Remove temporary files"
	@echo "  make deep-clean     - Remove all generated files and keep environment"