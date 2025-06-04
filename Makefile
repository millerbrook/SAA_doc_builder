.PHONY: run setup clean help build_docs all

# Python interpreter to use
PYTHON = python
VENV = .venv
VENV_ACTIVATE = source $(VENV)/Scripts/activate

# Directories
TEMP_DIR = temp
OUTPUT_DIR = generated_documents

# Default target
.DEFAULT_GOAL := help

# Run all steps
all: run build_docs

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

# Set up virtual environment and install dependencies
setup:
	@echo "Setting up virtual environment..."
	@$(PYTHON) -m venv $(VENV)
	@echo "Installing dependencies..."
	@$(VENV_ACTIVATE) && pip install pandas openpyxl python-docx

# Clean up temporary files and virtual environment
clean:
	@echo "Cleaning up..."
	@rm -rf __pycache__
	@rm -rf *.pyc
	@rm -rf $(TEMP_DIR)

# Deep clean - removes all generated files and environment
deep-clean: clean
	@echo "Performing deep clean..."
	@rm -rf $(OUTPUT_DIR)
# @rm -rf $(VENV)

# Help information
help:
	@echo "Available commands:"
	@echo "  make run        - Step 1: Run the data loading script"
	@echo "  make build_docs - Step 2: Build documents from template"
	@echo "  make all        - Run all steps in sequence"
	@echo "  make setup      - Set up virtual environment and install dependencies"
	@echo "  make clean      - Remove temporary files"
	@echo "  make deep-clean - Remove all generated files and environment"