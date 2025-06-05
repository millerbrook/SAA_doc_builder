import os
import pandas as pd
from docx import Document
import re
import shutil
import datetime
import docx.shared

# Define paths
current_dir = os.path.dirname(os.path.abspath(__file__))
data_folder = os.path.join(current_dir, 'data')
template_file = os.path.join(data_folder, 'SAA-DBL-TranscriptionTemplate.docx')
input_data_file = os.path.join(current_dir, 'temp', 'processed_data.pkl')

# Define output directory for generated documents
OUTPUT_DIR = os.path.join(current_dir, "generated_documents")
os.makedirs(OUTPUT_DIR, exist_ok=True)

def format_citation_text(row, content_type):
    """
    Format citation text based on volume and document type
    """
    doc_number = str(row.get('DBL - Doc number', ''))
    date_value = str(row.get('Date', ''))
    volume = str(row.get('volume', '')).strip()
    
    # Use correct capitalization for range columns
    transcript_range = str(row.get('Transcript range', ''))
    translate_range = str(row.get('Translate range', ''))
    
    # Base citation parts that don't change
    author = 'James W. Lowry'
    
    if volume == 'I':
        book_info = "Documents of Brotherly Love: Dutch Mennonite Aid to Swiss Anabaptists Volume 1, 1635-1709, edited by David J. Rempel Smucker and John L. Ruth (Millersburg, OH: Ohio Amish Library, 2007)"
    elif volume == 'II':
        book_info = "Documents of Brotherly Love: Dutch Mennonite Aid to Swiss Anabaptists Volume II, 1710-1711 (Millersburg, OH: Ohio Amish Library, 2015)"
    else:
        # Default case if volume is not recognized
        book_info = "Documents of Brotherly Love: Dutch Mennonite Aid to Swiss Anabaptists (Millersburg, OH: Ohio Amish Library)"
    
    # Format differs based on document type
    if content_type == 'Transcript':
        doc_type = "transcription"
        page_range = transcript_range
    else:  # Translation
        doc_type = "translation"
        page_range = translate_range
    
    # Assemble the full citation
    citation = f'{author}, "Document {doc_number}, {date_value}, {doc_type}," in {book_info}, {page_range}.'
    
    return citation

def update_digital_id_in_header(header, new_digital_id):
    """Update the Digital ID in a header, checking both paragraphs and tables"""
    # First check header paragraphs
    for para in header.paragraphs:
        if "Digital ID:" in para.text:
            print(f"Found Digital ID in header paragraph: '{para.text}'")
            # Clear the paragraph
            for run in para.runs:
                run.clear()
            # Create new text with proper formatting
            run1 = para.add_run("Digital ID: ")
            run1.bold = True
            para.add_run(new_digital_id)
            print(f"Updated paragraph text to: 'Digital ID: {new_digital_id}'")
            return True
    
    # If not found in paragraphs, check header tables
    for table in header.tables:
        print(f"Examining header table with {len(table.rows)} rows")
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "Digital ID:" in para.text:
                        print(f"Found Digital ID in table cell: '{para.text}'")
                        # Clear the paragraph
                        for run in para.runs:
                            run.clear()
                        # Create new text with proper formatting
                        run1 = para.add_run("Digital ID: ")
                        run1.bold = True
                        para.add_run(new_digital_id)
                        print(f"Updated cell text to: 'Digital ID: {new_digital_id}'")
                        return True
    
    print("WARNING: Digital ID not found in header paragraphs or tables")
    return False

def add_metadata_fields(doc, row, content_type):
    """
    Update metadata fields in the document:
    - Copyright: Leave as is in template
    - Other fields: Add tabs after ": " and add value, or delete line if no value exists
    
    Tab rules:
    - 5 or fewer characters: 3 tabs
    - 6-11 characters: 2 tabs
    - 12+ characters: 1 tab
    """
    print("Processing metadata fields...")
    
    # Define metadata fields to process (excluding Copyright)
    metadata_fields = [
        {"tag": "Date", "column": "Date"},
        {"tag": "Sender", "column": "Sender"},
        {"tag": "Sender Place", "column": "Sender Place"},
        {"tag": "Receiver", "column": "Receiver"},
        {"tag": "Receiver Place", "column": "Receiver Place"}
    ]
    
    # Handle Language field (might differ based on content type)
    if content_type == 'Translate':
        language_value = "English"
    else:
        language_value = row.get('Language', '')
    
    metadata_fields.append({"tag": "Language", "column": None, "value": language_value})
    
    # Find paragraphs that begin with each metadata tag
    paragraphs_to_delete = []
    
    for i, para in enumerate(doc.paragraphs):
        # Skip empty paragraphs
        if not para.text.strip():
            continue
            
        # Skip Copyright paragraph - leave it as is
        if para.text.strip().startswith("Copyright:"):
            print(f"Keeping original Copyright text: '{para.text}'")
            continue
        
        # Check for metadata fields
        for field in metadata_fields:
            if para.text.strip().startswith(f"{field['tag']}:"):
                print(f"Found {field['tag']} field: '{para.text}'")
                
                # Get the value from dataframe or from pre-defined value
                if 'value' in field:
                    value = field['value']
                else:
                    value = row.get(field['column'], '')
                
                # If the value is None, NaN, or empty string, mark paragraph for deletion
                if pd.isna(value) or str(value).strip() == '':
                    print(f"  No value or NaN for {field['tag']}, will remove line")
                    paragraphs_to_delete.append(i)
                else:
                    # Clear the paragraph and rebuild it
                    for run in para.runs:
                        run.clear()
                    
                    # Add tag with bold formatting
                    bold_run = para.add_run(f"{field['tag']}: ")
                    bold_run.bold = True
                    
                    # Add tabs based on tag length
                    if len(field['tag']) <= 5:
                        para.add_run("\t\t\t")  # 3 tabs for 5 or fewer chars
                    elif len(field['tag']) < 12:
                        para.add_run("\t\t")     # 2 tabs for 6-11 chars
                    else:
                        para.add_run("\t")       # 1 tab for 12+ chars
                    
                    # Add value
                    para.add_run(str(value))
                    print(f"  Updated with value: '{value}'")
                
                break
    
    # Delete paragraphs marked for deletion (in reverse order to maintain indices)
    for idx in sorted(paragraphs_to_delete, reverse=True):
        p = doc.paragraphs[idx]._p
        p.getparent().remove(p)
        print(f"Removed paragraph at index {idx}")
    
    return doc

def remove_instruction_text(doc):
    """
    Remove the specific instruction text from the document
    """
    instruction_text = "<For the following PRINT fields, if they aren't available for a document, remove them."
    
    print("Searching for instruction text...")
    
    # Find paragraphs containing any part of the instruction text (more robust)
    paragraphs_to_delete = []
    for i, para in enumerate(doc.paragraphs):
        # Check for the start of the instruction text to identify the paragraph
        if instruction_text in para.text:
            print(f"Found instruction text at paragraph {i}: '{para.text[:50]}...'")
            paragraphs_to_delete.append(i)
        # Also check for sections of text in case it's broken up
        elif "<For the following PRINT fields" in para.text or "remove them. For example" in para.text:
            print(f"Found partial instruction text at paragraph {i}: '{para.text[:50]}...'")
            paragraphs_to_delete.append(i)
    
    # Delete paragraphs with instruction text (in reverse order to maintain indices)
    if paragraphs_to_delete:
        for idx in sorted(paragraphs_to_delete, reverse=True):
            p = doc.paragraphs[idx]._p
            p.getparent().remove(p)
            print(f"Removed instruction paragraph at index {idx}")
        print(f"Removed {len(paragraphs_to_delete)} paragraphs containing instruction text")
    else:
        print("No instruction text found")
    
    return doc

def create_document(row, content_type):
    """Create a new document based on the template and row data"""
    # Skip if the specified column has no value
    if pd.isna(row[content_type]):
        return None
    
    # Get values for filename
    filename_value = str(row.get('Filename', ''))
    
    # Determine language code for the filename
    if content_type == 'Translate':
        # Always use EN for translations
        language_code = 'EN'
    else:
        # For transcriptions, map language names to codes
        language_value = str(row.get('Language', ''))
        
        # Check for combined language cases
        if language_value.lower() == 'german/dutch':
            language_code = 'DE-NL'
        elif language_value.lower() == 'french; german':
            language_code = 'FR-DE'
        elif language_value.lower() == 'german':
            language_code = 'DE'
        elif language_value.lower() == 'dutch':
            language_code = 'NL'
        elif language_value.lower() == 'french':
            language_code = 'FR'
        else:
            # Keep original value for other languages
            language_code = language_value
    
    # Create a filename using the specified format with appropriate language code
    safe_filename = f"{filename_value}_{language_code}_{content_type.lower()}.docx"
    # Clean up filename to be valid (remove illegal characters)
    safe_filename = re.sub(r'[<>:"/\\|?*]', '', safe_filename)
    # Replace spaces with underscores
    safe_filename = safe_filename.replace(" ", "_")
    
    output_path = os.path.join(OUTPUT_DIR, safe_filename)
    
    # Digital ID and other values still needed for document content
    digital_id = str(row.get('Digital ID', 'unknown'))
    doc_number = str(row.get('DBL - Doc number', 'unknown'))
    date_value = str(row.get('Date', 'unknown'))
    
    # Create a copy of the template
    print(f"Creating document: {output_path}")
    shutil.copy(template_file, output_path)
    
    # Open the document
    doc = Document(output_path)
    doc = remove_instruction_text(doc)
    
    # Process the header sections
    for section_idx, section in enumerate(doc.sections):
        header = section.header
        update_digital_id_in_header(header, digital_id)
    
    # Generate the enhanced citation text
    citation_text = format_citation_text(row, content_type)
    
    # Find and replace Citation line in the document body
    citation_found = False
    for para in doc.paragraphs:
        if para.text.strip().startswith("Citation:"):
            print(f"Found Citation line: '{para.text}'")
            # Clear the paragraph
            for run in para.runs:
                run.clear()
            
            # Create new text with proper formatting
            run1 = para.add_run("Citation: ")
            run1.bold = True
            para.add_run(citation_text)
            
            print(f"Replaced with Citation: {citation_text}")
            citation_found = True
            break
    
    if not citation_found:
        print("Warning: Citation line not found in document")
    
    # Process metadata fields (after citation handling)
    doc = add_metadata_fields(doc, row, content_type)
    
    # Add content type indication to the main document body
    paragraph = doc.add_paragraph()
    paragraph.add_run(f"Document Type: {content_type}").bold = True
    
    # REMOVED: Timestamp generation code
    # No longer adding timestamp at the end of documents
    
    # Save the modified document
    doc.save(output_path)
    
    return output_path

def main():
    # Check if the processed data file exists
    if not os.path.exists(input_data_file):
        print(f"Error: Processed data file not found: {input_data_file}")
        print("Please run step 1 first (make run)")
        return
    
    # Load the processed data
    try:
        df = pd.read_pickle(input_data_file)
        print(f"Loaded data with {len(df)} rows from {input_data_file}")
    except Exception as e:
        print(f"Error loading processed data: {e}")
        return
    
    # Track created documents
    transcript_docs = []
    translate_docs = []
    
    # Process each row in the dataframe
    for idx, row in df.iterrows():
        # Create transcript document if column has value
        if 'Transcript' in df.columns and not pd.isna(row.get('Transcript', pd.NA)):
            doc_path = create_document(row, 'Transcript')
            if doc_path:
                transcript_docs.append(doc_path)
                print(f"Created Transcript document: {os.path.basename(doc_path)}")
        
        # Create translate document if column has value
        if 'Translate' in df.columns and not pd.isna(row.get('Translate', pd.NA)):
            doc_path = create_document(row, 'Translate')
            if doc_path:
                translate_docs.append(doc_path)
                print(f"Created Translate document: {os.path.basename(doc_path)}")
    
    print(f"\nGenerated {len(transcript_docs)} Transcript documents")
    print(f"Generated {len(translate_docs)} Translate documents")
    print(f"Documents saved to: {os.path.abspath(OUTPUT_DIR)}")

if __name__ == "__main__":
    main()
