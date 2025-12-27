"""
Core driver module for the Report Generator.
Initializes Word, sets up the document, and exposes the main API for the GUI.
"""

import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path
from CTkMessagebox import CTkMessagebox
import pythoncom
import time

from .content_static import generate_static_pages_part1, generate_static_pages_part2
from .content_dynamic import replace_bookmarks as replace_bookmarks_dynamic, update_index_page_numbers
from .utils import cm_to_pt

# =================================================================================================
#                                       CONFIGURATION
# =================================================================================================

# Paths
# This file is in app/backend/, so parent.parent is app/
BASE_DIR = Path(__file__).resolve().parent.parent 
ASSET_DIR = BASE_DIR / "assets"
DOC_PATH = BASE_DIR / "reports" / "template.docx"

# Ensure reports directory exists
DOC_PATH.parent.mkdir(parents=True, exist_ok=True)

# Global Word state
word = None
doc = None
_document_finalized = False # Flag to prevent double-finalization

# =================================================================================================
#                                     INITIALIZATION
# =================================================================================================

def initialize():
    """
    Initializes the Microsoft Word application and creates a new document.
    Does nothing if the document is already initialized.
    
    Sets up:
    - Word Application (Visible)
    - New Document
    - Page Margins
    - Default Fonts
    - Static Content PART 1 ONLY (Title Page, Certificates, Acknowledgement, Abstract).
    
    NOTE: Chapters and References are generated later via `finalize_document()`.
    """
    global word, doc
    if doc:
        return

    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = True
        time.sleep(1) # Wait for Word to initialize fully to prevent RPC errors
        doc = word.Documents.Add()
    except Exception as e:
        print(f"Error initializing Word: {e}")
        word = None
        doc = None

    if doc:
        # --- Initial Setup ---
        try:
            # Global Font Defaults
            doc.Styles(c.wdStyleNormal).Font.Name = "Times New Roman"
            doc.Content.Font.Name = "Times New Roman"
            
            # Margins
            doc.PageSetup.TopMargin = cm_to_pt(1.7)
            doc.PageSetup.BottomMargin = cm_to_pt(1.7)
            doc.PageSetup.LeftMargin = cm_to_pt(2.1)
            doc.PageSetup.RightMargin = cm_to_pt(1.7)
        except Exception as e:
            print(f"Setup error: {e}")

        # Generate PART 1 structure immediately (Title Page → Abstract)
        print(f"DEBUG: Calling generate_static_pages_part1. Doc: {doc}, Word: {word}, BaseDir: {BASE_DIR}")
        generate_static_pages_part1(doc, word, BASE_DIR)


def finalize_document(num_chapters: int):
    """
    Generates (or regenerates) the dynamic parts of the document (TOC, Chapters, References).
    
    If Part 2 already exists, it will be deleted first to allow dynamic chapter count changes.
    This enables users to add/remove chapters and press "Done" multiple times.
    
    :param num_chapters: The final count of chapters from the GUI.
    """
    global doc, word, _document_finalized
    if not doc:
        return
    
    # If Part 2 already exists, delete it first
    if _document_finalized:
        print(f"DEBUG: Regenerating Part 2 - deleting old content first...")
        from .content_static import delete_part2_content
        delete_part2_content(doc)
        _document_finalized = False
        
    print(f"DEBUG: Calling generate_static_pages_part2 with {num_chapters} chapters.")
    generate_static_pages_part2(doc, word, BASE_DIR, num_chapters)
    _document_finalized = True


def is_document_finalized():
    """
    Returns True if the document structure has been finalized (Part 2 generated).
    NOTE: This is now informational only - regeneration is allowed.
    """
    return _document_finalized


# =================================================================================================
#                                         PUBLIC API
# =================================================================================================

def replace_bookmarks(data_dict: dict):
    """
    Updates the document content based on user inputs.
    Delegates to `content_dynamic.replace_bookmarks`.
    
    :param data_dict: Dictionary containing key-value pairs from the GUI inputs.
    """
    if doc:
        replace_bookmarks_dynamic(doc, word, data_dict, ASSET_DIR)


def save_document(num_chapters: int, full_data: dict):
    """
    Finalizes the document and saves it to the reports folder.
    
    :param num_chapters: Number of chapters from GUI tabs.
    :param full_data: Aggregated data from all pages (used for final bookmark replacement).
    
    Steps:
    1. Generate Phase 2 structure (TOC, Chapters, References).
    2. Replace all bookmarks with user data.
    3. Update page numbers (TOC, Index).
    4. Update all Word fields.
    5. Save as `template.docx`.
    """
    if not doc:
        return
        
    try:
        # PHASE 2: Generate Chapters/TOC structure
        finalize_document(num_chapters)
        
        # Replace all bookmarks with aggregated data
        replace_bookmarks(full_data)
        
        # Update page numbers in TOC
        update_index_page_numbers(doc)
        
        # Update Word fields
        doc.Fields.Update()
        for field in doc.Fields:
            field.Update()
            
        for section in doc.Sections:
            section.Headers(c.wdHeaderFooterPrimary).Range.Fields.Update()
            section.Footers(c.wdHeaderFooterPrimary).Range.Fields.Update()
            
        doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
        
        CTkMessagebox(title="Saved", message=f"The report has been successfully saved.\n\nSave Location: {DOC_PATH.resolve()}", icon="check")
        
    except Exception as e:
        CTkMessagebox(title="Error", message=f"Failed to save document: {e}", icon="cancel")
