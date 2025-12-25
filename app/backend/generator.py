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

from .content_static import generate_static_pages
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
    - Static Content (Title Page, Certificates, etc.) via `generate_static_pages`
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

        # Generate the structure immediately (Legacy behavior)
        print(f"DEBUG: Calling generate_static_pages. Doc: {doc}, Word: {word}, BaseDir: {BASE_DIR}")
        generate_static_pages(doc, word, BASE_DIR)


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


def save_document():
    """
    Finalizes the document and saves it to the reports folder.
    - Updates page numbers (TOC, Index).
    - Updates all Word fields (formulas, refs).
    - Updates Header/Footer fields.
    - Saves as `template.docx`.
    - Displays Success/Error message box.
    """
    if not doc:
        return
        
    try:
        update_index_page_numbers(doc)
        
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
