"""
Core driver module for the Report Generator.
Initializes Word, sets up the document, and exposes the main API.
"""
import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path
from CTkMessagebox import CTkMessagebox
import pythoncom

from .content_static import generate_static_pages
from .content_dynamic import replace_bookmarks as replace_bookmarks_dynamic, update_index_page_numbers
from .utils import cm_to_pt

# Paths
# This file is in app/backend/, so parent.parent is app/
BASE_DIR = Path(__file__).resolve().parent.parent 
ASSET_DIR = BASE_DIR / "assets"
DOC_PATH = BASE_DIR / "reports" / "template.docx"

# Ensure reports directory exists
DOC_PATH.parent.mkdir(parents=True, exist_ok=True)

# Initialize Word
word = None
doc = None

def initialize():
    global word, doc
    if doc:
        return

    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = True
        import time
        time.sleep(1) # Wait for Word to initialize fully
        doc = word.Documents.Add()
    except Exception as e:
        print(f"Error initializing Word: {e}")
        word = None
        doc = None

    if doc:
        # --- Initial Setup (Legacy behavior replication) ---
        # Delete any content (though new doc is empty) and set margins
        try:
            # Global Font Defaults
            doc.Styles(c.wdStyleNormal).Font.Name = "Times New Roman"
            doc.Content.Font.Name = "Times New Roman"
            
            doc.PageSetup.TopMargin = cm_to_pt(1.7)
            doc.PageSetup.BottomMargin = cm_to_pt(1.7)
            doc.PageSetup.LeftMargin = cm_to_pt(2.1)
            doc.PageSetup.RightMargin = cm_to_pt(1.7)
            # doc.Content.Delete() # New doc doesn't need this, but legacy had it.
        except Exception as e:
            print(f"Setup error: {e}")

        # Generate the structure immediately (Monolithic behavior)
        print(f"DEBUG: Calling generate_static_pages. Doc: {doc}, Word: {word}, BaseDir: {BASE_DIR}")
        generate_static_pages(doc, word, BASE_DIR)


def replace_bookmarks(data_dict: dict):
    if doc:
        replace_bookmarks_dynamic(doc, word, data_dict, ASSET_DIR)

def save_document():
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
        
        # Bring GUI to front if possible? CTkMessagebox usually handles parent.
        CTkMessagebox(title="Saved", message=f"The report has been successfully saved.\n\nSave Location: {DOC_PATH.resolve()}", icon="check")
        
    except Exception as e:
        CTkMessagebox(title="Error", message=f"Failed to save document: {e}", icon="cancel")
