"""
Formatting and bookmarking utilities for Word automation.
Provides helper functions to apply styles and create named ranges (bookmarks) programmatically.
"""

from win32com.client import constants as c


# =================================================================================================
#                                         FORMATTING TEXT
# =================================================================================================

def set_format(selection, font="Times New Roman", size=12, bold=False, align=None, underline=None):
    """
    Sets the formatting properties for the current selection in Word.
    
    :param selection: The Word Selection object.
    :param font: The font name. Defaults to "Times New Roman".
    :param size: The font size in points. Defaults to 12.
    :param bold: Boolean for bold text. Defaults to False.
    :param align: The paragraph alignment constant (e.g., c.wdAlignParagraphCenter). Defaults to None (unchanged).
    :param underline: The underline constant (e.g., c.wdUnderlineSingle). Defaults to None (unchanged).
    """
    if font is not None:
        selection.Font.Name = font
    if size is not None:
        selection.Font.Size = size
    if bold is not None:
        selection.Font.Bold = bold
    if align is not None:
        selection.ParagraphFormat.Alignment = align
    if underline is not None:
        selection.Font.Underline = underline


# =================================================================================================
#                                       MARKING BOOKMARKS
# =================================================================================================

def add_bookmark(doc, selection, name: str, placeholder: str = "___"):
    """
    Inserts a placeholder string into the document at the current selection 
    and wraps it in a named Bookmark for later replacement.
    
    Logic:
    1.  Types the placeholder text.
    2.  Calculates the range of that text.
    3.  Adds a Bookmark to that range.
    
    :param doc: The Word Document object.
    :param selection: The Word Selection object.
    :param name: The unique name for the bookmark.
    :param placeholder: The text to insert (e.g., "___" or "___\n").
    """
    selection.TypeText(placeholder)
    bm_range = selection.Range.Duplicate
    
    # Calculate start position based on placeholder length
    # ensuring we capture the just-inserted text
    bm_start = bm_range.Start - len(placeholder) 
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    
    doc.Bookmarks.Add(name, bm_range)
