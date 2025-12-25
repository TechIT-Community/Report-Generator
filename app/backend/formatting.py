"""
Formatting and bookmarking utilities for Word automation.
"""
from win32com.client import constants as c

def set_format(selection, font="Times New Roman", size=12, bold=False, align=None, underline=None):
    """
    Sets the formatting for the current selection in Word.
    
    :param selection: The Word Selection object.
    :param font: The font name. Defaults to "Times New Roman".
    :param size: The font size. Defaults to 12.
    :param bold: Whether the text should be bold. Defaults to False.
    :param align: The paragraph alignment constant. Defaults to None (keeps current).
    :param underline: The underline constant. Defaults to None (keeps current).
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

def add_bookmark(doc, selection, name: str, placeholder: str = "___"):
    """
    Inserts a placeholder text and adds a bookmark to it.
    
    :param doc: The Word Document object.
    :param selection: The Word Selection object.
    :param name: The name of the bookmark.
    :param placeholder: The placeholder text to insert. Defaults to "___".
    """
    selection.TypeText(placeholder)
    bm_range = selection.Range.Duplicate
    # Calculate start position based on placeholder length
    # Note: Range is slightly dynamic, ensuring we capture the just-inserted text
    bm_start = bm_range.Start - len(placeholder) 
    bm_range = doc.Range(bm_start, bm_start + len(placeholder))
    doc.Bookmarks.Add(name, bm_range)
