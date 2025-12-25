"""
make_test10_sections.py
-----------------------
Creates test10.docx in the folder where you run this script.

Prerequisites
-------------
• Windows with desktop Microsoft Word installed
• pip install pywin32
"""

import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path

# ----------------------------------------------------------------------
# 1.  Where to save
# ----------------------------------------------------------------------
DOC_PATH = Path.cwd() / "testing1" / "test11.docx"

# ----------------------------------------------------------------------
# 2.  Start Word (hidden)
# ----------------------------------------------------------------------
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()

# Wipe any default paragraph so we start truly at character 0
doc.Content.Delete()

# ----------------------------------------------------------------------
# 3.  Write 10 pages, inserting section breaks after page 3 and 7
# ----------------------------------------------------------------------
cursor = doc.Range(0, 0)
blank_pages = {4, 8, 11}                 # pages right after each section break
for page in range(1, 16):
    cursor.Collapse(c.wdCollapseEnd)

    # ----- TEXT GENERATION (only change) -----
    if page not in blank_pages:
        cursor.InsertAfter(f"Page {page}: filler text. " * 100)
    # -----------------------------------------

    cursor.InsertParagraphAfter()

    if page in (3, 7, 10):
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdSectionBreakNextPage)
    elif page < 15:
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdPageBreak)

# ----------------------------------------------------------------------
# 4.  Configure page‑numbering rules per section
#     • Section 1 (idx 1): no numbers
#     • Section 2+       : restart at 1, but *hide* number on first page
# ----------------------------------------------------------------------
for idx, sec in enumerate(doc.Sections, start=1):
    footer = sec.Footers(c.wdHeaderFooterPrimary)
    footer.LinkToPrevious = False                     # isolate every section

    if idx == 1:
        # blank footer → no page number at all
        footer.Range.Text = ""
    else:
        # Add centered page number; restart at 1
        pnums = footer.PageNumbers
        if idx == 2:
            pnums.RestartNumberingAtSection = True
            pnums.StartingNumber = 1
            pnums.Add(c.wdAlignParagraphCenter, True)
            show_first = True
        else:
            pnums.RestartNumberingAtSection = False
        # ShowOnFirstPage=False  ⇒ omit number on page after the section break
        pnums.Add(c.wdAlignParagraphCenter, False)

# ----------------------------------------------------------------------
# 5.  Save & quit
# ----------------------------------------------------------------------
doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
doc.Close(False)
word.Quit()

print("✅ Created:", DOC_PATH.resolve())
