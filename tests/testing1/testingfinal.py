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
DOC_PATH = Path.cwd() / "testing1" / "finaltest.docx"

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
# 4.  Configure page-numbering rules per section
#     • Section 1 (index 1): no numbers
#     • Section 2 (index 2): restart at 1
#     • Section 3 (index 3): restart at 1
# ----------------------------------------------------------------------
for idx, sec in enumerate(doc.Sections, start=1):
    for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
        sec.Footers(hf_type).LinkToPrevious = False

    if idx == 1:
        for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
            sec.Footers(hf_type).Range.Text = ""
        continue

    if idx == 2:
        sec.PageSetup.DifferentFirstPageHeaderFooter = False
        footer = sec.Footers(c.wdHeaderFooterPrimary)
        pnums = footer.PageNumbers
        pnums.RestartNumberingAtSection = True
        pnums.StartingNumber = 1
        pnums.Add(c.wdAlignParagraphCenter, False)

    if idx >= 3:
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
        pfooter = sec.Footers(c.wdHeaderFooterPrimary)
        ppnums = pfooter.PageNumbers
        ppnums.RestartNumberingAtSection = False
        ppnums.Add(c.wdAlignParagraphCenter, False)
        sec.Footers(c.wdHeaderFooterFirstPage).Range.Text = ""

# ----------------------------------------------------------------------
# 5. Border
# ----------------------------------------------------------------------
sec1 = doc.Sections(1)
borders = sec1.Borders
borders.DistanceFromTop = borders.DistanceFromBottom = 24
borders.DistanceFromLeft = borders.DistanceFromRight = 12

for side in (c.wdBorderTop, c.wdBorderLeft, c.wdBorderBottom, c.wdBorderRight):
    br = borders(side)
    br.LineStyle = c.wdLineStyleThinThickThinMedGap
    br.LineWidth = c.wdLineWidth300pt
    br.Color = c.wdColorAutomatic

# ----------------------------------------------------------------------
# 6.  Save & quit
# ----------------------------------------------------------------------

doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
doc.Close(False)
word.Quit()

print("✅ Created:", DOC_PATH.resolve())
