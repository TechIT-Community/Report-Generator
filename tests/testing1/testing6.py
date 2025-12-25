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
DOC_PATH = Path.cwd() / "testing1" / "test101.docx"

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
cursor = doc.Range(0, 0)                             # single "insertion-point"
for page in range(1, 11):
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertAfter(f"Random text filling page {page}. " * 150)
    cursor.InsertParagraphAfter()

    # Decide which break to insert
    if page in (3, 7):
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdSectionBreakNextPage)  # new section + new page
    elif page < 10:                                   # simple page break
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdPageBreak)

# ----------------------------------------------------------------------
# 4.  Configure page-numbering rules per section
#     • Section 1 (index 1): no numbers
#     • Section 2 (index 2): restart at 1
#     • Section 3 (index 3): restart at 1
# ----------------------------------------------------------------------
for idx, sec in enumerate(doc.Sections, start=1):
    footer = sec.Footers(c.wdHeaderFooterPrimary)
    footer.LinkToPrevious = False                     # isolate every section

    if idx == 1:
        # blank footer → no page number
        footer.Range.Text = ""
    else:
        # Add centered page number; restart at 1
        pnums = footer.PageNumbers
        pnums.RestartNumberingAtSection = True
        pnums.StartingNumber = 1
        pnums.Add(c.wdAlignParagraphCenter, True)     # (alignment, show-on-first)

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
