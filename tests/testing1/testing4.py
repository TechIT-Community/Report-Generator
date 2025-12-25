import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path

DOC_PATH = Path.cwd() / "testing1" / "test15.docx"

# Start Word
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
doc.Content.Delete()

# Write 15 pages with section breaks after 3, 7, 10
cursor = doc.Range(0, 0)
for page in range(1, 16):
    cursor.Collapse(c.wdCollapseEnd)
    cursor.InsertAfter(f"Page {page}: filler text. " * 100)
    cursor.InsertParagraphAfter()

    if page in (3, 7, 10):
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdSectionBreakNextPage)
    elif page < 15:
        cursor.Collapse(c.wdCollapseEnd)
        cursor.InsertBreak(c.wdPageBreak)

# Configure page numbering for each section
for idx, sec in enumerate(doc.Sections, start=1):
    # Break link to previous
    for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
        sec.Footers(hf_type).LinkToPrevious = False

    # Section 1 – no page numbers
    if idx == 1:
        for hf_type in [c.wdHeaderFooterPrimary, c.wdHeaderFooterFirstPage]:
            sec.Footers(hf_type).Range.Text = ""
        continue

    # Section 2: restart at 1, show page number on first page
    if idx == 2:
        sec.PageSetup.DifferentFirstPageHeaderFooter = False
        footer = sec.Footers(c.wdHeaderFooterPrimary)
        pnums = footer.PageNumbers
        pnums.RestartNumberingAtSection = True
        pnums.StartingNumber = 1
        pnums.Add(c.wdAlignParagraphCenter, True)

    # Sections 3 & 4: continue numbering, but hide number on first page
    if idx >= 3:
        sec.PageSetup.DifferentFirstPageHeaderFooter = True

        # Primary footer: used on 2nd page onward
        pfooter = sec.Footers(c.wdHeaderFooterPrimary)
        ppnums = pfooter.PageNumbers
        ppnums.RestartNumberingAtSection = False
        ppnums.Add(c.wdAlignParagraphCenter, False)

        # First-page footer: hide the number
        ffooter = sec.Footers(c.wdHeaderFooterFirstPage)
        ffooter.Range.Text = ""  # blank

# Save & quit
doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
doc.Close(False)
word.Quit()

print("✅ Created:", DOC_PATH.resolve())
