#Imports
import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path

#PATH
DOC_PATH = Path.cwd() / "testing1" / "test1.docx"
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
doc.Content.Delete()


# 2️⃣ Add 4 sections with real text (section break after each page)
for page in range(1, 5):
    rng = doc.Range()
    rng.Collapse(c.wdCollapseEnd)
    
    rng.InsertAfter(f"Random text filling page {page}. " * 150)
    rng.InsertParagraphAfter()
    
    if page < 4:
        rng.Collapse(c.wdCollapseEnd)
        rng.InsertBreak(c.wdSectionBreakNextPage)

# 3️⃣ Add page numbers to all sections
for sec in doc.Sections:
    footer = sec.Footers(c.wdHeaderFooterPrimary)
    footer.LinkToPrevious = False
    footer.Range.Text = "Page "
    rng = footer.Range
    rng.Collapse(c.wdCollapseEnd)
    rng.Fields.Add(rng, c.wdFieldPage)

# 4️⃣ Add border ONLY to section 2 (i.e., page 2)
sec3 = doc.Sections(3)
borders = sec3.Borders
borders.DistanceFromTop = borders.DistanceFromBottom = 24
borders.DistanceFromLeft = borders.DistanceFromRight = 12

for side in (c.wdBorderTop, c.wdBorderLeft, c.wdBorderBottom, c.wdBorderRight):
    br = borders(side)
    br.LineStyle = c.wdLineStyleSingle
    br.LineWidth = c.wdLineWidth150pt
    br.Color = c.wdColorAutomatic


# 5️⃣ Save and close
doc.SaveAs(str(DOC_PATH), FileFormat=c.wdFormatDocumentDefault)
doc.Close(False)
word.Quit()

print("✅ Created correctly at:", DOC_PATH.resolve())
