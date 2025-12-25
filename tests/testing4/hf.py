import win32com.client as win32
from win32com.client import constants as c
from pathlib import Path


class WordDocGenerator:
    def __init__(self, doc_path):
        self.doc_path = Path(doc_path)
        self.word = win32.gencache.EnsureDispatch("Word.Application")
        self.word.Visible = True
        self.doc = self.word.Documents.Add()
        self.doc.Content.Delete()

        self._generate_pages()
        self._configure_page_numbers()
        self._insert_initial_bookmark("PROJECT_TITLE")
        self.save()

    def _generate_pages(self):
        cursor = self.doc.Range(0, 0)
        blank_pages = {4, 8, 11}
        for page in range(1, 16):
            cursor.Collapse(c.wdCollapseEnd)
            if page not in blank_pages:
                cursor.InsertAfter(f"Page {page}: filler text. " * 100)
            cursor.InsertParagraphAfter()
            if page in (3, 7, 10):
                cursor.Collapse(c.wdCollapseEnd)
                cursor.InsertBreak(c.wdSectionBreakNextPage)
            elif page < 15:
                cursor.Collapse(c.wdCollapseEnd)
                cursor.InsertBreak(c.wdPageBreak)

    def _configure_page_numbers(self):
        for idx, sec in enumerate(self.doc.Sections, start=1):
            footer = sec.Footers(c.wdHeaderFooterPrimary)
            footer.LinkToPrevious = False
            if idx == 1:
                footer.Range.Text = ""
            else:
                pnums = footer.PageNumbers
                if idx == 2:
                    pnums.RestartNumberingAtSection = True
                    pnums.StartingNumber = 1
                    pnums.Add(c.wdAlignParagraphCenter, True)
                pnums.Add(c.wdAlignParagraphCenter, False)

    def _insert_initial_bookmark(self, name):
        rng = self.doc.Range(0, 0)
        rng.InsertBefore("___")
        bm_range = rng.Duplicate
        self.doc.Bookmarks.Add(name, bm_range)

    def replace_bookmarks(self, data_dict):
        def normalize(name):
            return name.strip().replace(" ", "_").upper()

        # Step 1: Replace any bookmarks that match keys in data_dict
        normalized_bookmarks = {normalize(b.Name): b.Name for b in self.doc.Bookmarks}

        for key, value in data_dict.items():
            norm_key = normalize(key)
            if norm_key in normalized_bookmarks:
                bm_name = normalized_bookmarks[norm_key]
                bm = self.doc.Bookmarks(bm_name)
                bm_range = bm.Range
                bm_range.Text = value + "\n\n"
                self.doc.Bookmarks.Add(bm_name, bm_range)  # re-insert the bookmark
                print(f"✅ Replaced bookmark '{bm_name}' with: {value}")
            else:
                print(f"⚠️ Bookmark '{key}' not found in document.")

        # Step 2: Special handling if 'Project Title' or 'Year' is present
        title = data_dict.get("Project Title")
        year = data_dict.get("Year")

        if title or year:
            for idx, section in enumerate(self.doc.Sections, start=1):
                if idx == 1:
                    continue

                # HEADER: Left-align project title
                header = section.Headers(c.wdHeaderFooterPrimary)
                header.LinkToPrevious = False
                if title:
                    header.Range.Text = title
                    header.Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

                # FOOTER: Left = dept, Center = year, Right = page number
                footer = section.Footers(c.wdHeaderFooterPrimary)
                footer.LinkToPrevious = False
                rng = footer.Range
                rng.Text = ""

                table = rng.Tables.Add(rng, NumRows=1, NumColumns=3)
                table.PreferredWidthType = c.wdPreferredWidthPercent
                table.PreferredWidth = 100
                table.Borders.Enable = False

                # Left = Dept.
                table.Cell(1, 1).Range.Text = "Dept. of CSE, BNMIT"
                table.Cell(1, 1).Range.ParagraphFormat.Alignment = c.wdAlignParagraphLeft

                # Center = Year (only if provided)
                if year:
                    table.Cell(1, 2).Range.Text = year
                table.Cell(1, 2).Range.ParagraphFormat.Alignment = c.wdAlignParagraphCenter

                # Right = Page number
                right_range = table.Cell(1, 3).Range
                right_range.Collapse(c.wdCollapseStart)
                right_range.Fields.Add(right_range, c.wdFieldPage)
                right_range.ParagraphFormat.Alignment = c.wdAlignParagraphRight

    def save(self):
        self.doc.SaveAs(str(self.doc_path), FileFormat=c.wdFormatDocumentDefault)
