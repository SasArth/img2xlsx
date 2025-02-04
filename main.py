from img2table.ocr import PaddleOCR
from img2table.document import Image

# Instantiation of OCR
ocr = PaddleOCR(lang="en")

# Instantiation of document, either an image or a PDF
doc = Image("Reddatabooktable.JPG")

# Extraction of tables and creation of a xlsx file containing tables
doc.to_xlsx(dest="here.xlsx",
            ocr=ocr,
            implicit_rows=False,
            implicit_columns=False,
            borderless_tables=False,
            min_confidence=50)