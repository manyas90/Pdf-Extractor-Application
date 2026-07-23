import os
import fitz


class PDFInfo:

    def __init__(self, file_path):

        self.file_name = os.path.basename(file_path)

        self.file_size = os.path.getsize(file_path)

        doc = fitz.open(file_path)

        self.total_pages = len(doc)

        doc.close()