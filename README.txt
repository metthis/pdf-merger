PDFMerger

Asks you for PDFs. Let's your specify which parts of which of the selected files and in which order you want to marge into a new PDF. Allows you to specify metadata and save the new PDF.

Consists of these files:
- pdf_merger.py
    - Defines the PDFMerger class which contains virtually all the code.
- script.py
    - Creates and instance of PDFMerger and runs the programme.

Libraries used:
- tkinter (including ttk and filedialog) for the UI
- pypdf (specifically PdfWriter) for the merging
- os to convert between file paths and the files' base names
- collections to create a namedtuple for handling metadata input