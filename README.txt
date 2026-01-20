The ./PDF SPLIT.bat file should install and launch the application once installed. If already installed, it will check dependencies and launch the application.



Mode 1: Explicit Range Splitter (Excel)

This tab splits a single PDF according to explicit page ranges and filenames listed in an Excel sheet.
Input: The ranges and filenames are read from the Excel file specified in the Excel Data field (e.g., ./manual_by_excel/data.xlsx).

Example:
Column B (Range)	Column C (Filename)				Result
1-3			1570000	Pages 1 through 3 are saved as  	1570000.pdf.
4-8			1570004	Pages 4 through 8 are saved as  	1570004.pdf.
9-10			1570045	Pages 9 through 10 are saved as 	1570045.pdf.



Mode 2: OCR PO Splitter (Excel)

This tab is intended to compare the POs (Purchase Orders) listed in Column A of the Excel sheet (e.g., ./OCR_by_excel/data.xlsx) with all letters and numbers picked up by OCR (Optical Character Recognition) from images in scanned documents, and split the PDF according to those matches.
Logic:

If a PO match is made on a page, that page is assigned to that PO group.
Pages occurring directly after a match with no new PO found are assumed to be part of the PO group from the previous page (e.g., a packing slip or a second page of a document).
This pattern continues until a new PO match is found.

Any page that falls out of this pattern (e.g., the first page of the PDF with no PO found) is assigned and saved as PO_unassigned.pdf.

Caveats: Problems can occur when a PO is not completely readable, is on the second page, or if multiple POs are listed on a single page (like in a comment or master bill of lading).



Mode 3: Sequential Splitter (Embedded Table)

This tab splits a single PDF sequentially based on the document lengths (page counts) entered directly into the embedded table.
Input: Data is entered directly into the table columns: File Name and Last Page. The value in the Last Page column is treated as the number of pages (document length) for that file.
Logic:

The first document starts on page 1 and runs for the number of pages specified in its Last Page column.
The next document starts on the page immediately following the previous document's end page.

Example:
File Name	Last Page (Length)				   Result
1200000	3	Pages 1 through 3 are saved as 			   1200000.pdf.
1210000	4	Pages 4 through 7 (4 pages total) are saved as     1210000.pdf.
1210001	2	Pages 8 through 9 (2 pages total) are saved as     1210001.pdf.

This mode also supports Excel-like features including arrow key navigation, Ctrl+C/Ctrl+V (Copy/Paste), Delete to clear cells, and Ctrl+Z/Ctrl+Y (Undo/Redo).