import sys
import os
import traceback
import pandas as pd

# PDF Libraries
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path 

# GUI Libraries
import PySide6.QtWidgets as qw
import PySide6.QtCore as qc
from PySide6.QtGui import QAction
import PySide6.QtGui as qg 

# ==========================================
# CONFIGURATION
# ==========================================
# Update these paths to match your local installation
try:
    pytesseract.pytesseract.tesseract_cmd = r'.\installed\tesseract\tesseract.exe'
except Exception:
    pass # Ignore if tesseract is not installed and OCR tab isn't used
POPPLER_PATH = r".\installed\poppler\Release-25.07.0-0\poppler-25.07.0\Library\bin"

# =====================================================================
# WORKER THREAD 1: Explicit Range Splitter (pypdf)
# =====================================================================
class RangeSplitterThread(qc.QThread):
    progress_signal = qc.Signal(str)
    finished_signal = qc.Signal(str)
    error_signal = qc.Signal(str)

    def __init__(self, pdf_path, excel_path, output_dir):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path
        self.output_dir = output_dir
        self._is_running = True

    def run(self):
        try:
            self.progress_signal.emit(f"--- Loading Excel: {self.excel_path} ---")
            
            # Logic from Script 1
            df_raw = pd.read_excel(self.excel_path, header=None)
            
            # Process Excel Data
            # Assuming Column 2 is Range (0-based index 2) and Column 3 is Filename (0-based index 3)
            page_ranges_split = df_raw.iloc[:, 2].astype(str).str.split('-', expand=True)
            df_raw['StartPage'] = pd.to_numeric(page_ranges_split[0], errors='coerce').astype('Int64')
            df_raw['EndPage'] = pd.to_numeric(page_ranges_split[1], errors='coerce').astype('Int64')
            df_raw['FileName'] = df_raw.iloc[:, 3]
            df_final = df_raw[['FileName', 'StartPage', 'EndPage']].dropna()

            self.split_pdf(df_final)
            self.finished_signal.emit("Range Splitting Complete!")

        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}\n{traceback.format_exc()}")

    def split_pdf(self, df):
        reader = PdfReader(self.pdf_path)
        total_pages = len(reader.pages)
        self.progress_signal.emit(f"--- Loaded PDF: {self.pdf_path} ({total_pages} pages) ---")

        os.makedirs(self.output_dir, exist_ok=True)

        for index, row in df.iterrows():
            if not self._is_running: return

            file_name = str(row['FileName'])
            start_page = row['StartPage']
            end_page = row['EndPage']

            self.progress_signal.emit(f"Processing: {file_name} | Range: {start_page}-{end_page}")

            try:
                writer = PdfWriter()
                # Convert 1-based Excel to 0-based Python indices
                start = int(start_page) - 1
                end = int(end_page)

                # Boundary Check
                if start < 0 or end > total_pages:
                    self.progress_signal.emit(f"  [SKIP] Range {start_page}-{end_page} exceeds PDF limits.")
                    continue

                for page_num in range(start, end):
                    writer.add_page(reader.pages[page_num])

                output_filename = os.path.join(self.output_dir, f"{file_name}.pdf")
                with open(output_filename, "wb") as f:
                    writer.write(f)
                self.progress_signal.emit(f"  -> Created: {output_filename}")

            except Exception as e:
                self.progress_signal.emit(f"  [ERROR] Row {index}: {str(e)}")

    def stop(self):
        self._is_running = False

# =====================================================================
# WORKER THREAD 2: OCR PO Splitter (fitz/tesseract)
# =====================================================================
class OCRSplitterThread(qc.QThread):
    progress_signal = qc.Signal(str)
    finished_signal = qc.Signal(str)
    error_signal = qc.Signal(str)

    def __init__(self, pdf_path, excel_path, output_dir):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path
        self.output_dir = output_dir
        self._is_running = True

    def run(self):
        try:
            self.process_pdf_splitting()
            self.finished_signal.emit("OCR Processing Complete!")
        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}\n{traceback.format_exc()}")

    def emit_progress(self, message):
        self.progress_signal.emit(message)

    def create_high_res_pdf(self, pdf_path, output_pdf_path, dpi_multiplier=1):

        self.emit_progress(f" -> Converting PDF pages to high-res images...")
        base_dpi = 300 # Reduced from 1000 for performance, usually 300 is enough for OCR
        target_dpi = base_dpi * dpi_multiplier
        
        pages = convert_from_path(pdf_path, dpi=target_dpi, poppler_path=POPPLER_PATH)
        
        if pages:
            first_page = pages[0]
            other_pages = pages[1:]
            first_page.save(output_pdf_path, save_all=True, append_images=other_pages, dpi=(target_dpi, target_dpi))
            self.emit_progress(f" -> Temporary high-res PDF created: {output_pdf_path}")
        else:
            raise RuntimeError("No pages found in the PDF during upscaling.")

    def process_pdf_splitting(self):
        self.emit_progress(f"Loading PO list from {self.excel_path}")
        df = pd.read_excel(self.excel_path)
        po_list_flat = df.values.flatten().tolist()
        target_pos = [str(po) for po in po_list_flat if pd.notna(po)]

        temp_high_res_path = "temp_analysis_high_res.pdf"
        os.makedirs(self.output_dir, exist_ok=True)

        # Step 1: High Res
        self.emit_progress("STEP 1: Creating High-Res Copy for Analysis...")
        self.create_high_res_pdf(self.pdf_path, temp_high_res_path)

        # Step 2: Scan
        self.emit_progress("STEP 2: Scanning Text & Grouping Pages...")
        doc_analysis = fitz.open(temp_high_res_path)
        po_page_map = {po: [] for po in target_pos}
        current_active_po = None

        for page_num, page in enumerate(doc_analysis):
            if not self._is_running: break
            self.emit_progress(f"Scanning Page {page_num + 1}...")
            
            word_list = page.get_text("words") # (x0, y0, x1, y1, "word", ...)
            found_new_po = False
            Y_TOLERANCE = 10

            for po in target_pos:
                if len(po) < 3: continue # Skip short noise

                for word_info in word_list:
                    word_text = word_info[4]
                    if word_text == po:
                        po_y = word_info[1]
                        is_disregarded = False
                        
                        # Check for "riding" on same line
                        for other in word_list:
                            if other[4].lower() == "riding" and abs(other[1] - po_y) < Y_TOLERANCE:
                                is_disregarded = True
                                self.emit_progress(f" -> Found 'riding' near {po}. Ignoring.")
                                break
                        
                        if not is_disregarded:
                            current_active_po = po
                            found_new_po = True
                            self.emit_progress(f" -> Found PO: {po}")
                            break
                if found_new_po: break

            if current_active_po and current_active_po in po_page_map:
                po_page_map[current_active_po].append(page_num)
            else:
                self.emit_progress(" -> Page unassigned (no PO found yet).")

        doc_analysis.close()

        # Step 3: Split Original
        if self._is_running:
            self.emit_progress("STEP 3: Splitting Original File...")
            doc_original = fitz.open(self.pdf_path)
            
            for po_name, page_numbers in po_page_map.items():
                if not page_numbers: continue
                
                new_doc = fitz.open()
                for p_num in page_numbers:
                    new_doc.insert_pdf(doc_original, from_page=p_num, to_page=p_num)
                
                out_name = os.path.join(self.output_dir, f"PO_{po_name}.pdf")
                new_doc.save(out_name)
                new_doc.close()
                self.emit_progress(f"Saved: {out_name}")
            
            doc_original.close()

        # Cleanup
        if os.path.exists(temp_high_res_path):
            os.remove(temp_high_res_path)

    def stop(self):
        self._is_running = False

# =====================================================================
# WORKER THREAD 3: Sequential Splitter (pypdf)
# =====================================================================
class SequentialSplitterThread(qc.QThread):
    progress_signal = qc.Signal(str)
    finished_signal = qc.Signal(str)
    error_signal = qc.Signal(str)

    # split_data is a list of dicts: [{'FileName': '1200000', 'LastPage': 3}, ...]
    def __init__(self, pdf_path, split_data, output_dir):
        super().__init__()
        self.pdf_path = pdf_path
        self.split_data = split_data
        self.output_dir = output_dir
        self._is_running = True

    def run(self):
        print("\n--- DEBUG THREAD: run() started ---")
        try:
            self.split_pdf()
            self.finished_signal.emit("Sequential Splitting Complete!")
            print("--- DEBUG THREAD: run() finished successfully ---")

        except Exception as e:
            error_msg = f"Error: {str(e)}\n{traceback.format_exc()}"
            self.error_signal.emit(error_msg)
            print(f"--- DEBUG THREAD: CRITICAL ERROR ---\n{error_msg}")

    def split_pdf(self):
        print("--- DEBUG THREAD: split_pdf() started ---")
        reader = PdfReader(self.pdf_path)
        total_pages = len(reader.pages)
        self.progress_signal.emit(f"--- Loaded PDF: {self.pdf_path} ({total_pages} pages) ---")
        
        os.makedirs(self.output_dir, exist_ok=True)
        
        current_start_page = 1 # 1-based logic

        for index, row in enumerate(self.split_data):
            if not self._is_running: return

            file_name = str(row['FileName'])
            doc_length = row['LastPage'] 
            start_page = current_start_page
            calculated_end_page = start_page + doc_length - 1 
            end_page = calculated_end_page 

            print(f"  [DEBUG THREAD] Processing Row {index}: File={file_name}, Range={start_page}-{end_page}")

            if doc_length <= 0:
                self.progress_signal.emit(f"  [SKIP] Row {index+1}: Document length is {doc_length}. Skipping row.")
                continue

            if index == len(self.split_data) - 1 and end_page > total_pages:
                 end_page = total_pages
                 
            if end_page > total_pages:
                 self.progress_signal.emit(f"  [SKIP] Row {index+1}: End page {end_page} exceeds total PDF pages ({total_pages}). Stopping.")
                 break

            self.progress_signal.emit(f"Processing: {file_name} | Range: {start_page}-{end_page}")

            try:
                writer = PdfWriter()
                start = int(start_page) - 1
                end = int(end_page)

                for page_num in range(start, end):
                    writer.add_page(reader.pages[page_num])

                output_filename = os.path.join(self.output_dir, f"{file_name}.pdf")
                
                print(f"  [DEBUG THREAD] Attempting to write file: {output_filename}")
                
                with open(output_filename, "wb") as f:
                    writer.write(f)
                    
                print(f"  [DEBUG THREAD] File written successfully.")
                self.progress_signal.emit(f"  -> Created: {output_filename}")

            except Exception as e:
                error_msg = f"  [ERROR] Row {index+1}: {str(e)}\n{traceback.format_exc()}"
                self.progress_signal.emit(error_msg)
                print(f"  [DEBUG THREAD] Inner loop error:\n{error_msg}")

            current_start_page = end_page + 1 
        
        print("--- DEBUG THREAD: split_pdf() finished loop ---")

    def stop(self):
        self._is_running = False


# =====================================================================
# GUI TABS
# =====================================================================

class BaseTab(qw.QWidget):
    """Helper class to reduce code duplication for file selection"""
    def __init__(self, default_pdf, default_excel, log_placeholder, include_excel=True):
        super().__init__()
        self.layout = qw.QVBoxLayout()
        
        # PDF Input
        self.pdf_input = qw.QLineEdit(default_pdf)
        self.btn_pdf = qw.QPushButton("Browse PDF")
        self.btn_pdf.clicked.connect(lambda: self.browse_file(self.pdf_input, "PDF Files (*.pdf)"))
        
        h1 = qw.QHBoxLayout()
        h1.addWidget(qw.QLabel("PDF File:"))
        h1.addWidget(self.pdf_input)
        h1.addWidget(self.btn_pdf)
        self.layout.addLayout(h1)

        # Excel Input (Only if needed)
        if include_excel:
            self.excel_input = qw.QLineEdit(default_excel)
            self.btn_excel = qw.QPushButton("Browse Excel")
            self.btn_excel.clicked.connect(lambda: self.browse_file(self.excel_input, "Excel Files (*.xlsx *.xls)"))

            h2 = qw.QHBoxLayout()
            h2.addWidget(qw.QLabel("Excel Data:"))
            h2.addWidget(self.excel_input)
            h2.addWidget(self.btn_excel)
            self.layout.addLayout(h2)
        else:
            self.excel_input = None # Placeholder to prevent errors in subclasses

        # Output Folder
        self.out_input = qw.QLineEdit(os.path.abspath("/split_saved_docs"))
        self.btn_out = qw.QPushButton("Browse Folder")
        self.btn_out.clicked.connect(self.browse_folder)

        h3 = qw.QHBoxLayout()
        h3.addWidget(qw.QLabel("Output Dir:"))
        h3.addWidget(self.out_input)
        h3.addWidget(self.btn_out)
        self.layout.addLayout(h3)

        # Start Button
        self.btn_start = qw.QPushButton("Start Processing")
        self.btn_start.setStyleSheet("background-color: lightgreen; font-weight: bold; padding: 5px;")
        self.layout.addWidget(self.btn_start)

        # Log
        self.log_area = qw.QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setPlaceholderText(log_placeholder)
        self.layout.addWidget(self.log_area)

        self.setLayout(self.layout)

    def browse_file(self, line_edit, filter_str):
        fname, _ = qw.QFileDialog.getOpenFileName(self, "Select File", "", filter_str)
        if fname: line_edit.setText(fname)

    def browse_folder(self):
        dname = qw.QFileDialog.getExistingDirectory(self, "Select Output Folder", self.out_input.text())
        if dname: self.out_input.setText(dname)

    def log(self, msg):
        self.log_area.append(msg)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())

class RangeSplitterTab(BaseTab):
    def __init__(self):
        super().__init__("manual_by_excel/source_docs/example.pdf", "manual_by_excel/data.xlsx", "Logs for Range Splitter (Page 1-5, etc)...", include_excel=True)
        self.btn_start.clicked.connect(self.start_process)
        self.thread = None

    def start_process(self):
        pdf = self.pdf_input.text()
        xls = self.excel_input.text()
        out = self.out_input.text()

        if not os.path.exists(pdf) or not os.path.exists(xls):
            qw.QMessageBox.warning(self, "Error", "Input files not found.")
            return

        self.log_area.clear()
        self.btn_start.setEnabled(False)
        
        self.thread = RangeSplitterThread(pdf, xls, out)
        self.thread.progress_signal.connect(self.log)
        self.thread.finished_signal.connect(self.on_finish)
        self.thread.error_signal.connect(self.on_error)
        self.thread.start()

    def on_finish(self, msg):
        self.log(msg)
        self.btn_start.setEnabled(True)
        qw.QMessageBox.information(self, "Done", msg)

    def on_error(self, msg):
        self.log(f"CRITICAL ERROR: {msg}")
        self.btn_start.setEnabled(True)
        qw.QMessageBox.critical(self, "Error", "An error occurred. See logs.")

class OCRSplitterTab(BaseTab):
    def __init__(self):
        super().__init__("OCR_by_excel/source_docs/example.pdf", "OCR_by_excel/data.xlsx", "Logs for OCR PO Splitter...", include_excel=True)
        self.btn_start.clicked.connect(self.start_process)
        self.thread = None

    def start_process(self):
        pdf = self.pdf_input.text()
        xls = self.excel_input.text()
        out = self.out_input.text()

        if not os.path.exists(pdf) or not os.path.exists(xls):
            qw.QMessageBox.warning(self, "Error", "Input files not found.")
            return

        self.log_area.clear()
        self.btn_start.setEnabled(False)
        
        self.thread = OCRSplitterThread(pdf, xls, out)
        self.thread.progress_signal.connect(self.log)
        self.thread.finished_signal.connect(self.on_finish)
        self.thread.error_signal.connect(self.on_error)
        self.thread.start()

    def on_finish(self, msg):
        self.log(msg)
        self.btn_start.setEnabled(True)
        qw.QMessageBox.information(self, "Done", msg)

    def on_error(self, msg):
        self.log(f"CRITICAL ERROR: {msg}")
        self.btn_start.setEnabled(True)
        qw.QMessageBox.critical(self, "Error", "An error occurred. See logs.")


# =====================================================================
# GUI COMPONENT: Custom QTableWidget for Auto-Row Addition (NEW)
# =====================================================================

class SetCellTextCommand(qg.QUndoCommand):
    """Undo/Redo command for setting the text of a single cell."""
    def __init__(self, table, row, col, new_text, old_text, description="Set Cell Text", parent=None):
        super().__init__(description, parent)
        self.table = table
        self.row = row
        self.col = col
        self.new_text = new_text
        self.old_text = old_text
        self.index = table.model().index(row, col)
        print(f"[DEBUG COMMAND] Command created for R{row}, C{col}. Old='{old_text}', New='{new_text}'")

    def _set_text_via_model(self, text):
        model = self.table.model()
        
        # CRUCIAL: If the item doesn't exist, we must create it before setting data.
        if self.table.item(self.row, self.col) is None:
            print(f"[DEBUG COMMAND] Item R{self.row}, C{self.col} was None. Creating item.")
            self.table.setItem(self.row, self.col, qw.QTableWidgetItem())
            
        # Now, set the data via the model
        model.setData(self.index, text, qc.Qt.ItemDataRole.EditRole)
        print(f"[DEBUG COMMAND] Model setData called for R{self.row}, C{self.col} with text: '{text}'")

    def redo(self):
        self._set_text_via_model(self.new_text)

    def undo(self):
        self._set_text_via_model(self.old_text)





class ArrowKeyCommitDelegate(qw.QStyledItemDelegate):
    """
    A delegate that installs an event filter on the cell editor 
    to intercept arrow keys and force commit/navigation.
    """
    def __init__(self, parent=None):
        super().__init__(parent)

    def createEditor(self, parent, option, index):
        # Create the default editor (usually a QLineEdit)
        editor = super().createEditor(parent, option, index)
        # Install the event filter on the editor
        editor.installEventFilter(self)
        return editor

    def eventFilter(self, source, event):
        if event.type() == qc.QEvent.KeyPress:
            key = event.key()
            
            # Keys that should commit data and navigate
            is_navigation_key = key in (
                qc.Qt.Key.Key_Down, 
                qc.Qt.Key.Key_Up, 
                qc.Qt.Key.Key_Left, 
                qc.Qt.Key.Key_Right
            )

            if is_navigation_key:
                table = self.parent() # The parent is the QTableWidget
                
                # A. Force commit data by sending a Return key event to the editor
                return_event = qg.QKeyEvent(
                    qc.QEvent.KeyPress, 
                    qc.Qt.Key.Key_Return, 
                    qc.Qt.KeyboardModifier.NoModifier
                )
                qw.QApplication.sendEvent(source, return_event)
                
                # B. Manually calculate the new index and move the cursor
                # --- FIX: Removed 'MoveOperation' prefix ---
                if key == qc.Qt.Key.Key_Right:
                    cursor_move = qw.QAbstractItemView.MoveNext
                elif key == qc.Qt.Key.Key_Left:
                    cursor_move = qw.QAbstractItemView.MovePrevious
                elif key == qc.Qt.Key.Key_Down:
                    cursor_move = qw.QAbstractItemView.MoveDown
                elif key == qc.Qt.Key.Key_Up:
                    cursor_move = qw.QAbstractItemView.MoveUp
                else:
                    cursor_move = qw.QAbstractItemView.NoMove
                # -------------------------------------------

                new_index = table.moveCursor(cursor_move, event.modifiers())
                table.setCurrentIndex(new_index)
                
                # Start editing the new cell immediately
                if new_index.isValid():
                    table.edit(new_index)
                
                # C. Accept the original arrow key event
                return True # Event handled

        return super().eventFilter(source, event)


class AutoAddRowTable(qw.QTableWidget):
    """
    A QTableWidget subclass that handles Excel-like copy/paste, 
    delete, and undo/redo functionality.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        # --- NEW: Undo Stack Initialization ---
        self.undo_stack = qg.QUndoStack(self)
        # --------------------------------------
        
        # Install the custom delegate
        self.setItemDelegate(ArrowKeyCommitDelegate(self)) 
        
        # --- Setup Copy/Paste/Undo/Redo Actions ---
        self.copy_action = qg.QAction("Copy", self)
        self.copy_action.setShortcut(qg.QKeySequence.StandardKey.Copy)
        self.copy_action.triggered.connect(self.copy_selection_to_clipboard)
        self.addAction(self.copy_action)
        
        self.paste_action = qg.QAction("Paste", self)
        self.paste_action.setShortcut(qg.QKeySequence.StandardKey.Paste)
        self.paste_action.triggered.connect(self.paste_from_clipboard)
        self.addAction(self.paste_action)
        
        # NEW: Undo/Redo Actions
        self.undo_action = self.undo_stack.createUndoAction(self, "Undo")
        self.undo_action.setShortcut(qg.QKeySequence.StandardKey.Undo)
        self.addAction(self.undo_action)
        
        self.redo_action = self.undo_stack.createRedoAction(self, "Redo")
        self.redo_action.setShortcut(qg.QKeySequence.StandardKey.Redo)
        self.addAction(self.redo_action)
        # ------------------------------------------

    # --- NEW: Override setData to push the command (for all edits) ---
    def setData(self, index, value, role=qc.Qt.ItemDataRole.EditRole):
        if role == qc.Qt.ItemDataRole.EditRole:
            # 1. Ensure the item exists before getting old_text
            item = self.item(index.row(), index.column())
            if item is None:
                item = qw.QTableWidgetItem()
                self.setItem(index.row(), index.column(), item)
            
            old_text = item.text()
            new_text = str(value)
            
            print(f"[DEBUG TABLE] setData called for R{index.row()}, C{index.column()}. Old='{old_text}', New='{new_text}'")
            
            if new_text != old_text:
                # 2. Create and push the command
                command = SetCellTextCommand(
                    self, 
                    index.row(), 
                    index.column(), 
                    new_text, 
                    old_text,
                    description=f"Set {index.row()},{index.column()}"
                )
                self.undo_stack.push(command)
                
                # 3. Return True to indicate the data was handled by the command's redo.
                return True
        
        # For all other roles, or if text didn't change, call the base implementation
        return super().setData(index, value, role)
    # --------------------------------------------------------------------


    def add_row(self):
        """Adds a new row to the table."""
        row_count = self.rowCount()
        self.setRowCount(row_count + 1)
        return row_count # Return the index of the new row

    def copy_selection_to_clipboard(self):
        # ... (Copy logic remains the same) ...
        selected_indexes = self.selectedIndexes()
        if not selected_indexes:
            return
            
        selected_indexes.sort(key=lambda index: (index.row(), index.column()))
        
        output_data = []
        current_row = -1
        row_data = []
        
        for index in selected_indexes:
            item = self.item(index.row(), index.column())
            text = item.text() if item else ""
            
            if index.row() != current_row:
                if row_data:
                    output_data.append(row_data)
                row_data = []
                current_row = index.row()
            
            row_data.append(text)

        if row_data:
            output_data.append(row_data)

        text_to_copy = ""
        for row in output_data:
            text_to_copy += "\t".join(row) + "\n"
            
        qw.QApplication.clipboard().setText(text_to_copy.strip())

    def paste_from_clipboard(self):
        """Pastes data from the clipboard into the table starting at the current cell."""
        clipboard = qw.QApplication.clipboard()
        clipboard_text = clipboard.text()
        
        if not clipboard_text:
            return

        current_index = self.currentIndex()
        if not current_index.isValid():
            return

        start_row = current_index.row()
        start_col = current_index.column()
        
        rows = clipboard_text.strip().split('\n')
        parsed_data = [row.split('\t') for row in rows]
        
        paste_rows = len(parsed_data)
        paste_cols = max(len(col) for col in parsed_data) if parsed_data else 0
        
        # Ensure the table has enough rows
        required_rows = start_row + paste_rows
        if required_rows > self.rowCount():
            self.setRowCount(required_rows)

        # --- Use QUndoStack.beginMacro() and endMacro() ---
        self.undo_stack.beginMacro("Paste Data")
        
        for r_idx, row_data in enumerate(parsed_data):
            target_row = start_row + r_idx
            
            for c_idx, cell_data in enumerate(row_data):
                target_col = start_col + c_idx
                
                if target_col >= self.columnCount():
                    break
                
                # Get old text for the undo command
                item = self.item(target_row, target_col)
                old_text = item.text() if item else ""
                
                if cell_data != old_text:
                    command = SetCellTextCommand(
                        self, 
                        target_row, 
                        target_col, 
                        cell_data, 
                        old_text
                    )
                    # Push the sub-command directly to the stack while the macro is open
                    self.undo_stack.push(command)

        self.undo_stack.endMacro()
        # ----------------------------------------------------------

        # Ensure the last pasted cell is visible
        if parsed_data:
            last_row = start_row + paste_rows - 1
            last_col = min(start_col + paste_cols - 1, self.columnCount() - 1)
            self.scrollToItem(self.item(last_row, last_col))


    def keyPressEvent(self, event):
        """Overrides the key press event to handle auto-row addition and Delete."""
        
        # --- Delete Key Handler ---
        if event.key() == qc.Qt.Key.Key_Delete:
            selected_indexes = self.selectedIndexes()
            if not selected_indexes:
                return
            
            # Use a Macro Command for Undo/Redo
            self.undo_stack.beginMacro("Clear Cells")
            
            for index in selected_indexes:
                item = self.item(index.row(), index.column())
                if item and item.text():
                    command = SetCellTextCommand(
                        self, 
                        index.row(), 
                        index.column(), 
                        "", 
                        item.text()
                    )
                    self.undo_stack.push(command)
            
            self.undo_stack.endMacro()
            event.accept()
            return

        # 1. Check for Down Arrow key (Auto-row logic)
        if event.key() == qc.Qt.Key.Key_Down:
            current_row = self.currentRow()
            row_count = self.rowCount()
            
            # 2. Check if the cursor is in the last row
            if current_row == row_count - 1:
                # 3. Add a new row
                new_row_index = self.add_row()
                
                # 4. Move the cursor to the new row, keeping the same column
                new_index = self.model().index(new_row_index, self.currentColumn())
                self.setCurrentIndex(new_index)
                self.edit(new_index) # Start editing the new cell immediately
                
                event.accept()
                return

        # Handle all other key presses normally
        super().keyPressEvent(event)



    # --------------------------------------------------------------------


    def add_row(self):
        """Adds a new row to the table."""
        row_count = self.rowCount()
        self.setRowCount(row_count + 1)
        return row_count # Return the index of the new row

    def copy_selection_to_clipboard(self):
        # ... (Copy logic remains the same) ...
        selected_indexes = self.selectedIndexes()
        if not selected_indexes:
            return
            
        selected_indexes.sort(key=lambda index: (index.row(), index.column()))
        
        output_data = []
        current_row = -1
        row_data = []
        
        for index in selected_indexes:
            item = self.item(index.row(), index.column())
            text = item.text() if item else ""
            
            if index.row() != current_row:
                if row_data:
                    output_data.append(row_data)
                row_data = []
                current_row = index.row()
            
            row_data.append(text)

        if row_data:
            output_data.append(row_data)

        text_to_copy = ""
        for row in output_data:
            text_to_copy += "\t".join(row) + "\n"
            
        qw.QApplication.clipboard().setText(text_to_copy.strip())

    def paste_from_clipboard(self):
        """Pastes data from the clipboard into the table starting at the current cell."""
        clipboard = qw.QApplication.clipboard()
        clipboard_text = clipboard.text()
        
        if not clipboard_text:
            return

        current_index = self.currentIndex()
        if not current_index.isValid():
            return

        start_row = current_index.row()
        start_col = current_index.column()
        
        rows = clipboard_text.strip().split('\n')
        parsed_data = [row.split('\t') for row in rows]
        
        paste_rows = len(parsed_data)
        paste_cols = max(len(col) for col in parsed_data) if parsed_data else 0
        
        # Ensure the table has enough rows
        required_rows = start_row + paste_rows
        if required_rows > self.rowCount():
            self.setRowCount(required_rows)

        # --- FIX: Use QUndoStack.beginMacro() and endMacro() ---
        self.undo_stack.beginMacro("Paste Data")
        
        for r_idx, row_data in enumerate(parsed_data):
            target_row = start_row + r_idx
            
            for c_idx, cell_data in enumerate(row_data):
                target_col = start_col + c_idx
                
                if target_col >= self.columnCount():
                    break
                
                # Get old text for the undo command
                item = self.item(target_row, target_col)
                old_text = item.text() if item else ""
                
                if cell_data != old_text:
                    command = SetCellTextCommand(
                        self, 
                        target_row, 
                        target_col, 
                        cell_data, 
                        old_text
                    )
                    # Push the sub-command directly to the stack while the macro is open
                    self.undo_stack.push(command)

        self.undo_stack.endMacro()
        # ----------------------------------------------------------

        # Ensure the last pasted cell is visible
        if parsed_data:
            last_row = start_row + paste_rows - 1
            last_col = min(start_col + paste_cols - 1, self.columnCount() - 1)
            self.scrollToItem(self.item(last_row, last_col))



    def keyPressEvent(self, event):
        """Overrides the key press event to handle auto-row addition and Delete."""
        
        # --- NEW: Delete Key Handler (FIXED) ---
        if event.key() == qc.Qt.Key.Key_Delete:
            selected_indexes = self.selectedIndexes()
            if not selected_indexes:
                return
            
            # Use a Macro Command for Undo/Redo
            self.undo_stack.beginMacro("Clear Cells")
            
            for index in selected_indexes:
                item = self.item(index.row(), index.column())
                if item and item.text():
                    command = SetCellTextCommand(
                        self, 
                        index.row(), 
                        index.column(), 
                        "", 
                        item.text()
                    )
                    self.undo_stack.push(command)
            
            self.undo_stack.endMacro()
            event.accept()
            return

        # -------------------------------
        
        # 1. Check for Down Arrow key (Auto-row logic)
        if event.key() == qc.Qt.Key.Key_Down:
            current_row = self.currentRow()
            row_count = self.rowCount()
            
            # 2. Check if the cursor is in the last row
            if current_row == row_count - 1:
                # 3. Add a new row
                new_row_index = self.add_row()
                
                # 4. Move the cursor to the new row, keeping the same column
                new_index = self.model().index(new_row_index, self.currentColumn())
                self.setCurrentIndex(new_index)
                self.edit(new_index) # Start editing the new cell immediately
                
                event.accept()
                return

        # Handle all other key presses normally
        super().keyPressEvent(event)









# =====================================================================
# GUI COMPONENT: Embedded Spreadsheet (MODIFIED)
# =====================================================================

class ExcelLikeTable(qw.QWidget):
    """A QWidget embedding the editable AutoAddRowTable with add/remove row buttons."""
    def __init__(self):
        super().__init__()
        self.layout = qw.QVBoxLayout(self)
        
        # Use the custom AutoAddRowTable instead of QTableWidget
        self.table = AutoAddRowTable() 
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["File Name (e.g., 1200000)", "Last Page (e.g., 3)"])
        
        # Set default rows based on user's request
        self.table.setRowCount(2)
        self.table.setItem(0, 0, qw.QTableWidgetItem("1200000"))
        self.table.setItem(0, 1, qw.QTableWidgetItem("3"))
        self.table.setItem(1, 0, qw.QTableWidgetItem("1210000"))
        self.table.setItem(1, 1, qw.QTableWidgetItem("4"))

        # Stretch the columns
        self.table.horizontalHeader().setSectionResizeMode(0, qw.QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, qw.QHeaderView.ResizeMode.ResizeToContents)
        
        # Add/Remove Buttons
        self.btn_add = qw.QPushButton("Add Row")
        self.btn_remove = qw.QPushButton("Remove Selected Row(s)")
        
        # Connect button to the new table's add_row method
        self.btn_add.clicked.connect(self.table.add_row) 
        self.btn_remove.clicked.connect(self.remove_rows)
        
        h_buttons = qw.QHBoxLayout()
        h_buttons.addWidget(self.btn_add)
        h_buttons.addWidget(self.btn_remove)
        
        self.layout.addWidget(qw.QLabel("Define Sequential Split Points (File A gets pages 1-X, File B gets X+1-Y, etc):"))
        self.layout.addWidget(self.table)
        self.layout.addLayout(h_buttons)
        
    def remove_rows(self):
        selected_rows = sorted(list(set(index.row() for index in self.table.selectedIndexes())), reverse=True)
        for row in selected_rows:
            self.table.removeRow(row)

    def get_data(self):
        """Extracts and validates data from the QTableWidget by ensuring items exist."""
        data = []
        
        print("\n--- DEBUG: get_data() START (Forcing Item Creation) ---")
        
        for row in range(self.table.rowCount()):
            
            # --- CRUCIAL FIX: Ensure QTableWidgetItem exists for the cell ---
            # If the item is None, create it. This forces the QTableWidget to 
            # synchronize the item with any data set via the model (like from QUndoCommand).
            filename_item = self.table.item(row, 0)
            if filename_item is None:
                filename_item = qw.QTableWidgetItem()
                self.table.setItem(row, 0, filename_item)
                
            page_item = self.table.item(row, 1)
            if page_item is None:
                page_item = qw.QTableWidgetItem()
                self.table.setItem(row, 1, page_item)
            # ----------------------------------------------------------------
            
            filename = filename_item.text() if filename_item else ""
            page_str = page_item.text() if page_item else ""
            
            print(f"  [DEBUG] Row {row}: filename='{filename}', page_str='{page_str}'")
            
            try:
                page = int(page_str)
            except ValueError:
                page = None
            
            if filename.strip() and page is not None and page > 0:
                data.append({
                    'FileName': filename.strip(),
                    'LastPage': page
                })
        print(f"--- DEBUG: get_data() END. Final data count: {len(data)} ---\n")
        return data






class SequentialSplitterTab(qw.QWidget):
    """The new Tab 3 implementation using the embedded table."""
    def __init__(self):
        super().__init__()
        self.layout = qw.QVBoxLayout()
        self.thread = None
        default_pdf = "sequential_by_table/source_docs/example.pdf"
        default_out = os.path.abspath("sequential_by_table/split_saved_docs")

        # --- File Selection (PDF and Output Dir) ---
        self.pdf_input = qw.QLineEdit(default_pdf)
        self.btn_pdf = qw.QPushButton("Browse PDF")
        self.btn_pdf.clicked.connect(lambda: self.browse_file(self.pdf_input, "PDF Files (*.pdf)"))
        h1 = qw.QHBoxLayout()
        h1.addWidget(qw.QLabel("PDF File:"))
        h1.addWidget(self.pdf_input)
        h1.addWidget(self.btn_pdf)
        self.layout.addLayout(h1)

        self.out_input = qw.QLineEdit(default_out)
        self.btn_out = qw.QPushButton("Browse Folder")
        self.btn_out.clicked.connect(self.browse_folder)
        h2 = qw.QHBoxLayout()
        h2.addWidget(qw.QLabel("Output Dir:"))
        h2.addWidget(self.out_input)
        h2.addWidget(self.btn_out)
        self.layout.addLayout(h2)

        # --- Embedded Spreadsheet (The Key Component) ---
        self.excel_table_widget = ExcelLikeTable()
        self.layout.addWidget(self.excel_table_widget)

        # --- Start Button ---
        self.btn_start = qw.QPushButton("Start Sequential Splitting")
        self.btn_start.setStyleSheet("background-color: lightblue; font-weight: bold; padding: 5px;")
        self.btn_start.clicked.connect(self.start_process)
        self.layout.addWidget(self.btn_start)

        # --- Log ---
        self.log_area = qw.QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setPlaceholderText("Logs for Sequential Page Splitter...")
        self.layout.addWidget(self.log_area)

        self.setLayout(self.layout)

    def browse_file(self, line_edit, filter_str):
        fname, _ = qw.QFileDialog.getOpenFileName(self, "Select File", "", filter_str)
        if fname: line_edit.setText(fname)

    def browse_folder(self):
        dname = qw.QFileDialog.getExistingDirectory(self, "Select Output Folder", self.out_input.text())
        if dname: self.out_input.setText(dname)

    def log(self, msg):
        self.log_area.append(msg)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())

    def start_process(self):
        pdf = self.pdf_input.text()
        output_data = self.excel_table_widget.get_data()
        out = self.out_input.text()

        if not os.path.exists(pdf):
            qw.QMessageBox.warning(self, "Error", "PDF file not found.")
            return

        if not output_data:
            qw.QMessageBox.warning(self, "Error", "No valid split definitions entered in the table.")
            return

        # Sequential check: LastPage numbers must be strictly increasing
        #last_page_numbers = [item['LastPage'] for item in output_data]
        #if any(last_page_numbers[i] >= last_page_numbers[i+1] for i in range(len(last_page_numbers)-1)):
        #    qw.QMessageBox.critical(self, "Error", "Last Page numbers must be in strictly ascending order for Sequential Split (e.g., 3, 7, 10).")
        #    return

        self.log_area.clear()
        self.btn_start.setEnabled(False)
        
        self.thread = SequentialSplitterThread(pdf, output_data, out)
        self.thread.progress_signal.connect(self.log)
        self.thread.finished_signal.connect(self.on_finish)
        self.thread.error_signal.connect(self.on_error)
        self.thread.start()

    def on_finish(self, msg):
        self.log(msg)
        self.btn_start.setEnabled(True)
        qw.QMessageBox.information(self, "Done", msg)

    def on_error(self, msg):
        self.log(f"CRITICAL ERROR: {msg}")
        self.btn_start.setEnabled(True)
        qw.QMessageBox.critical(self, "Error", "An error occurred. See logs.")

# =====================================================================
# MAIN WINDOW
# =====================================================================
class MainWindow(qw.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Splitter Tool")
        self.setGeometry(100, 100, 700, 600)

        self.tabs = qw.QTabWidget()
        self.setCentralWidget(self.tabs)

        # Add Tabs
        self.tab1 = RangeSplitterTab()
        self.tab2 = OCRSplitterTab()
        self.tab3 = SequentialSplitterTab()

        self.tabs.addTab(self.tab1, "Mode 1: Explicit Range Splitter (Excel)")
        self.tabs.addTab(self.tab2, "Mode 2: OCR PO Splitter (Excel)")
        self.tabs.addTab(self.tab3, "Mode 3: Sequential Splitter (Embedded Table)")

if __name__ == "__main__":
    # Ensure a QApplication instance exists before creating widgets
    if not qw.QApplication.instance():
        app = qw.QApplication(sys.argv)
    else:
        app = qw.QApplication.instance()
    
    # Set style for better look
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())