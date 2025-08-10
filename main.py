import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import pikepdf
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import letter, A4, A5, legal
from reportlab.lib.units import inch
import io
import threading
import math
import os
import re
from collections import deque
import traceback
from pdf2image import convert_from_path

# --- BACKEND PDF PROCESSING LOGIC ---

class PdfProcessor:
    """Handles all the backend PDF manipulation tasks."""

    def __init__(self, file_paths, options, progress_callback, status_callback):
        self.file_paths = file_paths
        self.options = options
        self.progress_callback = progress_callback
        self.status_callback = status_callback
        self.paper_sizes = {
            "A4": A4,
            "A5": A5,
            "Letter": letter,
            "Legal": legal
        }

    def create_booklet(self):
        """Main method to orchestrate the booklet creation process."""
        try:
            self.status_callback("Starting booklet creation...")
            self.progress_callback(5)

            # 1. Merge PDFs if multiple are provided
            self.status_callback("Merging PDF files...")
            merged_pdf_path = self._merge_pdfs()
            if not merged_pdf_path:
                raise Exception("Failed to merge PDFs.")
            self.progress_callback(20)

            # 2. Open the merged PDF
            with pikepdf.open(merged_pdf_path) as source_pdf:
                # 3. Parse page range
                pages_to_process = self._parse_page_range(len(source_pdf.pages))
                if not pages_to_process:
                    raise ValueError("Invalid page range specified.")
                
                # 4. Handle splitting
                self.status_callback("Organizing pages for booklet layout...")
                booklet_chunks = self._split_into_booklets(pages_to_process)
                self.progress_callback(40)

                output_files = []
                total_chunks = len(booklet_chunks)

                for i, chunk in enumerate(booklet_chunks):
                    chunk_status = f"Processing booklet {i+1} of {total_chunks}..."
                    self.status_callback(chunk_status)
                    
                    # 5. Reorder pages for booklet layout
                    ordered_pages_indices = self._reorder_for_booklet(chunk)
                    
                    # 6. Create the final output PDF
                    output_pdf = self._create_final_pdf(source_pdf, ordered_pages_indices)
                    
                    # 7. Save the output
                    base_path, extension = os.path.splitext(self.options['output_path'])
                    output_filename = f"{base_path}_part_{i+1}{extension}" if total_chunks > 1 else self.options['output_path']
                    output_pdf.save(output_filename)
                    output_pdf.close()
                    output_files.append(output_filename)
                    
                    progress = 40 + int(60 * (i + 1) / total_chunks)
                    self.progress_callback(progress)

            self.status_callback(f"Booklet(s) created successfully: {', '.join(output_files)}")

        except Exception as e:
            print("Exception:", e)
            traceback.print_exc()
            self.status_callback(f"Error: {e}")
            return False
        finally:
            # Clean up temporary merged file
            if hasattr(self, '_temp_merged_file') and os.path.exists(self._temp_merged_file):
                os.remove(self._temp_merged_file)
        
        return True

    def _merge_pdfs(self):
        """Merges multiple PDFs into a single temporary file."""
        if not self.file_paths:
            return None
        if len(self.file_paths) == 1:
            # If only one file, check for password and return its path
            try:
                with pikepdf.open(self.file_paths[0], password=self.options.get('password', '')):
                    return self.file_paths[0]
            except pikepdf.PasswordError:
                raise Exception(f"Incorrect password for {os.path.basename(self.file_paths[0])}")

        merged_pdf = pikepdf.Pdf.new()
        for path in self.file_paths:
            try:
                with pikepdf.open(path, password=self.options.get('password', '')) as pdf:
                    merged_pdf.pages.extend(pdf.pages)
            except pikepdf.PasswordError:
                raise Exception(f"Incorrect password for {os.path.basename(path)}")
        
        # Save to a temporary file
        self._temp_merged_file = "temp_merged_booklet.pdf"
        merged_pdf.save(self._temp_merged_file)
        merged_pdf.close()
        return self._temp_merged_file

    def _parse_page_range(self, total_pages):
        """Parses a page range string (e.g., '1-5, 8, 10-12') into a list of indices."""
        range_str = self.options.get('page_range', '').strip()
        if not range_str:
            return list(range(total_pages))

        indices = set()
        parts = re.split(r'[,\s]+', range_str)
        for part in parts:
            if not part: continue
            if '-' in part:
                start, end = map(int, part.split('-'))
                for i in range(start, end + 1):
                    if 1 <= i <= total_pages:
                        indices.add(i - 1)
            else:
                i = int(part)
                if 1 <= i <= total_pages:
                    indices.add(i - 1)
        return sorted(list(indices))

    def _split_into_booklets(self, page_indices):
        """Splits the list of page indices into smaller chunks for separate booklets."""
        if not self.options.get('split_booklet', False):
            return [page_indices]
        
        sheets_per_booklet = self.options.get('sheets_per_booklet', 10)
        pages_per_booklet = sheets_per_booklet * 4
        
        chunks = []
        for i in range(0, len(page_indices), pages_per_booklet):
            chunks.append(page_indices[i:i + pages_per_booklet])
        return chunks

    def _reorder_for_booklet(self, page_indices):
        """Reorders page indices for booklet printing, supporting LTR and RTL."""
        num_pages = len(page_indices)
        
        # Each sheet has 4 pages, so total pages must be a multiple of 4.
        # We add blank pages at the end of the list of indices. A special value (-1) represents a blank page.
        padding = (4 - (num_pages % 4)) % 4
        padded_indices = page_indices + [-1] * padding
        n = len(padded_indices)
        
        reordered = []
        
        # Use a deque for efficient popping from both ends
        page_deque = deque(padded_indices)

        is_rtl = self.options.get('direction', 'LTR') == 'RTL'

        while page_deque:
            if is_rtl:
                # For RTL: front-left, front-right, back-left, back-right
                reordered.append(page_deque.popleft())  # 1
                reordered.append(page_deque.pop())      # n
                if page_deque:
                    reordered.append(page_deque.pop())      # n-1
                    reordered.append(page_deque.popleft())  # 2
            else: # LTR
                # For LTR: front-left, front-right, back-left, back-right
                reordered.append(page_deque.pop())      # n
                reordered.append(page_deque.popleft())  # 1
                if page_deque:
                    reordered.append(page_deque.popleft())  # 2
                    reordered.append(page_deque.pop())      # n-1

        return reordered

    def _create_final_pdf(self, source_pdf, ordered_indices):
        """Creates the final 2-up PDF booklet using reportlab and pdf2image, preserving aspect ratio and centering images."""
        # اختيار حجم الصفحة حسب الاتجاه
        base_size = self.paper_sizes.get(self.options['paper_size'], A4)
        if self.options.get('orientation', 'Portrait') == 'Landscape':
            output_size = (base_size[1], base_size[0])  # عكس العرض والارتفاع
        else:
            output_size = base_size
        output_width, output_height = output_size
        packet = io.BytesIO()
        can = rl_canvas.Canvas(packet, pagesize=output_size)

        temp_pdf_path = self._temp_merged_file if hasattr(self, '_temp_merged_file') else self.file_paths[0]
        images = convert_from_path(temp_pdf_path, dpi=300)

        for i in range(0, len(ordered_indices), 2):
            left_idx = ordered_indices[i]
            right_idx = ordered_indices[i+1] if i+1 < len(ordered_indices) else -1

            # رسم الصفحة اليسرى في منتصف النصف الأيسر
            if left_idx != -1:
                img_left = images[left_idx]
                orig_w, orig_h = img_left.size
                max_w, max_h = output_width / 2 * 0.9, output_height * 0.9  # استخدم 90% من نصف الصفحة كحد أقصى
                scale = min(max_w / orig_w, max_h / orig_h, 1.0)
                new_w, new_h = int(orig_w * scale), int(orig_h * scale)
                x_left = (output_width / 2 - new_w) / 2
                y_left = (output_height - new_h) / 2
                can.drawInlineImage(img_left, x_left, y_left, width=new_w, height=new_h)

            # رسم الصفحة اليمنى في منتصف النصف الأيمن
            if right_idx != -1:
                img_right = images[right_idx]
                orig_w, orig_h = img_right.size
                max_w, max_h = output_width / 2 * 0.9, output_height * 0.9
                scale = min(max_w / orig_w, max_h / orig_h, 1.0)
                new_w, new_h = int(orig_w * scale), int(orig_h * scale)
                x_right = output_width / 2 + (output_width / 2 - new_w) / 2
                y_right = (output_height - new_h) / 2
                can.drawInlineImage(img_right, x_right, y_right, width=new_w, height=new_h)

            can.showPage()

        can.save()
        packet.seek(0)
        output_pdf = pikepdf.open(packet)
        return output_pdf

    def _place_page(self, source_pdf, page_idx, target_page, target_rect):
        """Places a source page onto a target page within a specified rectangle."""
        if page_idx == -1:
            # Skip blank pages
            return
        source_page = source_pdf.pages[page_idx]
        media_box = source_page.MediaBox
        src_width = float(media_box[2]) - float(media_box[0])
        src_height = float(media_box[3]) - float(media_box[1])
        
        # Handle page rotation
        if src_width > src_height: # Landscape
            source_page.rotate(90, relative=True)

        # Add page numbers if enabled
        if self.options.get('add_page_numbers', False):
            page_num_to_display = page_idx + 1
            overlay_pdf_bytes = self._create_page_number_overlay(
                (src_width, src_height), 
                page_num_to_display
            )
            overlay_pdf = pikepdf.open(io.BytesIO(overlay_pdf_bytes))
            source_page.add_overlay(overlay_pdf.pages[0])

        # استخدم overlay لإضافة الصفحة الأصلية إلى الصفحة الجديدة
        target_page.add_overlay(source_page)
        # ملاحظة: هذا سيضع الصفحة الأصلية فوق الصفحة الجديدة بالكامل
        # إذا أردت وضعها في نصف الصفحة، تحتاج إلى معالجة إضافية بمكتبة أخرى

        # إذا أردت فقط معالجة مكان الصفحة، يجب استخدام مكتبة أخرى مثل PyPDF2 أو fitz (PyMuPDF)

    def _create_page_number_overlay(self, page_size, page_number):
        """Creates a temporary PDF with just a page number."""
        packet = io.BytesIO()
        can = rl_canvas.Canvas(packet, pagesize=page_size)
        width, height = page_size
        
        # Position at bottom center
        can.setFont("Helvetica", 9)
        can.drawCentredString(width / 2, 0.25 * inch, str(page_number))
        can.save()
        
        packet.seek(0)
        return packet.read()


# --- GUI APPLICATION ---

class BookletCreatorApp:
    """The main GUI application class."""

    def __init__(self, root):
        self.root = root
        self.root.title("Booklet Creator")
        self.root.geometry("800x650")

        self.file_paths = []
        self._setup_ui()
        self._update_ui_state()

    def _setup_ui(self):
        """Initializes and places all GUI widgets."""
        main_frame = tb.Frame(self.root, padding=15)
        main_frame.pack(fill=BOTH, expand=YES)

        # --- File Selection Frame ---
        file_frame = tb.Labelframe(main_frame, text="1. Select PDF Files", padding=10)
        file_frame.pack(fill=X, pady=(0, 10))
        
        # Treeview to list files
        tree_frame = tb.Frame(file_frame)
        tree_frame.pack(fill=X, expand=YES)
        
        self.file_tree = tb.Treeview(tree_frame, columns=("filename", "pages"), show="headings", height=5)
        self.file_tree.heading("filename", text="File Name")
        self.file_tree.heading("pages", text="Pages")
        self.file_tree.column("filename", width=400)
        self.file_tree.column("pages", width=50, anchor=CENTER)
        self.file_tree.pack(side=LEFT, fill=X, expand=YES)

        # Scrollbar for treeview
        scrollbar = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)

        # File action buttons
        file_button_frame = tb.Frame(file_frame)
        file_button_frame.pack(fill=X, pady=(10, 0))
        
        tb.Button(file_button_frame, text="Add PDF(s)", command=self._add_files, bootstyle=SUCCESS).pack(side=LEFT, padx=(0, 5))
        self.remove_btn = tb.Button(file_button_frame, text="Remove Selected", command=self._remove_selected_file, bootstyle=DANGER)
        self.remove_btn.pack(side=LEFT, padx=5)
        self.move_up_btn = tb.Button(file_button_frame, text="Move Up", command=lambda: self._move_item(-1))
        self.move_up_btn.pack(side=LEFT, padx=5)
        self.move_down_btn = tb.Button(file_button_frame, text="Move Down", command=lambda: self._move_item(1))
        self.move_down_btn.pack(side=LEFT, padx=5)


        # --- Options Frame ---
        options_frame = tb.Labelframe(main_frame, text="2. Configure Options", padding=10)
        options_frame.pack(fill=X, pady=10)

        # Left and Right option columns
        left_col = tb.Frame(options_frame)
        left_col.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))
        right_col = tb.Frame(options_frame)
        right_col.pack(side=RIGHT, fill=X, expand=YES, padx=(10, 0))

        # Page Range
        tb.Label(left_col, text="Page Range (e.g., 1-10, 15):").pack(anchor=W)
        self.page_range_var = tk.StringVar()
        tb.Entry(left_col, textvariable=self.page_range_var).pack(fill=X, pady=(0, 10))

        # Language Direction
        tb.Label(left_col, text="Language Direction:").pack(anchor=W)
        self.direction_var = tk.StringVar(value="LTR")
        tb.Radiobutton(left_col, text="Left-to-Right (LTR)", variable=self.direction_var, value="LTR").pack(anchor=W)
        tb.Radiobutton(left_col, text="Right-to-Left (RTL)", variable=self.direction_var, value="RTL").pack(anchor=W, pady=(0, 10))

        # Page Numbering
        self.add_numbers_var = tk.BooleanVar(value=True)
        tb.Checkbutton(left_col, text="Add page numbers", variable=self.add_numbers_var, bootstyle="primary-round-toggle").pack(anchor=W, pady=(5,0))


        # Paper Size
        tb.Label(right_col, text="Paper Size:").pack(anchor=W)
        self.paper_size_var = tk.StringVar(value="Auto")
        tb.Combobox(right_col, textvariable=self.paper_size_var, values=["Auto", "A4", "A5", "Letter", "Legal"], state="readonly").pack(fill=X, pady=(0, 10))

        # Page Orientation
        tb.Label(right_col, text="Page Orientation:").pack(anchor=W)
        self.orientation_var = tk.StringVar(value="Portrait")
        tb.Combobox(right_col, textvariable=self.orientation_var, values=["Portrait", "Landscape"], state="readonly").pack(fill=X, pady=(0, 10))

        # Split Booklet
        self.split_booklet_var = tk.BooleanVar(value=False)
        split_check = tb.Checkbutton(right_col, text="Split into smaller booklets", variable=self.split_booklet_var, command=self._toggle_split_options, bootstyle="primary-round-toggle")
        split_check.pack(anchor=W, pady=(10, 5))
        
        split_options_frame = tb.Frame(right_col)
        split_options_frame.pack(fill=X, padx=(20, 0))
        tb.Label(split_options_frame, text="Sheets per booklet:").pack(side=LEFT)
        self.sheets_per_booklet_var = tk.IntVar(value=10)
        self.split_spinbox = tb.Spinbox(split_options_frame, from_=1, to=100, textvariable=self.sheets_per_booklet_var, width=5)
        self.split_spinbox.pack(side=LEFT, padx=5)
        self._toggle_split_options() # Set initial state

        # --- Action Frame ---
        action_frame = tb.Labelframe(main_frame, text="3. Create Booklet", padding=10)
        action_frame.pack(fill=X, pady=10)

        self.create_btn = tb.Button(action_frame, text="Create Booklet", command=self._start_booklet_creation, bootstyle=(SUCCESS, OUTLINE))
        self.create_btn.pack(side=LEFT, padx=(0, 10))
        tb.Button(action_frame, text="Reset", command=self._clear_all, bootstyle=(WARNING, OUTLINE)).pack(side=LEFT)

        # --- Status Frame ---
        status_frame = tb.Frame(main_frame)
        status_frame.pack(fill=X, pady=(10, 0))

        self.status_var = tk.StringVar(value="Ready")
        tb.Label(status_frame, textvariable=self.status_var).pack(side=LEFT, anchor=W)

        self.progress_bar = tb.Progressbar(status_frame, mode='determinate', bootstyle=STRIPED)
        self.progress_bar.pack(fill=X, expand=YES, side=RIGHT)

    def _add_files(self):
        """Opens file dialog to select PDFs and adds them to the list."""
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if paths:
            for path in paths:
                if path not in self.file_paths:
                    try:
                        with pikepdf.open(path) as pdf:
                            pages = len(pdf.pages)
                        self.file_paths.append(path)
                        self.file_tree.insert("", END, values=(os.path.basename(path), pages))
                    except pikepdf.PasswordError:
                        messagebox.showwarning("Password Protected", f"{os.path.basename(path)} is password protected. Password handling is not yet implemented in this GUI version.")
                    except Exception as e:
                        messagebox.showerror("Error", f"Could not open {os.path.basename(path)}: {e}")
            self._update_ui_state()

    def _remove_selected_file(self):
        """Removes the selected file from the list."""
        selected_items = self.file_tree.selection()
        if not selected_items:
            return
        
        for item in selected_items:
            index = self.file_tree.index(item)
            self.file_paths.pop(index)
            self.file_tree.delete(item)
        
        self._update_ui_state()

    def _move_item(self, direction):
        """Moves a selected item up or down in the list."""
        selected = self.file_tree.selection()
        if not selected:
            return

        for item in selected:
            self.file_tree.move(item, self.file_tree.parent(item), self.file_tree.index(item) + direction)
            # Also update the underlying file_paths list
            idx = self.file_tree.index(item)
            path = self.file_paths.pop(idx - direction)
            self.file_paths.insert(idx, path)


    def _clear_all(self):
        """Resets the application to its initial state."""
        self.file_paths.clear()
        for i in self.file_tree.get_children():
            self.file_tree.delete(i)
        
        self.page_range_var.set("")
        self.direction_var.set("LTR")
        self.add_numbers_var.set(True)
        self.paper_size_var.set("Auto")
        self.split_booklet_var.set(False)
        self.sheets_per_booklet_var.set(10)
        self.status_var.set("Ready")
        self.progress_bar['value'] = 0
        self._toggle_split_options()
        self._update_ui_state()

    def _update_ui_state(self):
        """Enables or disables widgets based on the current state."""
        has_files = len(self.file_paths) > 0
        self.create_btn.config(state=NORMAL if has_files else DISABLED)
        self.remove_btn.config(state=NORMAL if has_files else DISABLED)
        self.move_up_btn.config(state=NORMAL if has_files else DISABLED)
        self.move_down_btn.config(state=NORMAL if has_files else DISABLED)

    def _toggle_split_options(self):
        """Enables or disables the split booklet spinbox."""
        state = NORMAL if self.split_booklet_var.get() else DISABLED
        self.split_spinbox.config(state=state)

    def _update_progress(self, value):
        self.progress_bar['value'] = value
        self.root.update_idletasks()

    def _update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()
        
    def _start_booklet_creation(self):
        """Gathers options and starts the PDF processing in a new thread."""
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please add at least one PDF file.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Booklet As",
            filetypes=[("PDF Files", "*.pdf")],
            defaultextension=".pdf",
            initialfile="booklet_output.pdf"
        )
        if not output_path:
            return

        options = {
            'output_path': output_path,
            'page_range': self.page_range_var.get(),
            'direction': self.direction_var.get(),
            'add_page_numbers': self.add_numbers_var.get(),
            'paper_size': self.paper_size_var.get(),
            'split_booklet': self.split_booklet_var.get(),
            'sheets_per_booklet': self.sheets_per_booklet_var.get(),
            'orientation': self.orientation_var.get(),
        }

        # Reorder file_paths based on Treeview order
        tree_items = self.file_tree.get_children()
        ordered_filenames = [self.file_tree.item(item)['values'][0] for item in tree_items]
        
        # Create a map of basename to full path for reordering
        path_map = {os.path.basename(p): p for p in self.file_paths}
        ordered_paths = [path_map[fname] for fname in ordered_filenames]
        self.file_paths = ordered_paths

        self.create_btn.config(state=DISABLED)
        self.progress_bar['value'] = 0

        # Run the PDF processing in a separate thread to avoid freezing the GUI
        thread = threading.Thread(
            target=self._run_processor,
            args=(self.file_paths, options)
        )
        thread.daemon = True
        thread.start()

    def _run_processor(self, file_paths, options):
        """The target function for the processing thread."""
        processor = PdfProcessor(file_paths, options, self._update_progress, self._update_status)
        success = processor.create_booklet()
        
        if success:
            messagebox.showinfo("Success", "Booklet created successfully!")
        else:
            messagebox.showerror("Error", f"An error occurred. Check status for details.")
            
        self.create_btn.config(state=NORMAL)
        self.progress_bar['value'] = 0


if __name__ == "__main__":
    # Use a modern ttkbootstrap theme
    root = tb.Window(themename="litera")
    app = BookletCreatorApp(root)
    root.mainloop()


