import fitz  # PyMuPDF
import os
from PIL import Image

class PDFEngine:
    def __init__(self):
        self.doc = None
        self.file_path = None
        self.clipboard_pages = [] # List of pixmaps or temp files? Keeping it simple for now

    def open_pdf(self, path):
        """Opens a PDF file."""
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(path)
            self.file_path = path
            return True, f"Loaded {len(self.doc)} pages."
        except Exception as e:
            return False, str(e)

    def save_pdf(self, path=None):
        """Saves the current PDF."""
        if not self.doc:
            return False, "No document open."
        
        target_path = path if path else self.file_path
        try:
            # Incremental save if overwriting and possible, else full lossless save
            if target_path == self.file_path:
                try:
                    self.doc.saveIncr() 
                except:
                    # Fallback to full save if incremental fails
                    self.doc.save(target_path, deflate=True, garbage=0)
            else:
                # Full save, deflate for lossless compression, no garbage collection to prevent recompression
                self.doc.save(target_path, deflate=True, garbage=0)
            return True, "Saved successfully."
        except Exception as e:
            # If incremental save fails (e.g. major changes), try full save
            try:
                self.doc.save(target_path, deflate=True, garbage=0)
                return True, "Saved successfully (Full)."
            except Exception as e2:
                return False, str(e2)

    def save_subset(self, page_indices, path):
        """Saves specific pages to a new PDF file."""
        if not self.doc:
            return False, "No document open."
        
        try:
            new_doc = fitz.open()
            # Sort indices to maintain order? Or follow selection order?
            # User might select 3, 1, 2. Usually we want 1, 2, 3.
            # But "Custom: 3, 1" might mean specific order.
            # Let's trust the input list order.
            
            for idx in page_indices:
                if 0 <= idx < len(self.doc):
                    new_doc.insert_pdf(self.doc, from_page=idx, to_page=idx)
            
            new_doc.save(path, deflate=True, garbage=0)
            new_doc.close()
            return True, "Saved successfully."
        except Exception as e:
            return False, str(e)

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None
            self.file_path = None

    def get_page_count(self):
        return len(self.doc) if self.doc else 0

    def get_page_image(self, page_index, scale=1.0):
        """Returns a PIL Image for a specific page."""
        if not self.doc or not (0 <= page_index < len(self.doc)):
            return None

        page = self.doc[page_index]
        matrix = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=matrix)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def rotate_page(self, page_index, angle):
        """Rotates a page by angle (90, -90, 180)."""
        if not self.doc: return
        page = self.doc[page_index]
        page.set_rotation(page.rotation + angle)

    def delete_pages(self, page_indices):
        """Deletes pages. Indices should be a list of integers."""
        if not self.doc: return
        # Delete in reverse order to avoid index shifting problems
        for idx in sorted(page_indices, reverse=True):
            self.doc.delete_page(idx)

    def move_page(self, from_index, to_index):
        """Moves a page from one index to another."""
        if not self.doc: return
        self.doc.move_page(from_index, to_index)

    def create_blank_page(self, width=595, height=842, insert_at=-1):
        """Creates a blank page."""
        if not self.doc:
             self.doc = fitz.open() # Create new if none
        
        self.doc.new_page(pno=insert_at, width=width, height=height)

    def insert_pdf(self, path, insert_at=-1):
        """Merges another PDF into the current one."""
        if not self.doc:
             self.doc = fitz.open() # Create new if none
             
        try:
            src = fitz.open(path)
            self.doc.insert_pdf(src, start_at=insert_at)
            src.close()
            return True
        except Exception as e:
            return False

    def export_selection(self, page_indices, output_path):
        """Exports selected pages to a new PDF."""
        if not self.doc: return False
        
        try:
            new_doc = fitz.open()
            new_doc.insert_pdf(self.doc, from_page=page_indices[0], to_page=page_indices[-1]) # This is range, need list support
            # PyMuPDF insert_pdf takes from/to. For arbitrary list, we can use select?
            # actually better: new_doc.insert_pdf(self.doc, links=False, annots=False, show_progress=0) and then layout?
            # Or just:
            new_doc = fitz.open()
            # method 2: insert one by one
            for pid in page_indices:
                new_doc.insert_pdf(self.doc, from_page=pid, to_page=pid)
            
            new_doc.save(output_path)
            new_doc.close()
            return True
        except Exception as e:
            print(e)
            return False

    def extract_text(self, page_index):
        """Extracts text from a specific page."""
        if not self.doc or not (0 <= page_index < len(self.doc)):
            return None
        return self.doc[page_index].get_text()

    def merge_pdf_list(self, file_paths):
        """Merges multiple PDFs into the current document or a new one."""
        if not self.doc:
            self.doc = fitz.open()

        count = 0
        for path in file_paths:
            try:
                self.doc.insert_pdf(fitz.open(path))
                count += 1
            except Exception as e:
                print(f"Failed to merge {path}: {e}")
        return count > 0

    def add_watermark(self, text, page_indices=None):
        """Adds text watermark to specified pages (or all)."""
        if not self.doc: return

        pages = page_indices if page_indices else range(len(self.doc))
        
        for pid in pages:
            page = self.doc[pid]
            # Calculate center
            rect = page.rect
            center = fitz.Point(rect.width/2, rect.height/2)
            
            # Insert text with rotation
            # morph = (center, fitz.Matrix(45)) 
            # page.insert_text(center, text, fontsize=50, color=(0.7, 0.7, 0.7), rotate=45, align=1)
            # Use insert_textbox for better centering if needed, but simple text for now
            
            # Create a watermark using Shape
            shape = page.new_shape()
            shape.insert_text(center, text, fontsize=60, color=(0.8, 0.8, 0.8), rotate=45, align=1)
            shape.commit()
            
            
