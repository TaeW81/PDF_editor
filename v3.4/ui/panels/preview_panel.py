import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk

class PreviewPanel(ttk.Frame):
    def __init__(self, master, pdf_engine, on_page_change=None, bootstyle="dark", **kwargs):
        super().__init__(master, bootstyle=bootstyle, **kwargs)
        self.pdf = pdf_engine
        self.on_page_change = on_page_change
        self.current_page_index = 0
        self.photo_image = None
        
        # UI Components
        self.lbl_title = ttk.Label(self, text="Preview", font=("Segoe UI", 10, "bold"), bootstyle="inverse-dark", padding=5)
        self.lbl_title.pack(fill=X)
        
        self.canvas = tk.Canvas(self, bg="#555555")
        self.v_scroll = ttk.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self, orient=HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        
        self.v_scroll.pack(side=RIGHT, fill=Y)
        self.h_scroll.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        
        # Zoom state
        self.zoom_scale = 1.0

        # Mouse Wheel binding
        self.canvas.bind("<Control-MouseWheel>", self._on_zoom)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

    def scroll(self, delta):
        # returns (start_fraction, end_fraction)
        # e.g. (0.0, 1.0) means all visible.
        # (0.0, 0.5) means top half visible.
        
        view_start, view_end = self.v_scroll.get()
        
        # Scroll Sensitivity
        units = int(-1*(delta/120))
        
        if delta > 0 and view_start <= 0.0:
            # At top, going up -> Prev Page
            self.change_page(-1)
        elif delta < 0 and view_end >= 1.0:
            # At bottom, going down -> Next Page
            self.change_page(1)
        else:
            # Just scroll content
            self.canvas.yview_scroll(units, "units")

    def change_page(self, offset):
        new_idx = self.current_page_index + offset
        if 0 <= new_idx < self.pdf.get_page_count():
            self.show_page(new_idx)
            # If we went to next page, scroll to top?
            if offset > 0:
                self.canvas.yview_moveto(0.0)
            else:
                self.canvas.yview_moveto(1.0) # Bottom of prev page?
                # Actually usually top of prev page is better UX for reading?
                # But if scrolling UP, bottom makes sense.
                pass

    def _on_mousewheel(self, event):
        self.scroll(event.delta)

    def _on_zoom(self, event):
        self.zoom(event.delta)
        return "break"

    def zoom(self, delta):
        if delta > 0:
            self.zoom_scale = min(4.0, self.zoom_scale + 0.1)
        else:
            self.zoom_scale = max(0.1, self.zoom_scale - 0.1)
        
        # Refresh current page with new scale
        self.show_page(self.current_page_index)

    def show_page(self, index):
        if not self.pdf.doc or not (0 <= index < len(self.pdf.doc)):
            return

        if self.current_page_index != index:
            self.current_page_index = index
            if self.on_page_change:
                self.on_page_change(index)
        
        self.current_page_index = index
        
        # Get image from engine with ZOOM SCALE
        pil_img = self.pdf.get_page_image(index, scale=self.zoom_scale) 
        
        if pil_img:
            self.photo_image = ImageTk.PhotoImage(pil_img)
            
            # Center image on canvas
            c_width = self.canvas.winfo_width()
            c_height = self.canvas.winfo_height()
            
            if c_width < 100: c_width = 800 # Fallback if not mapped
            if c_height < 100: c_height = 600

            img_w = self.photo_image.width()
            img_h = self.photo_image.height()
            
            x = max(0, (c_width - img_w) // 2)
            y = max(0, (c_height - img_h) // 2)

            self.canvas.delete("all")
            self.canvas.create_image(x, y, anchor="nw", image=self.photo_image)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def show_logo(self):
        import os
        # Try to find logo
        # Assuming data dir is parallel to core/ui, or in ../data
        # We need to find the correct path. existing code used os.path.dirname(__file__) + data
        # In v3.4, __file__ is ui/panels/preview_panel.py. 
        # root is ../../. data is ../../data ??
        # Let's use config.settings.DATA_DIR if available, or relative path
        
        try:
             # Go up 3 levels from ui/panels/preview_panel.py -> ui/panels -> ui -> v3.4
             base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
             logo_path = os.path.join(base_dir, "data", "kunhwa_logo.png")
             
             # Fallback to older location if needed
             if not os.path.exists(logo_path):
                  # Try d:/Gits/PDF_editor/data/kunhwa_logo.png
                  logo_path = os.path.join(os.path.dirname(base_dir), "data", "kunhwa_logo.png")

             if os.path.exists(logo_path):
                 img = Image.open(logo_path)
                 # Resize to fit reasonable size
                 img.thumbnail((400, 300))
                 self.logo_image = ImageTk.PhotoImage(img) # Keep ref
                 
                 c_width = self.canvas.winfo_width()
                 c_height = self.canvas.winfo_height()
                 x = c_width // 2
                 y = c_height // 2
                 
                 self.canvas.delete("all")
                 self.canvas.create_image(x, y, anchor="center", image=self.logo_image)
             else:
                 self.canvas.delete("all")
                 self.canvas.create_text(200, 200, text="Kunhwa PDF Editor", font=("Segoe UI", 20, "bold"), fill="gray")
                 
        except Exception as e:
             print(f"Logo error: {e}")

    def clear(self):
        self.canvas.delete("all")
        self.show_logo()
