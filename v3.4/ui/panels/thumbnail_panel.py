import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk

class ThumbnailPanel(ttk.Frame):
    def __init__(self, master, pdf_engine, on_selection_change, drag_manager=None, bootstyle="secondary", **kwargs):
        super().__init__(master, bootstyle=bootstyle, **kwargs)
        self.pdf = pdf_engine
        self.on_selection_change = on_selection_change
        self.drag_manager = drag_manager
        
        self.thumbnails = [] # List of PhotoImage
        self.thumb_widgets = [] # List of (frame, index)
        self.selected_indices = set()
        self.scale = 0.2
        
        # Drag and Drop state
        self.drag_start_index = None
        self.dragged_item = None
        
        # UI Components
        self.lbl_title = ttk.Label(self, text="썸네일", font=("Segoe UI", 10, "bold"), bootstyle="inverse-secondary", padding=5)
        self.lbl_title.pack(fill=X)
        
        self.canvas = tk.Canvas(self, bg="#f0f0f0") 
        self.scrollbar = ttk.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        
        self.canvas.bind("<Configure>", self._on_resize)
        
        # Explicit bindings to ensure they work when panel has focus
        self.bind("<Control-a>", lambda e: self.select_all())
        self.bind("<Escape>", lambda e: self.deselect_all())
        self.bind("<Delete>", lambda e: self.winfo_toplevel().on_delete_pages())
        self.bind("<Control-c>", lambda e: self.winfo_toplevel().on_copy())
        self.bind("<Control-v>", lambda e: self.winfo_toplevel().on_paste())
        
        # Mousewheel Focus Handling

    def scroll(self, delta):
        self.canvas.yview_scroll(int(-1*(delta/120)), "units")

    def zoom(self, delta):
        if delta > 0:
            self.scale = min(2.0, self.scale + 0.1) 
        else:
            self.scale = max(0.1, self.scale - 0.1)
        self.refresh()

    def select_all(self):
        if not self.pdf.doc: return
        self.selected_indices = set(range(len(self.pdf.doc)))
        self.on_selection_change(self.selected_indices)
        self.refresh_selection_visuals()

    def deselect_all(self):
        self.selected_indices.clear()
        self.on_selection_change(self.selected_indices)
        self.refresh_selection_visuals()

    def _on_resize(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.update_grid_layout()

    def refresh(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.thumbnails.clear()
        self.thumb_widgets.clear()
        
        if not self.pdf.doc:
            return

        for i in range(len(self.pdf.doc)):
            # Frame style depends on selection
            style = "primary" if i in self.selected_indices else "light"
            
            frame = ttk.Frame(self.scroll_frame, padding=5, bootstyle=style)
            
            pil_img = self.pdf.get_page_image(i, scale=self.scale)
            if pil_img:
                tk_img = ImageTk.PhotoImage(pil_img)
                self.thumbnails.append(tk_img)
                
                lbl_img = ttk.Label(frame, image=tk_img)
                lbl_img.pack(side=TOP, pady=2)
                self._bind_events(lbl_img, i)
            
            lbl_num = ttk.Label(frame, text=f"{i+1}", font=("Segoe UI", 9))
            lbl_num.pack(side=BOTTOM)
            self._bind_events(lbl_num, i)
            
            self._bind_events(frame, i)
            self.thumb_widgets.append(frame)

        self.update_grid_layout()
        
    def update_grid_layout(self):
        if not self.thumb_widgets: return
        canvas_width = self.canvas.winfo_width()
        if canvas_width < 50: canvas_width = 300
        
        self.thumb_widgets[0].update_idletasks()
        item_width = self.thumb_widgets[0].winfo_reqwidth() + 10
        if item_width < 50: item_width = 150
        
        columns = max(1, canvas_width // item_width)
        
        for i, frame in enumerate(self.thumb_widgets):
            row = i // columns
            col = i % columns
            frame.grid(row=row, column=col, padx=5, pady=5, sticky="n")

    def _bind_events(self, widget, index):
        # Click
        widget.bind("<Button-1>", lambda e, idx=index: self._on_click_proxy(e, idx))
        # Drag and Drop
        widget.bind("<ButtonPress-1>", lambda e, idx=index: self._on_drag_start(e, idx))
        widget.bind("<B1-Motion>", self._on_drag_motion)
        widget.bind("<ButtonRelease-1>", self._on_drag_release)

    def _on_click_proxy(self, event, index):
        # Allow MainWindow to detect focus or just let click handle selection
        pass 
        # We don't need focus_set() if MainWindow handles global events.
        # But focus is good for visual cues if we add them later.
        self.focus_set()

    def _on_delete(self, event=None):
        if not self.selected_indices: return
        # Logic to delete pages.
        # We need to call a method on MainWindow or handle it via callback?
        # Using on_selection_change to signal? No, that's for UI update.
        # We should probably expose a delete callback or let MainWindow handle the bind?
        # But bindings are local to widget focus.
        # Let's bubble event or call a method if we have reference?
        # We have self.pdf. 
        # But we need UI refresh and main window status update.
        # Ideally MainWindow binds Delete.
        # For now, let's try to find MainWindow.
        try:
            self.winfo_toplevel().on_delete_pages()
        except AttributeError:
            pass # Method might not exist yet
        return "break"

    def _on_click(self, event, index):
        # Logic is now in _on_drag_start to prevent conflict?
        # Actually, Button-1 fires on press. ButtonPress-1 also fires on press.
        # They are duplicates.
        # Let's remove _on_click binding and put EVERYTHING in _on_drag_start?
        # Or keep _on_click for selection and _on_drag_start for drag init.
        # But if we drag, we usually select first.
        pass

    def _on_drag_start(self, event, index):
        self.drag_start_index = index
        
        # Check modifiers
        if event.state & 0x0004: # Control
            # Toggle
            if index in self.selected_indices:
                self.selected_indices.remove(index)
            else:
                self.selected_indices.add(index)
        elif event.state & 0x0001: # Shift
            # Range
            if self.selected_indices:
                start = sorted(list(self.selected_indices))[-1] # From last clicked?
                # A better approach requires tracking 'anchor'. For now:
                if not hasattr(self, 'last_clicked_index'): self.last_clicked_index = 0
                start = self.last_clicked_index
                end = index
                
                # Select range
                self.selected_indices = set(range(min(start, end), max(start, end) + 1))
            else:
                self.selected_indices = {index}
        else:
            # No modifier
            # If item is NOT selected, select it (and deselect others).
            # If item IS selected, don't deselect others YET (because we might be dragging multiple).
            # But if we just click (release without drag), we want to select only this.
            # This is the tricky part of DnD + Selection.
            
            if index not in self.selected_indices:
                self.selected_indices = {index}
            # If it IS in selection, we keep it as is (so we can drag the group)
            # BUT if we release without dragging, we should select only this.
            # That requires logic in ButtonRelease.
            
        self.last_clicked_index = index
        self.on_selection_change(self.selected_indices)
        self.refresh_selection_visuals()

    def _on_drag_motion(self, event):
        if self.drag_start_index is not None and self.drag_manager:
            if not self.drag_manager.dragging:
                # Start Global Drag
                win = self.winfo_toplevel()
                self.drag_manager.start_drag(win, self.selected_indices, event)

    def _on_drag_release(self, event):
        # If we didn't drag (or dragged very little), and no modifiers were pressed,
        # we should select ONLY the clicked item (deselect others).
        # We need to know if it was a drag or a click.
        # DragManager handles actual global drag stop.
        # But locally, if drag_start_index is set but Global Drag didn't start...
        
        if self.drag_start_index is not None:
             # Check if it was a simple click on a multi-selection
             if not (event.state & 0x0004) and not (event.state & 0x0001):
                 # No modifiers
                 # If we are here, we might have skipped deselecting in start.
                 # Only deselect if we didn't actually drag?
                 if not self.drag_manager or not self.drag_manager.dragging:
                     self.selected_indices = {self.drag_start_index}
                     self.on_selection_change(self.selected_indices)
                     self.refresh_selection_visuals()

        self.drag_start_index = None
        pass

    def refresh_selection_visuals(self):
        for i, frame in enumerate(self.thumb_widgets):
            style = "primary" if i in self.selected_indices else "light"
            frame.configure(bootstyle=style)

    def select_and_scroll_to(self, index):
        if not self.pdf.doc or not (0 <= index < len(self.pdf.doc)):
            return
            
        # Select
        self.selected_indices = {index}
        self.refresh_selection_visuals() # Don't trigger on_selection_change to avoid loop if possible?
        # If we trigger on_selection_change, it calls Preview.show_page. 
        # If Preview called this, we have a loop!
        # Pass a flag to suppress callback?
        # Or just don't call on_selection_change here. 
        # But we want to update status bar?
        # Let's call it, but guard in MainWindow?
        # Or Update PreviewPanel to not call callback if page is same?
        
        # Scroll
        if 0 <= index < len(self.thumb_widgets):
            widget = self.thumb_widgets[index]
            # Ensure layout is updated
            self.update_idletasks()
            
            # Get widget position relative to scroll_frame
            y = widget.winfo_y()
            h = widget.winfo_height()
            
            # Get scrollframe height
            sf_h = self.scroll_frame.winfo_height()
            canvas_h = self.canvas.winfo_height()
            
            if sf_h > canvas_h:
                # Center the item
                target_y = y - (canvas_h - h) / 2
                fraction = target_y / sf_h
                fraction = max(0.0, min(1.0, fraction))
                self.canvas.yview_moveto(fraction)
