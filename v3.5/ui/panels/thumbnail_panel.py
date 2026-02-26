import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES
from config.settings import APP_NAME, VERSION

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
        self.lbl_title = ttk.Label(self, text="썸네일", font=("맑은 고딕", 10, "bold"), bootstyle="inverse-secondary", padding=5)
        self.lbl_title.pack(fill=X)
        
        self.canvas = tk.Canvas(self, bg="#f0f0f0") 
        self.scrollbar = ttk.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack Scrollbar FIRST to ensure it reserves space
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        
        self.canvas.bind("<Configure>", self._on_resize)
        
        # Explicit bindings to ensure they work when panel has focus
        self.bind("<Control-a>", lambda e: self.select_all())
        self.bind("<Escape>", lambda e: self.deselect_all())
        self.bind("<Delete>", lambda e: self.winfo_toplevel().on_delete_pages())
        self.bind("<Control-c>", lambda e: self.winfo_toplevel().on_copy())
        self.bind("<Control-v>", lambda e: self.winfo_toplevel().on_paste())
        
        # OS File Drag and Drop Bindings
        self.canvas.drop_target_register(DND_FILES)
        self.canvas.dnd_bind('<<Drop>>', self._on_os_file_drop)
        
        # Bind Map event to trigger initial refresh when actually visible
        self.bind("<Map>", self._on_map_event)
        self._first_map = True
        
        # Mousewheel Focus Handling

    def _on_os_file_drop(self, event):
        """Handle files dropped from OS."""
        files = self.canvas.tk.splitlist(event.data)
        if not files: return
        
        main_win = self.winfo_toplevel()
        
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if not pdf_files: return
        
        if not self.pdf.doc:
            file_path = pdf_files[0]
            # Call open_pdf logic from MainWindow
            success, msg = main_win.pdf.open_pdf(file_path)
            if success:
                main_win.status_bar.config(text=f"열림: {file_path}")
                self.set_filename(os.path.basename(file_path))
                
                # Update Title
                filename = os.path.basename(file_path)
                user = main_win.auth.get_current_user_name()
                main_win.title(f"{filename} - {APP_NAME} {VERSION} [{user}]")
                
                main_win.after(200, main_win._refresh_on_open)
            else:
                import tkinter.messagebox as messagebox
                messagebox.showerror("오류", msg)
        else:
            target_index = self.get_index_at(event.x_root, event.y_root)
            if target_index == -1:
                target_index = -1
                
            success_count = 0
            for file_path in pdf_files:
                old_len = len(self.pdf.doc)
                if main_win.pdf.insert_pdf(file_path, insert_at=target_index):
                    success_count += 1
                    if target_index != -1:
                        target_index += (len(self.pdf.doc) - old_len)
                        
            if success_count > 0:
                self.refresh()
                main_win.status_bar.config(text=f"PDF {success_count}개 파일 병합 완료.")
            else:
                import tkinter.messagebox as messagebox
                messagebox.showerror("오류", "병합 실패.")

    def set_filename(self, filename):
        if filename:
            self.lbl_title.config(text=f"썸네일 - {filename}")
        else:
            self.lbl_title.config(text="썸네일")

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
    def _on_map_event(self, event):
        """Called when the widget becomes mapped (visible) on screen."""
        if self._first_map and self.pdf.doc:
            print("ThumbnailPanel Mapped - Triggering Initial Refresh")
            self._first_map = False
            # Force a refresh now that we have geometry
            self.refresh()
            
    def refresh(self):
        # Clear existing
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        # Ensure frame is above canvas background
        self.scroll_frame.lift()
        
        self.thumbnails.clear()
        self.thumb_widgets.clear()
        
        if not self.pdf.doc:
            return

        for i in range(len(self.pdf.doc)):
            try:
                # Frame style depends on selection
                style = "primary" if i in self.selected_indices else "light"
                
                frame = ttk.Frame(self.scroll_frame, padding=5, bootstyle=style)
                
                # Use consistent scale
                pil_img = self.pdf.get_page_image(i, scale=self.scale)
                if pil_img:
                    tk_img = ImageTk.PhotoImage(pil_img)
                    self.thumbnails.append(tk_img)
                    
                    lbl_img = ttk.Label(frame, image=tk_img)
                    lbl_img.pack(side=TOP, pady=2)
                    self._bind_events(lbl_img, i)
                else:
                    # Fallback
                    lbl_img = ttk.Label(frame, text="Error", width=10)
                    lbl_img.pack(side=TOP, pady=2)
                
                lbl_num = ttk.Label(frame, text=f"{i+1}", font=("맑은 고딕", 9))
                lbl_num.pack(side=BOTTOM)
                self._bind_events(lbl_num, i)
                
                self._bind_events(frame, i)
                self.thumb_widgets.append(frame)
            except Exception as e:
                print(f"Error loading thumbnail {i}: {e}")

        # Update Grid
        self.update_grid_layout()
        
        # Note: update_grid_layout now handles the scrollregion update robustly.
        # We don't force yview_moveto(0) here because it causes the view to jump to the top when zooming,
        # which is annoying for the user. We just let the canvas stay at its scroll position.
        
    def _update_scrollregion(self):
        # Kept for compatibility if used elsewhere, but refresh does it inline now
        self.scroll_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_resize(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.update_grid_layout(event.width)

    def update_grid_layout(self, width=None):
        if not self.thumb_widgets: return
        
        if width is not None:
             canvas_width = width
        else:
             canvas_width = self.canvas.winfo_width()
        
        if canvas_width < 50: canvas_width = 300 
        
        # Force the inner frame width slightly smaller than canvas to prevent horizontal overflow
        safe_canvas_width = canvas_width - 15
        if safe_canvas_width < 50: safe_canvas_width = 50
        
        self.canvas.itemconfig(self.canvas_window, width=safe_canvas_width)
        
        # Ensure sizes are mathematically computed
        self.scroll_frame.update_idletasks()
        max_frame_width = max((f.winfo_reqwidth() for f in self.thumb_widgets), default=50)
        
        item_width = max_frame_width + 10 # Adding safe edge margin for grid (padx=5 -> 10px)
        if item_width < 50: item_width = 50
        
        columns = max(1, safe_canvas_width // item_width)
        
        # Reset and apply column weights for centering so nothing clips at the right edge
        for c in range(50):
            self.scroll_frame.grid_columnconfigure(c, weight=0)
        for c in range(columns):
            self.scroll_frame.grid_columnconfigure(c, weight=1)
            
        for i, frame in enumerate(self.thumb_widgets):
            row = i // columns
            col = i % columns
            frame.grid(row=row, column=col, padx=5, pady=5, sticky="n")
            
        # Critically important: update layout then set scrollregion
        self.scroll_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Fallback for dynamic resizing weirdness
        self.after(50, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

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
        self.drag_start_pos = (event.x_root, event.y_root)
        self.has_dragged = False
        
        # Check modifiers
        # Note: In start, we might select immediately if modifiers are used.
        # Logic: 
        # Modifiers -> Immediate Action (Toggle/Range)
        # No Modifier -> Deferred Action (Wait for Release to see if dragged)
        
        if event.state & 0x0004: # Control
            # Toggle
            if index in self.selected_indices:
                self.selected_indices.remove(index)
            else:
                self.selected_indices.add(index)
            self.on_selection_change(self.selected_indices)
            self.refresh_selection_visuals()
            
        elif event.state & 0x0001: # Shift
            # Range
            if self.selected_indices:
                start = sorted(list(self.selected_indices))[-1] 
                if hasattr(self, 'last_clicked_index'): 
                    # Prefer anchor if we track it properly
                    start = self.last_clicked_index
                    
                end = index
                self.selected_indices = set(range(min(start, end), max(start, end) + 1))
            else:
                self.selected_indices = {index}
            self.on_selection_change(self.selected_indices)
            self.refresh_selection_visuals()
            
        else:
            # No modifier
            # If item is NOT selected, select it immediately (adding to others? No, standard is Select Only)
            # Standard Explorer: Left down on unselected -> Select Only.
            # Left down on SELECTED -> Do nothing (prepare to drag group).
            
            if index not in self.selected_indices:
                self.selected_indices = {index}
                self.on_selection_change(self.selected_indices)
                self.refresh_selection_visuals()
                
        self.last_clicked_index = index

    def _on_drag_motion(self, event):
        if self.drag_start_index is not None:
            # Check Threshold
            if not self.has_dragged:
                dx = abs(event.x_root - self.drag_start_pos[0])
                dy = abs(event.y_root - self.drag_start_pos[1])
                if dx > 5 or dy > 5:
                    self.has_dragged = True
                    
            if self.has_dragged and self.drag_manager:
                if not self.drag_manager.dragging:
                    # Start Global Drag
                    win = self.winfo_toplevel()
                    self.drag_manager.start_drag(win, self.selected_indices, event)

    def _on_drag_release(self, event):
        # Click Logic (Mouse Up without Drag)
        # Only for No-Modifier case where we deferred deselecting.
        
        if self.drag_start_index is not None and not self.has_dragged:
             # It was a click.
             
             # Check modifiers again (just in case they were released? No, we care if they were used in start)
             # But if I Ctrl+Click, logic ran in Start. Release does nothing.
             # If I Click (No Mod), logic ran in Start for UNSELECTED.
             # If I Click (No Mod) on SELECTED, Start did nothing. Release must Select Only it.
             
             is_ctrl = event.state & 0x0004
             is_shift = event.state & 0x0001
             
             if not is_ctrl and not is_shift:
                 # Select only this one (Deselect others)
                 self.selected_indices = {self.drag_start_index}
                 self.on_selection_change(self.selected_indices)
                 self.refresh_selection_visuals()

        self.drag_start_index = None
        self.has_dragged = False
        self.drag_start_pos = None

    def refresh_selection_visuals(self):
        for i, frame in enumerate(self.thumb_widgets):
            style = "primary" if i in self.selected_indices else "light"
            frame.configure(bootstyle=style)

    def select_and_scroll_to(self, index):
        if not self.pdf.doc or not (0 <= index < len(self.pdf.doc)):
            return
            
        # Select
        self.selected_indices = {index}
        self.refresh_selection_visuals() 
        
        # Scroll
        if 0 <= index < len(self.thumb_widgets):
            widget = self.thumb_widgets[index]
            self.update_idletasks()
            
            y = widget.winfo_y()
            h = widget.winfo_height()
            sf_h = self.scroll_frame.winfo_height()
            canvas_h = self.canvas.winfo_height()
            
            if sf_h > canvas_h:
                target_y = y - (canvas_h - h) / 2
                fraction = target_y / sf_h
                fraction = max(0.0, min(1.0, fraction))
                self.canvas.yview_moveto(fraction)

    def get_index_at(self, x, y):
        """Returns the index of the item at screen coordinates (x, y)."""
        # Iterate all widgets and check intersection
        # Since we have thumb_widgets list corresponding to indices
        for i, widget in enumerate(self.thumb_widgets):
            wx = widget.winfo_rootx()
            wy = widget.winfo_rooty()
            w = widget.winfo_width()
            h = widget.winfo_height()
            
            if wx <= x <= wx + w and wy <= y <= wy + h:
                return i
        return -1
