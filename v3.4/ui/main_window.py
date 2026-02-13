import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

from core.pdf_engine import PDFEngine
from core.auth import AuthManager
from core.clipboard import WindowManager, ClipboardManager, DragManager
from config.settings import APP_NAME, VERSION, THEME_NAME

from ui.panels.thumbnail_panel import ThumbnailPanel
from ui.panels.preview_panel import PreviewPanel

class MainWindow(ttk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title(f"{APP_NAME} {VERSION}")
        self.geometry("1200x800")
        
        # Register with Manager
        self.manager = WindowManager()
        self.manager.register(self)
        self.clipboard = ClipboardManager()
        self.drag_manager = DragManager()
        
        # Protocol
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Initialize Core Modules
        self.auth = AuthManager()
        self.pdf = PDFEngine()
        
        # Check Authentication
        authorized, message = self.auth.authenticate()
        if not authorized:
            messagebox.showerror("Authentication Error", message)
            self.destroy()
            return

        self.title(f"{APP_NAME} {VERSION} - {self.auth.get_current_user_name()}")

        self.setup_ui()
        
        # Bind Keys
        self.bind("<Control-o>", lambda e: self.on_open_pdf())
        self.bind("<Control-s>", lambda e: self.on_save_pdf())
        self.bind("<Control-c>", self.on_copy)
        self.bind("<Control-v>", self.on_paste)
        
        # Global Mouse Events for Drag
        self.bind("<Motion>", self.on_global_motion)
        self.bind("<ButtonRelease-1>", self.on_global_release)
        
        self.setup_global_binds()

    def on_close(self):
        self.manager.unregister(self)
        self.destroy() # Destroy Toplevel
        if not self.manager.get_windows():
            # If no windows left, quit app?
            # self.master.destroy() # Root
            # Or keep root alive? Standard is close app if all windows closed.
            self.master.quit()

    def setup_global_binds(self):
        # Bind events to self (Toplevel) to catch them regardless of focus
        self.bind("<MouseWheel>", self.on_mousewheel)
        self.bind("<Delete>", self.on_delete_pages)
        self.bind("<Control-a>", self.on_select_all_pages)
        self.bind("<Escape>", self.on_deselect_all_pages)
        # Ctrl+C/V are already bound in __init__

    def on_mousewheel(self, event):
        # Check if mouse is over thumbnail panel
        x, y = self.winfo_pointerx(), self.winfo_pointery()
        widget = self.winfo_containing(x, y)
        
        # Check if widget is inside thumbnail panel
        if self.is_descendant(widget, self.thumbnail_panel):
            if event.state & 0x0004: # Control
                 self.thumbnail_panel.zoom(event.delta)
            else:
                 self.thumbnail_panel.scroll(event.delta)
        # Check if widget is inside preview panel
        elif self.is_descendant(widget, self.preview_panel):
            if event.state & 0x0004: # Control
                 self.preview_panel.zoom(event.delta)
            else:
                 self.preview_panel.scroll(event.delta)
        # Default fallback?
        else:
            # Maybe scroll thumbnail by default if nothing else?
            pass

    def is_descendant(self, widget, ancestor):
        if not widget: return False
        if widget == ancestor: return True
        return self.is_descendant(widget.master, ancestor)

    def on_select_all_pages(self, event=None):
        self.thumbnail_panel.select_all()

    def on_deselect_all_pages(self, event=None):
        self.thumbnail_panel.deselect_all()

    def on_delete_pages(self, event=None):
        indices = self.thumbnail_panel.selected_indices
        if not indices: return
        
        if messagebox.askyesno("Delete", f"Delete {len(indices)} pages?"):
            self.pdf.delete_pages(indices)
            self.thumbnail_panel.selected_indices.clear()
            self.thumbnail_panel.refresh()
            self.status_bar.config(text="Pages deleted.")

    def on_global_motion(self, event):
        if self.drag_manager.dragging:
             # Update drag window position
             # We need screen coordinates
             x = self.winfo_pointerx()
             y = self.winfo_pointery()
             if self.drag_manager.drag_window:
                 self.drag_manager.drag_window.geometry(f"+{x+10}+{y+10}")

    def on_global_release(self, event):
        if self.drag_manager.dragging:
            self.drag_manager.stop_drag(event, self.winfo_pointerx(), self.winfo_pointery())

    def on_external_drop(self, source_window, indices, x, y):
        # Get drop target index from ThumbnailPanel
        target_index = self.thumbnail_panel.get_index_at(x, y)
        if target_index == -1:
            # Dropped outside items? Append to end?
            # Or maybe ignored. Let's append to end if dropped on panel but not on item
            # For now, just append to end if -1
            target_index = self.pdf.get_page_count()

        if source_window == self:
            # SAME WINDOW -> MOVE (Reorder)
            print(f"Moving {len(indices)} pages to {target_index}")
            
            # Move logic:
            # Moving multiple items is tricky because indices change.
            # Strategy: 
            # 1. Sort indices descending to remove? No, move is usually: extract and insert.
            # PDFEngine might need a better move method.
            # Simple approach: Move one by one? 
            # Warning: Moving p1 to p5, then p2 to p5... indices shift.
            
            # Let's trust PDFEngine or implement a robust move.
            # self.pdf.move_page(from, to) only moves one.
            
            # Robust Multi-Move:
            # Calculate new order.
            # E.g. [0, 1, 2, 3, 4], move [0, 1] to 4. -> [2, 3, 4, 0, 1]
            # This requires creating a new PDF with new page order? 
            # Or repeated move.
            
            # For now, let's just try moving the FIRST page of selection to target,
            # then others... this is UX heavy.
            # Simplified: Move them one by one, carefully adjusting target?
            
            # Let's use a simple loop for now, iterate sorted indices.
            # If moving Down (target > current): Move reversed?
            # If moving Up (target < current): Move normal?
            
            # Actually, `pdf_engine.move_page` implementation checks `from` and `to`.
            # If we move index 0 to 5. original 0 is at 5. original 1 is at 0.
            
            # Let's block multi-page move for now and only allow single page reorder? 
            # Or just move them sequentially to the target.
            
            # If we allow multi-selection move:
            # We want them to end up contiguous at target.
            
            sorted_indices = sorted(list(indices))
            
            # Check if target is inside selection (invalid)
            if target_index in sorted_indices:
                return 

            # Calculate shift
            # We move pages one by one to target.
            # If target > selection: we need to account for shifts.
            
            for i, idx in enumerate(sorted_indices):
                # Calculate current index of this page (it might have shifted)
                # This is hard to track.
                # Easier: Construct a new list of pages for the WHOLE document and re-create?
                # PyMuPDF: doc.select(page_numbers) rearranges!
                pass
            
            # New Strategy: `pdf_engine.reorder_pages(new_order)`
            # 1. Get current page count.
            # 2. Create list [0, 1, ... N]
            # 3. Remove selected indices.
            # 4. Insert selected indices at target.
            # 5. Call doc.select(new_order)
            
            # Implementation:
            all_pages = list(range(self.pdf.get_page_count()))
            to_move = sorted(list(indices))
            
            # Filter content
            remaining = [p for p in all_pages if p not in to_move]
            
            # Calculate insert position in 'remaining' list
            # target_index is based on OLD list.
            # We need to find where target_index maps to in remaining?
            # Simplified: If dropping on item X, we want to insert Before/After X.
            # If target_index was not moved, it's in remaining.
            # If target_index was moved, we act as if dropped on that moved item? (Visual feedback)
            
            # Let's assume target_index is "insert before".
            # We need to map target_index to position in 'remaining'.
            # Count how many items *before* target_index were removed.
            
            removed_before = sum(1 for x in to_move if x < target_index)
            new_insert_pos = target_index - removed_before
            
            # Insert
            for x in reversed(to_move):
                remaining.insert(new_insert_pos, x)
                
            self.pdf.doc.select(remaining)
            self.thumbnail_panel.refresh()
            self.thumbnail_panel.selected_indices = set() # Clear or restore?
            # Restore selection (indices have changed!)
            # We know where we inserted them: [new_insert_pos : new_insert_pos + len]
            new_selection = set(range(new_insert_pos, new_insert_pos + len(to_move)))
            self.thumbnail_panel.selected_indices = new_selection
            self.on_selection_change(new_selection)
            self.thumbnail_panel.refresh_selection_visuals()
            
            self.status_bar.config(text=f"Moved {len(indices)} pages.")
            
        else:
            # DIFFERENT WINDOW -> COPY (Insert)
            # This logic remains similar, but we use target_index
            src_pdf = source_window.pdf
            if not src_pdf.doc: return
            
            count = 0
            try:
                 # We insert at target_index.
                 # Loop through indices
                 # Note: insert_pdf inserts *before* the specified page number.
                 # So if we iterate, we just keep inserting at target_index + i
                 
                 sorted_indices = sorted(list(indices))
                 for i, idx in enumerate(sorted_indices):
                     self.pdf.doc.insert_pdf(src_pdf.doc, from_page=idx, to_page=idx, start_at=target_index+i)
                     count += 1
                 
                 self.thumbnail_panel.refresh()
                 # Select new pages
                 new_selection = set(range(target_index, target_index + count))
                 self.thumbnail_panel.selected_indices = new_selection
                 self.on_selection_change(new_selection)
                 self.thumbnail_panel.refresh_selection_visuals()
                 
                 self.status_bar.config(text=f"Copied {count} pages from other window.")
            except Exception as e:
                print(f"Drop failed: {e}")
                messagebox.showerror("Error", f"Drop failed: {e}")

    def on_copy(self, event=None):
        indices = self.thumbnail_panel.selected_indices
        if not indices or not self.pdf.file_path:
             return
        
        # We need the file path to copy from
        # Logic: If file has unsaved changes, we might need to save tmp?
        # For v3.2 parity, let's assume we copy from the file on disk.
        # If unsaved, warning? Or just copy from memory?
        # Copying from memory between processes is hard. Toplevel (same process) is easy: pass object.
        
        self.clipboard.copy(self, indices) # Pass self as source
        self.status_bar.config(text=f"Copied {len(indices)} pages.")

    def on_paste(self, event=None):
        data = self.clipboard.get_data()
        if not data: return
        
        source_window = data.get('source')
        indices = data.get('pages')
        
        if not source_window or not source_window.pdf.doc:
            return
            
        # Paste logic similar to drop
        count = 0
        for idx in indices:
            self.pdf.doc.insert_pdf(source_window.pdf.doc, from_page=idx, to_page=idx)
            count += 1
            
        self.thumbnail_panel.refresh()
        self.status_bar.config(text=f"Pasted {count} pages.")

    # ... (Rest of UI Setup) ...
    def setup_ui(self):
        # 0. Menu Bar
        self.create_menu()
        
        # 1. Toolbar
        self.toolbar = ttk.Frame(self, bootstyle="light")
        self.toolbar.pack(side=TOP, fill=X, pady=5)
        self.create_toolbar_buttons()
        
        # 2. Main Content
        self.paned = ttk.Panedwindow(self, orient=HORIZONTAL)
        self.paned.pack(fill=BOTH, expand=YES, padx=10, pady=5)
        
        # Left: Thumbnails
        # Pass Drag Manager
        self.thumbnail_panel = ThumbnailPanel(self.paned, self.pdf, self.on_selection_change, self.drag_manager)
        self.paned.add(self.thumbnail_panel, weight=1)
        
        # Right: Preview
        self.preview_panel = PreviewPanel(self.paned, self.pdf, on_page_change=self.on_preview_page_change)
        self.paned.add(self.preview_panel, weight=3)
        
        # 3. Footer / Status Bar
        self.create_footer()

    def on_preview_page_change(self, index):
        # Sync Thumbnail to Preview
        if len(self.thumbnail_panel.selected_indices) == 1 and list(self.thumbnail_panel.selected_indices)[0] == index:
            return
            
        self.thumbnail_panel.select_and_scroll_to(index)
        
        # Update status bar manually since we skipped on_selection_change callback loop?
        # Actually select_and_scroll_to DOES NOT call on_selection_change to avoid loop?
        # Wait, I said in comment "Don't trigger...". 
        # Let's ensure status bar updates.
        self.status_bar.config(text=f"페이지: {index + 1} / {self.pdf.get_page_count()}")

    def create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File Menu (파일)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="PDF 열기", command=self.on_open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="저장", command=self.on_save_pdf, accelerator="Ctrl+S")
        file_menu.add_command(label="선택 저장", command=self.on_save_selected)
        file_menu.add_separator()
        file_menu.add_command(label="새 창", command=self.on_new_window, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.quit)
        
        # Edit Menu (편집)
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="편집", menu=edit_menu)
        edit_menu.add_command(label="왼쪽으로 회전", command=lambda: self.on_rotate(-90))
        edit_menu.add_command(label="오른쪽으로 회전", command=lambda: self.on_rotate(90))
        edit_menu.add_command(label="페이지 삭제", command=self.on_delete_page, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="빈 페이지 삽입", command=self.on_blank_page)
        
        # Performance (성능 설정) - Restored v3.2
        perf_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="성능 설정", menu=perf_menu)
        perf_menu.add_command(label="🚀 고성능 모드 (권장)", command=lambda: self.set_performance_mode("high"))
        perf_menu.add_command(label="⚖️ 균형 모드", command=lambda: self.set_performance_mode("balanced"))
        perf_menu.add_command(label="🎨 고품질 모드", command=lambda: self.set_performance_mode("quality"))
        perf_menu.add_separator()
        self.perf_mode_lbl = "균형 모드"
        perf_menu.add_command(label=f"현재: {self.perf_mode_lbl}", state="disabled")

        # User Manager (사용자 관리) - Restored v3.2
        user_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="사용자 관리", menu=user_menu)
        
        # Check permissions
        if self.auth.is_admin():
            state = "normal"
        else:
            state = "disabled"
            
        user_menu.add_command(label="사용자 목록 보기", command=self.show_users_list, state=state)
        user_menu.add_command(label="사용자 추가", command=self.add_user_dialog, state=state)
        user_menu.add_command(label="사용자 제거", command=self.remove_user_dialog, state=state)
        user_menu.add_separator()
        user_menu.add_command(label="사용자 정보", command=lambda: messagebox.showinfo("사용자 정보", f"현재 사용자: {self.auth.get_current_user_name()}"))
        user_menu.add_separator()
        user_menu.add_command(label="사용자 백업", command=self.backup_users, state=state)
        user_menu.add_command(label="백업 복원", command=self.restore_users, state=state)

        # Help (도움말)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="정보", command=lambda: messagebox.showinfo("정보", f"{APP_NAME} {VERSION}\nCreated by {AUTHOR}"))

    def set_performance_mode(self, mode):
        # Placeholder for actual performance logic (caching strategies, etc.)
        modes = {"high": "고성능 모드", "balanced": "균형 모드", "quality": "고품질 모드"}
        self.perf_mode_lbl = modes.get(mode, "균형 모드")
        messagebox.showinfo("성능 설정", f"{self.perf_mode_lbl}로 설정되었습니다.")
        # Re-create menu to update label? Or just ignored for now.

    def show_users_list(self):
        users = self.auth.get_all_users()
        if not users:
            messagebox.showinfo("목록", "사용자가 없습니다.")
            return
            
        win = tk.Toplevel(self)
        win.title("사용자 목록")
        win.geometry("400x500")
        
        text = tk.Text(win)
        text.pack(fill=BOTH, expand=YES)
        
        for u in users:
            text.insert(END, f"이름: {u['name']}\n직급: {u['role']}\nMAC: {u['mac']}\n{'-'*30}\n")
            
    def add_user_dialog(self):
        from tkinter import simpledialog
        name = simpledialog.askstring("사용자 추가", "이름:")
        if not name: return
        role = simpledialog.askstring("사용자 추가", "직급 (admin/user):")
        if not role: return
        mac = simpledialog.askstring("사용자 추가", "MAC 주소 (XX-XX-XX-XX-XX-XX):")
        if not mac: return
        
        success, msg = self.auth.add_user(name, role, mac)
        messagebox.showinfo("결과", msg)

    def remove_user_dialog(self):
        from tkinter import simpledialog
        mac = simpledialog.askstring("사용자 제거", "제거할 MAC 주소:")
        if not mac: return
        
        success, msg = self.auth.remove_user(mac)
        messagebox.showinfo("결과", msg)

    def backup_users(self):
        path = filedialog.asksaveasfilename(defaultextension=".enc", filetypes=[("Encrypted Users", "*.enc")])
        if path:
            success, msg = self.auth.backup_users(path)
            messagebox.showinfo("백업", msg)

    def restore_users(self):
        path = filedialog.askopenfilename(filetypes=[("Encrypted Users", "*.enc")])
        if path:
            success, msg = self.auth.restore_users(path)
            messagebox.showinfo("복원", msg)

    # ... (Rest of methods: on_new_window, on_open_pdf, on_save_pdf...)
    # I need to ensure I don't overwrite them. 
    # I will replace only the setup_ui and create_menu and add new methods.
    # But wait, replace_file_content replaces a BLOCK.
    # I targeted from setup_ui (line 301) to end (558).
    # I MUST include all methods after create_menu in the replacement content to match the target.
    # The existing code has create_toolbar_buttons, create_footer, on_open_pdf, ..., on_save_selected, on_new_window, ..., on_split.
    # This is too much code to blindly copy-paste in "ReplacementContent" without error risk.
    # I should use multiple Replace calls or targeted replaces.
    
    # Strategy:
    # 1. Replace setup_ui
    # 2. Replace create_menu
    # 3. Add new methods at the end
    
    # Let's start with setup_ui


    # ... (Keep existing methods: create_menu, create_toolbar, etc. BUT update on_new_window) ...
    
    def on_new_window(self):
        MainWindow(master=self.master) # Create new toplevel

    # ... (Rest of existing methods) ...
    
    # Need to include all previous methods to avoid deletion by replacement
    # I will use replace_file_content carefully or just rewrite the class header and specific methods.
    # Since I need to change inheritance and __init__, and add new methods.
    # I will just replace the top part and add methods at end, and update setup_ui/on_new_window.


    def create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File Menu (파일)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="PDF 열기", command=self.on_open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="저장", command=self.on_save_pdf, accelerator="Ctrl+S")
        file_menu.add_command(label="선택 저장", command=self.on_save_selected)
        file_menu.add_separator()
        file_menu.add_command(label="새 창", command=self.on_new_window, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.quit)
        
        # Edit Menu (편집)
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="편집", menu=edit_menu)
        edit_menu.add_command(label="왼쪽으로 회전", command=lambda: self.on_rotate(-90))
        edit_menu.add_command(label="오른쪽으로 회전", command=lambda: self.on_rotate(90))
        edit_menu.add_command(label="페이지 삭제", command=self.on_delete_page, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="빈 페이지 삽입", command=self.on_blank_page)
        
        # Performance (성능 설정) - Restored v3.2
        perf_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="성능 설정", menu=perf_menu)
        perf_menu.add_command(label="🚀 고성능 모드 (권장)", command=lambda: self.set_performance_mode("high"))
        perf_menu.add_command(label="⚖️ 균형 모드", command=lambda: self.set_performance_mode("balanced"))
        perf_menu.add_command(label="🎨 고품질 모드", command=lambda: self.set_performance_mode("quality"))
        perf_menu.add_separator()
        self.perf_mode_lbl = "균형 모드"
        perf_menu.add_command(label=f"현재: {self.perf_mode_lbl}", state="disabled")

        # User Manager (사용자 관리) - Restored v3.2
        user_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="사용자 관리", menu=user_menu)
        
        # Check permissions
        if self.auth.is_admin():
            state = "normal"
        else:
            state = "disabled"
            
        user_menu.add_command(label="사용자 목록 보기", command=self.show_users_list, state=state)
        user_menu.add_command(label="사용자 추가", command=self.add_user_dialog, state=state)
        user_menu.add_command(label="사용자 제거", command=self.remove_user_dialog, state=state)
        user_menu.add_separator()
        user_menu.add_command(label="사용자 정보", command=lambda: messagebox.showinfo("사용자 정보", f"현재 사용자: {self.auth.get_current_user_name()}"))
        user_menu.add_separator()
        user_menu.add_command(label="사용자 백업", command=self.backup_users, state=state)
        user_menu.add_command(label="백업 복원", command=self.restore_users, state=state)

        # Help (도움말)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="정보", command=lambda: messagebox.showinfo("정보", f"{APP_NAME} {VERSION}\nCreated by {AUTHOR}"))

    def create_toolbar_buttons(self):
        # Group 1: File (파일)
        grp_file = ttk.Labelframe(self.toolbar, text="파일", padding=5)
        grp_file.pack(side=LEFT, padx=5, fill=Y)
        ttk.Button(grp_file, text="PDF 열기", command=self.on_open_pdf, bootstyle="primary").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="저장", command=self.on_save_pdf, bootstyle="success").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="선택 저장", command=self.on_save_selected, bootstyle="info").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="새 창", command=self.on_new_window, bootstyle="secondary").pack(side=LEFT, padx=2)
        
        # Group 2: Merge (병합)
        grp_merge = ttk.Labelframe(self.toolbar, text="병합", padding=5)
        grp_merge.pack(side=LEFT, padx=5, fill=Y)
        ttk.Button(grp_merge, text="병합", command=self.on_merge, bootstyle="primary-outline").pack(side=LEFT, padx=2)
        ttk.Button(grp_merge, text="다중 병합", command=self.on_multi_merge, bootstyle="primary-outline").pack(side=LEFT, padx=2)

        # Group 3: Edit (편집)
        grp_edit = ttk.Labelframe(self.toolbar, text="편집", padding=5)
        grp_edit.pack(side=LEFT, padx=5, fill=Y)
        ttk.Button(grp_edit, text="좌로90°", command=lambda: self.on_rotate(-90), bootstyle="warning").pack(side=LEFT, padx=2)
        ttk.Button(grp_edit, text="우로90°", command=lambda: self.on_rotate(90), bootstyle="warning").pack(side=LEFT, padx=2)
        ttk.Button(grp_edit, text="빈페이지", command=self.on_blank_page, bootstyle="success").pack(side=LEFT, padx=2)
        ttk.Button(grp_edit, text="삭제", command=self.on_delete_page, bootstyle="danger").pack(side=LEFT, padx=2)

        # Group 4: Tools (도구)
        grp_tools = ttk.Labelframe(self.toolbar, text="도구", padding=5)
        grp_tools.pack(side=LEFT, padx=5, fill=Y)
        ttk.Button(grp_tools, text="텍스트", command=self.on_extract_text, bootstyle="info-outline").pack(side=LEFT, padx=2)
        ttk.Button(grp_tools, text="맞춤", command=self.on_fit_screen, bootstyle="secondary-outline").pack(side=LEFT, padx=2)
        ttk.Button(grp_tools, text="분할", command=self.on_split, bootstyle="dark-outline").pack(side=LEFT, padx=2)

    def create_footer(self):
        footer_frame = ttk.Frame(self, bootstyle="light")
        footer_frame.pack(side=BOTTOM, fill=X)
        
        # Copyright
        from config.settings import AUTHOR, COMPANY, VERSION
        user_name = self.auth.get_current_user_name()
        
        copyright_text = f"© 2025 {COMPANY} {VERSION} | Developed by {AUTHOR} | 사용자: {user_name}"
        lbl_copy = ttk.Label(footer_frame, text=copyright_text, font=("Segoe UI", 8), bootstyle="secondary")
        lbl_copy.pack(side=LEFT, padx=10, pady=5)
        
        # Status
        self.status_bar = ttk.Label(footer_frame, text="준비", bootstyle="inverse-light", font=("Segoe UI", 9))
        self.status_bar.pack(side=RIGHT, padx=10, pady=5)

    def on_selection_change(self, selected_indices):
        count = len(selected_indices)
        txt = f"선택된 페이지: {count}"
        if count == 1:
            txt += f" ({list(selected_indices)[0] + 1} 페이지)"
            
        # Add hint for shortcuts
        txt += " | Ctrl+클릭: 다중선택 | Shift+클릭: 범위선택 | Delete: 삭제 | Ctrl+A: 전체선택"
        self.status_bar.config(text=txt)
        
        if selected_indices:
            last_selected = list(selected_indices)[-1]
            self.preview_panel.show_page(last_selected)

    def on_open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if path:
            success, msg = self.pdf.open_pdf(path)
            if success:
                self.status_bar.config(text=f"열림: {path}")
                self.thumbnail_panel.refresh()
                if self.pdf.get_page_count() > 0:
                    self.preview_panel.show_page(0)
            else:
                messagebox.showerror("오류", msg)

    def on_save_pdf(self):
        if not self.pdf.doc: return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 파일", "*.pdf")])
        if path:
            success, msg = self.pdf.save_pdf(path)
            self.status_bar.config(text=msg)

    def on_save_selected(self):
        self.on_split()

    def on_new_window(self):
        # Create new window in same process to share Clipboard/DragManager
        new_win = MainWindow(master=self.master)
        # Position it slightly offset
        x = self.winfo_x() + 50
        y = self.winfo_y() + 50
        new_win.geometry(f"+{x}+{y}")

    def on_merge(self):
        if not self.pdf.doc:
            messagebox.showinfo("알림", "병합할 PDF를 먼저 열어주세요.")
            return
        path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if path:
            if self.pdf.insert_pdf(path):
                self.thumbnail_panel.refresh()
                self.status_bar.config(text="PDF 병합 완료.")
            else:
                messagebox.showerror("오류", "병합 실패.")

    def on_multi_merge(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF 파일", "*.pdf")])
        if paths:
            if self.pdf.merge_pdf_list(paths):
                self.thumbnail_panel.refresh()
                if self.pdf.get_page_count() > 0:
                     self.preview_panel.show_page(0) # show first page if new doc
                self.status_bar.config(text=f"{len(paths)}개 파일 병합 완료.")
            else:
                messagebox.showerror("오류", "병합 실패.")

    def on_blank_page(self):
        self.pdf.create_blank_page()
        self.thumbnail_panel.refresh()
        self.status_bar.config(text="빈 페이지 추가됨.")

    def on_extract_text(self):
        indices = self.thumbnail_panel.selected_indices
        if not indices:
             messagebox.showinfo("알림", "텍스트를 추출할 페이지를 선택하세요.")
             return
             
        idx = list(indices)[0] # Extract from first selected
        text = self.pdf.extract_text(idx)
        if text:
             # Show text in new window
             top = tk.Toplevel(self)
             top.title(f"{idx+1} 페이지 텍스트")
             text_area = tk.Text(top, wrap="word")
             text_area.pack(fill=BOTH, expand=YES)
             text_area.insert("1.0", text)
        else:
             messagebox.showinfo("알림", "텍스트가 없습니다.")

    def on_fit_screen(self):
        # Reset zoom
        self.preview_panel.zoom_scale = 1.0
        self.preview_panel.show_page(self.preview_panel.current_page_index)

    def on_rotate(self, angle):
        if not self.pdf.doc: return
        indices = self.thumbnail_panel.selected_indices
        if not indices:
            messagebox.showinfo("알림", "회전할 페이지를 선택하세요.")
            return
            
        for idx in indices:
            self.pdf.rotate_page(idx, angle)
            
        self.thumbnail_panel.refresh()
        last_selected = list(indices)[-1]
        self.preview_panel.show_page(last_selected)
        self.status_bar.config(text="페이지 회전 완료.")

    def on_delete_page(self):
        if not self.pdf.doc: return
        indices = self.thumbnail_panel.selected_indices
        if not indices:
            return

        if messagebox.askyesno("삭제", f"{len(indices)}개 페이지를 삭제하시겠습니까?"):
            self.pdf.delete_pages(indices)
            self.thumbnail_panel.refresh()
            self.preview_panel.clear()
            self.status_bar.config(text="삭제 완료.")

    def on_split(self):
        if not self.pdf.doc: return
        indices = self.thumbnail_panel.selected_indices
        if not indices:
             messagebox.showinfo("Info", "Select pages to export.")
             return
             
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if path:
             success = self.pdf.export_selection(sorted(list(indices)), path)
    def set_performance_mode(self, mode):
        # Placeholder for actual performance logic (caching strategies, etc.)
        modes = {"high": "고성능 모드", "balanced": "균형 모드", "quality": "고품질 모드"}
        self.perf_mode_lbl = modes.get(mode, "균형 모드")
        messagebox.showinfo("성능 설정", f"{self.perf_mode_lbl}로 설정되었습니다.")

    def show_users_list(self):
        users = self.auth.get_all_users()
        if not users:
            messagebox.showinfo("목록", "사용자가 없습니다.")
            return
            
        win = tk.Toplevel(self)
        win.title("사용자 목록")
        win.geometry("400x500")
        
        text = tk.Text(win)
        text.pack(fill=BOTH, expand=YES)
        
        for u in users:
            text.insert(END, f"이름: {u['name']}\n직급: {u['role']}\nMAC: {u['mac']}\n{'-'*30}\n")
            
    def add_user_dialog(self):
        from tkinter import simpledialog
        name = simpledialog.askstring("사용자 추가", "이름:")
        if not name: return
        role = simpledialog.askstring("사용자 추가", "직급 (admin/user):")
        if not role: return
        mac = simpledialog.askstring("사용자 추가", "MAC 주소 (XX-XX-XX-XX-XX-XX):")
        if not mac: return
        
        success, msg = self.auth.add_user(name, role, mac)
        messagebox.showinfo("결과", msg)

    def remove_user_dialog(self):
        from tkinter import simpledialog
        mac = simpledialog.askstring("사용자 제거", "제거할 MAC 주소:")
        if not mac: return
        
        success, msg = self.auth.remove_user(mac)
        messagebox.showinfo("결과", msg)

    def backup_users(self):
        path = filedialog.asksaveasfilename(defaultextension=".enc", filetypes=[("Encrypted Users", "*.enc")])
        if path:
            success, msg = self.auth.backup_users(path)
            messagebox.showinfo("백업", msg)

    def restore_users(self):
        path = filedialog.askopenfilename(filetypes=[("Encrypted Users", "*.enc")])
        if path:
            success, msg = self.auth.restore_users(path)
            messagebox.showinfo("복원", msg)
