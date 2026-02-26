import os
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
from tkinterdnd2 import TkinterDnD
class MainWindow(ttk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        
        # Apply theme explicitly since we might not use ttk.Window as root
        style = ttk.Style(theme=THEME_NAME)
        
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
            self.show_auth_failure_dialog(message)
            self.destroy()
            if self.master:
                self.master.destroy()
            return
        self.title(f"{APP_NAME} {VERSION} - {self.auth.get_current_user_name()}")
        
        self.setup_ui()
        
        # Bind Keys
        self.bind("<Control-o>", lambda e: self.on_open_pdf())
        self.bind("<Control-s>", lambda e: self.on_save_pdf())
        self.bind("<Control-S>", lambda e: self.on_save_as_file()) # Shift+S usually maps to Capital S
        self.bind("<Control-Shift-s>", lambda e: self.on_save_as_file()) # Explicit just in case
        self.bind("<Control-c>", self.on_copy)
        self.bind("<Control-v>", self.on_paste)
        
        # Global Mouse Events for Drag
        self.bind("<Motion>", self.on_global_motion)
        self.bind("<ButtonRelease-1>", self.on_global_release)
        
        self.setup_global_binds()
    def show_auth_failure_dialog(self, message):
        """Shows authentication failure dialog with Copy MAC button."""
        # Extract MAC if present in message
        mac = ""
        import re
        mac_match = re.search(r'([0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2}-[0-9A-F]{2})', message, re.IGNORECASE)
        if mac_match:
            mac = mac_match.group(1).upper()
            
        dialog = tk.Toplevel(self)
        dialog.title("인증 실패")
        dialog.geometry("400x250")
        
        # Center
        x = self.winfo_screenwidth() // 2 - 200
        y = self.winfo_screenheight() // 2 - 125
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="⛔ 인증되지 않은 사용자입니다.", font=("맑은 고딕", 12, "bold"), bootstyle="danger").pack(pady=20)
        
        msg_frame = ttk.Frame(dialog, padding=10)
        msg_frame.pack(fill=BOTH, expand=YES)
        ttk.Label(msg_frame, text=message, justify="center").pack()
        
        btn_frame = ttk.Frame(dialog, padding=20)
        btn_frame.pack(fill=X)
        
        def copy_mac():
            self.clipboard_clear()
            self.clipboard_append(mac)
            self.update()
            messagebox.showinfo("복사 완료", "MAC 주소가 클립보드에 복사되었습니다.\n관리자에게 전달해주세요.")
            
        if mac:
            ttk.Button(btn_frame, text="MAC 주소 복사", command=copy_mac, bootstyle="info").pack(side=LEFT, expand=YES, padx=5)
            
        ttk.Button(btn_frame, text="종료", command=dialog.destroy, bootstyle="secondary").pack(side=RIGHT, expand=YES, padx=5)
        
        # Wait window
        self.wait_window(dialog)
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
        self.bind("<Control-z>", self.on_undo)
        self.bind("<Control-Z>", self.on_undo)
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
            self.pdf.push_undo_state()
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

    def on_undo(self, event=None):
        if not self.pdf.doc: return
        if hasattr(self.pdf, 'undo_stack') and len(self.pdf.undo_stack) > 0:
            if self.pdf.undo():
                self.thumbnail_panel.refresh()
                self.on_preview_page_change(0)
                self.status_bar.config(text="실행 취소 완료")
            else:
                self.status_bar.config(text="실행 취소에 실패했습니다.")
        else:
            self.status_bar.config(text="더 이상 실행 취소할 항목이 없습니다.")

    def on_external_drop(self, source_window, indices, x, y):
        # Get drop target index from ThumbnailPanel
        target_index = self.thumbnail_panel.get_drop_index_at(x, y)
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
            
            # For now, let's just try moving the FIRST page of selection to target,
            # then others... this is UX heavy.
            # Simplified: Move them one by one, carefully adjusting target?
            
            # Let's use a simple loop for now, iterate sorted indices.
            # If moving Down (target > current): Move reversed?
            # If moving Up (target < current): Move normal?
            
            # Actually, `pdf_engine.move_page` implementation checks `from` and `to`.
            # Reorder Logic
            self.pdf.push_undo_state()
            total_pages = self.pdf.get_page_count()
            all_pages_list = list(range(total_pages))
            moving_indices = sorted(list(indices))
            # Remove indices from list
            for idx in reversed(moving_indices): # Reverse to avoid index shifting
                if idx < len(all_pages_list):
                    all_pages_list.pop(idx)
            
            # Calculate where to insert
            # We need to account that removing items shifts indices.
            # target_index is based on OLD indices.
            
            # Count how many items BEFORE target were removed
            removed_before = sum(1 for idx in moving_indices if idx < target_index)
            insert_pos = max(0, target_index - removed_before)
            
            # Insert moving items
            for idx in reversed(moving_indices):
                all_pages_list.insert(insert_pos, idx)
                
            # Apply
            self.pdf.doc.select(all_pages_list)
            self.thumbnail_panel.refresh()
            
            # Reselect moved items (they are now at insert_pos)
            new_selection = set(range(insert_pos, insert_pos + len(moving_indices)))
            self.thumbnail_panel.selected_indices = new_selection
            self.thumbnail_panel.refresh_selection_visuals()
            self.on_preview_page_change(insert_pos)
            
        else:
            # DIFFERENT WINDOW -> COPY (Insert)
            # This logic remains similar, but we use target_index
            src_pdf = source_window.pdf
            if not src_pdf.doc: return
            
            self.pdf.push_undo_state()
            
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

    def on_drag_hover(self, source_window, indices, x, y):
        """Pass drag hover events down to thumbnail panel to draw insertion guides."""
        if hasattr(self, 'thumbnail_panel'):
            self.thumbnail_panel.draw_drag_guide(x, y)
            
    def clear_drag_guide(self):
        """Clear drag guide when drag stops or exits window bounds."""
        if hasattr(self, 'thumbnail_panel'):
            self.thumbnail_panel.clear_drag_guide()

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
    def export_users_local(self):
        """Export users data to a local unencrypted file (Backup)."""
        users_data = self.auth.load_users()
        if not users_data:
             messagebox.showerror("오류", "사용자 데이터가 없습니다.")
             return
        import json
        
        types = [
            ("JSON 파일", "*.json"),
            ("텍스트 파일", "*.txt")
        ]
        
        path = filedialog.asksaveasfilename(
            title="사용자 목록 저장 (백업 - 비암호화)",
            filetypes=types,
            defaultextension=".json",
            initialfile="users_backup.json"
        )
        
        if not path: return
        
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(users_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("완료", f"사용자 목록이 저장되었습니다.\n{path}")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")
    def show_usage_dialog(self):
        """Shows usage instructions."""
        msg = """
[단축키 안내]
- PDF 열기: Ctrl + O
- 저장: Ctrl + S
- 다른 이름으로 저장: Ctrl + Shift + S
- 페이지 삭제: Delete 키
- 전체 선택: Ctrl + A
- 다중 선택: Ctrl + 클릭
- 범위 선택: Shift + 클릭
[기능 안내]
- 썸네일 드래그 앤 드롭으로 페이지 순서를 변경할 수 있습니다.
- 썸네일 창에서 우클릭 메뉴를 사용할 수 있습니다.
- 사용자 관리 메뉴에서 팀원 권한을 관리하세요.
        """
        messagebox.showinfo("사용법", msg.strip())
    def create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File Menu (파일)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="PDF 열기", command=self.on_open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="저장", command=self.on_save_pdf, accelerator="Ctrl+S")
        file_menu.add_command(label="다른 이름으로 저장", command=self.on_save_as_file, accelerator="Ctrl+Shift+S")
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
        
        # User Manager (사용자 관리)
        user_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="사용자 관리", menu=user_menu)
        
        # Check permissions
        if self.auth.is_admin():
            state = "normal"
        else:
            state = "disabled"
            
        user_menu.add_command(label="사용자 관리 (수정/삭제)", command=self.manage_users_dialog, state=state)
        user_menu.add_command(label="사용자 추가", command=self.add_user_dialog, state=state)
        user_menu.add_separator()
        user_menu.add_command(label="사용자 목록 저장 (백업)", command=self.export_users_local, state=state)
        user_menu.add_command(label="Gist용 데이터 내보내기", command=self.export_gist_data, state=state)
        
        # Help (도움말)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_usage_dialog)
        help_menu.add_separator()
        help_menu.add_command(label="정보", command=lambda: messagebox.showinfo("정보", f"{APP_NAME} {VERSION}\nCreated by {AUTHOR}", parent=self))
    def on_open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if path:
            success, msg = self.pdf.open_pdf(path)
            if success:
                self.status_bar.config(text=f"열림: {path}")
                self.thumbnail_panel.set_filename(os.path.basename(path))
                
                # Update Title
                filename = os.path.basename(path)
                user = self.auth.get_current_user_name()
                self.title(f"{filename} - {APP_NAME} {VERSION} [{user}]")
                
                # Delay refresh to ensure UI is ready
                self.after(200, self._refresh_on_open)
            else:
                messagebox.showerror("오류", msg)
    def create_toolbar_buttons(self):
        # Group 1: File (파일)
        grp_file = ttk.Labelframe(self.toolbar, text="파일", padding=5)
        grp_file.pack(side=LEFT, padx=5, fill=Y)
        ttk.Button(grp_file, text="PDF 열기", command=self.on_open_pdf, bootstyle="primary").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="저장", command=self.on_save_pdf, bootstyle="success").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="선택 저장", command=self.on_save_selected, bootstyle="info").pack(side=LEFT, padx=2)
        ttk.Button(grp_file, text="이미지로 저장", command=self.on_export_to_image, bootstyle="warning").pack(side=LEFT, padx=2)
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
            

    def update_title(self):
        filename = "제목 없음"
        if self.pdf.file_path:
            filename = os.path.basename(self.pdf.file_path)
        
        user = self.auth.get_current_user_name()
        # Requested Format: [Filename] - AppName
        title_str = f"[{filename}] - {APP_NAME} {VERSION} [{user}]"
        self.title(title_str)
        
        # Also update root title if it exists (for Taskbar)
        if self.master:
            self.master.title(title_str)
    def create_footer(self):
        footer_frame = ttk.Frame(self, bootstyle="light")
        footer_frame.pack(side=BOTTOM, fill=X)
        
        # Copyright
        from config.settings import AUTHOR, COMPANY, VERSION
        user = self.auth.get_current_user_name()
        
        # Filename in footer too? No, title is enough.
        
        copyright_text = f"© 2025 {COMPANY} {VERSION} | Developed by {AUTHOR} | 사용자: {user}"
        lbl_copy = ttk.Label(footer_frame, text=copyright_text, font=("맑은 고딕", 8), bootstyle="secondary")
        lbl_copy.pack(side=LEFT, padx=10, pady=5)
        
        # Status
        self.status_bar = ttk.Label(footer_frame, text="준비", bootstyle="inverse-light", font=("맑은 고딕", 9))
        self.status_bar.pack(side=RIGHT, padx=10, pady=5)
    def on_selection_change(self, selected_indices):
        count = len(selected_indices)
        if count == 0:
            txt = "선택된 페이지: 없음"
        else:
            sorted_indices = sorted(list(selected_indices))
            pages = [str(i + 1) for i in sorted_indices]
            
            # If too many, truncate? User example showed 3 items. 
            # Let's show all if reasonable, or truncate if very long.
            if len(pages) > 10:
                pages_str = ", ".join(pages[:10]) + "..."
            else:
                pages_str = ", ".join(pages)
                
            txt = f"선택된 페이지: {count}개 ({pages_str})"
            
        # Add hint for shortcuts
        txt += " | Ctrl+클릭: 다중선택 | Shift+클릭: 범위선택 | Delete: 삭제 | Ctrl+A: 전체선택"
        self.status_bar.config(text=txt)
        
        if selected_indices:
            last_selected = list(selected_indices)[-1]
            self.preview_panel.show_page(last_selected, from_thumbnail=True)
    def on_open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if path:
            success, msg = self.pdf.open_pdf(path)
            if success:
                self.status_bar.config(text=f"열림: {path}")
                self.thumbnail_panel.set_filename(os.path.basename(path))
                
                # Delay refresh to ensure UI is ready
                self.after(200, self._refresh_on_open)
            else:
                messagebox.showerror("오류", msg)
    def _refresh_on_open(self):
        """Called after delay"""
        self.update_idletasks()
        self.thumbnail_panel.refresh()
        if self.pdf.get_page_count() > 0:
            self.preview_panel.show_page(0)
            # Try to fit to window if panel is ready
            self.preview_panel.fit_to_window()
            
    def on_save_pdf(self):
        """Quick Save (Overwrite)"""
        # If file exists, overwrite. Else Save As.
        if self.pdf.file_path and os.path.exists(self.pdf.file_path):
            success, msg = self.pdf.save_pdf(self.pdf.file_path)
            self.status_bar.config(text=msg)
            # Show simple feedback?
            # messagebox.showinfo("저장", "저장되었습니다.")
        else:
            self.on_save_as_file()
    def on_save_as_file(self):
        """Advanced Save As (PDF, JPG, PNG) with Page Selection."""
        if not self.pdf.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        # 1. Show Export Dialog
        dialog = ExportDialog(self, self.pdf.get_page_count(), len(self.thumbnail_panel.selected_indices))
        self.wait_window(dialog)
        
        if not dialog.result:
            return # Cancelled
            
        fmt = dialog.result['format'] # 'pdf', 'jpg', 'png'
        range_type = dialog.result['range'] # 'all', 'selected', 'custom'
        custom_pages = dialog.result['custom_pages'] # string
        
        # 2. Determine Page Indices
        page_indices = []
        total_pages = self.pdf.get_page_count()
        
        if range_type == 'all':
            page_indices = list(range(total_pages))
        elif range_type == 'selected':
            page_indices = sorted(list(self.thumbnail_panel.selected_indices))
            if not page_indices:
                messagebox.showwarning("경고", "선택된 페이지가 없습니다.")
                return
        elif range_type == 'custom':
            # Parse "1, 3-5, 8"
            try:
                indices = set()
                parts = custom_pages.split(',')
                for p in parts:
                    p = p.strip()
                    if '-' in p:
                        start, end = map(int, p.split('-'))
                        # User input 1-based, convert to 0-based
                        for i in range(start, end + 1):
                            indices.add(i - 1)
                    else:
                        indices.add(int(p) - 1)
                
                # Filter valid
                page_indices = sorted([i for i in indices if 0 <= i < total_pages])
                
                if not page_indices:
                     messagebox.showwarning("오류", "유효한 페이지 번호가 없습니다.")
                     return
            except ValueError:
                messagebox.showwarning("오류", "페이지 형식이 올바르지 않습니다.\n예: 1, 3-5")
                return
        # 3. File Save Dialog
        filetypes = []
        ext = ""
        if fmt == 'pdf':
            filetypes = [("PDF 파일", "*.pdf")]
            ext = ".pdf"
        elif fmt == 'jpg':
            filetypes = [("JPEG 이미지", "*.jpg")]
            ext = ".jpg"
        elif fmt == 'png':
            filetypes = [("PNG 이미지", "*.png")]
            ext = ".png"
            
        initial = "document"
        if self.pdf.file_path:
            initial = os.path.splitext(os.path.basename(self.pdf.file_path))[0]
            
        path = filedialog.asksaveasfilename(
            title=f"다른 이름으로 저장 ({fmt.upper()})",
            defaultextension=ext,
            filetypes=filetypes,
            initialfile=initial
        )
        
        if not path: return
        
        # 4. Execute Save
        try:
            if fmt == 'pdf':
                if range_type == 'all':
                    # Full save (or just copy if simple?)
                    # Use save_subset even for all to ensure consistent clean save?
                    # Or use existing save_pdf for efficiency/metadata preservation?
                    # Existing save_pdf saves the opened document.
                    # If we used 'all', we clone the doc?
                    # Let's use save_subset for 'all' too if we want to "Export".
                    # But Save As usually implies keeping everything.
                    # save_subset creates NEW doc (loses bookmarks etc maybe).
                    # save_pdf saves CURRENT doc.
                    
                    if len(page_indices) == total_pages:
                         # Full Check
                         success, msg = self.pdf.save_pdf(path)
                    else:
                         success, msg = self.pdf.save_subset(page_indices, path)
                         
                    if success:
                         messagebox.showinfo("완료", "PDF 저장이 완료되었습니다.")
                    else:
                         messagebox.showerror("오류", msg)
                         
            else: # Image
                self.export_images(path, os.path.splitext(path)[1], page_indices)
                
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류 발생: {e}")
    def export_images(self, path, ext, page_indices):
        """Export specific pages as images."""
        try:
            base_path = os.path.splitext(path)[0]
            count = len(page_indices)
            saved_files = []
            
            self.status_bar.config(text="이미지 내보내는 중...")
            self.update()
            
            for i, page_idx in enumerate(page_indices):
                # Scale 4.0 ~ 288 DPI. High quality.
                img = self.pdf.get_page_image(page_idx, scale=4.0)
                
                # Naming strategy
                if count > 1:
                     filename = f"{base_path}_p{page_idx+1}{ext}"
                else:
                     filename = f"{base_path}{ext}"
                
                # Save with quality options
                if ext.lower() in ['.jpg', '.jpeg']:
                    img.save(filename, quality=95, subsampling=0)
                else:
                    img.save(filename)
                
                saved_files.append(filename)
            
            self.status_bar.config(text=f"{count}개 이미지 저장 완료.")
            messagebox.showinfo("완료", f"{count}개 파일이 저장되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"이미지 저장 실패: {e}")

    def on_export_to_image(self):
        """Quickly export selected (or all) pages to images."""
        if not self.pdf.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
            
        indices = sorted(list(self.thumbnail_panel.selected_indices))
        if not indices:
            if messagebox.askyesno("전체 내보내기", "선택된 페이지가 없습니다. 전체 페이지를 이미지(JPG)로 내보내시겠습니까?"):
                indices = list(range(self.pdf.get_page_count()))
            else:
                return
                
        folder = filedialog.askdirectory(title="이미지를 저장할 폴더 선택")
        if not folder: return
        
        base_name = "image_export"
        if self.pdf.file_path:
             base_name = os.path.splitext(os.path.basename(self.pdf.file_path))[0]
             
        path = os.path.join(folder, base_name + ".jpg") 
        self.export_images(path, ".jpg", indices)

    def on_save_selected(self):
        """Save currently selected pages as a new PDF."""
        # Shortcut to 'Save As PDF' with 'Selected' range?
        # Or distinct logic? 
        # Let's map it to the Export Dialog pre-filled?
        # Or just do it directly?
        # Previous implementation was self.on_split() which seemingly did something similar?
        # Let's restore as direct call to on_split if it exists, or implement using save_subset.
        
        # If I look at my previous read of line 402: "command=self.on_save_selected".
        # And line 592: "def on_save_selected(self): self.on_split()"
        
        # Checking if on_split exists...
        if hasattr(self, 'on_split'):
            self.on_split()
        else:
             # Fallback: Open Export Dialog with 'Selected' preset?
             # Or just show error?
             # Let's implement a simple version or redirect.
             if not self.thumbnail_panel.selected_indices:
                 messagebox.showwarning("경고", "선택된 페이지가 없습니다.")
                 return
             
             # Use new save_subset
             path = filedialog.asksaveasfilename(
                 title="선택된 페이지 저장",
                 defaultextension=".pdf",
                 filetypes=[("PDF 파일", "*.pdf")]
             )
             if path:
                 indices = sorted(list(self.thumbnail_panel.selected_indices))
                 success, msg = self.pdf.save_subset(indices, path)
                 if success:
                     messagebox.showinfo("완료", "선택된 페이지가 저장되었습니다.")
                 else:
                     messagebox.showerror("오류", msg)
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
        if not paths: return
        
        # Show Reordering Dialog
        dialog = MergeOrderingDialog(self, paths)
        self.wait_window(dialog)
        
        if not dialog.result: return
        
        ordered_paths = dialog.result
        
        if self.pdf.merge_pdf_list(ordered_paths):
            self.thumbnail_panel.refresh()
            if self.pdf.get_page_count() > 0:
                 self.preview_panel.show_page(0) 
                 self.preview_panel.fit_to_window()
            self.status_bar.config(text=f"{len(ordered_paths)}개 파일 병합 완료.")
            messagebox.showinfo("완료", "파일 병합이 완료되었습니다.")
        else:
            messagebox.showerror("오류", "병합 실패.")
    def on_blank_page(self):
        self.show_insert_blank_page_dialog()
    def show_insert_blank_page_dialog(self):
        """빈페이지 삽입 다이얼로그 표시"""
        if not self.pdf.doc:
            messagebox.showwarning("경고", "먼저 PDF를 열어주세요.")
            return
        
        # 새 창 생성
        dialog = tk.Toplevel(self)
        dialog.title("빈페이지 삽입")
        dialog.geometry("300x250")
        dialog.transient(self)
        dialog.grab_set()
        
        # 중앙 정렬
        x = self.winfo_x() + (self.winfo_width() // 2) - 150
        y = self.winfo_y() + (self.winfo_height() // 2) - 125
        dialog.geometry(f"+{x}+{y}")
        
        # 페이지 크기 선택
        ttk.Label(dialog, text="페이지 크기를 선택하세요:", font=("맑은 고딕", 10)).pack(pady=10)
        
        # A4 가로/세로, A3 가로/세로 버튼들
        ttk.Button(dialog, text="A4 가로 (297×210mm)", 
                 command=lambda: self.insert_blank_page("A4", "landscape", dialog)).pack(pady=5, fill=X, padx=20)
        
        ttk.Button(dialog, text="A4 세로 (210×297mm)", 
                 command=lambda: self.insert_blank_page("A4", "portrait", dialog)).pack(pady=5, fill=X, padx=20)
        
        ttk.Button(dialog, text="A3 가로 (420×297mm)", 
                 command=lambda: self.insert_blank_page("A3", "landscape", dialog)).pack(pady=5, fill=X, padx=20)
        
        ttk.Button(dialog, text="A3 세로 (297×420mm)", 
                 command=lambda: self.insert_blank_page("A3", "portrait", dialog)).pack(pady=5, fill=X, padx=20)
        
        # 취소 버튼
        ttk.Button(dialog, text="취소", command=dialog.destroy, bootstyle="secondary").pack(pady=10, fill=X, padx=20)
    def insert_blank_page(self, page_size, orientation, dialog):
        """빈페이지 삽입 로직"""
        try:
            # 포인트 단위 변환 (1mm = 2.83465 pt)
            # A4: 595.276 x 841.89
            # A3: 841.89 x 1190.55
            
            w, h = 0, 0
            if page_size == "A4":
                w, h = 595, 842
            elif page_size == "A3":
                w, h = 842, 1191
                
            if orientation == "landscape":
                w, h = h, w
                
            # 삽입 위치 결정: 선택된 페이지 앞, 없으면 맨 앞
            indices = self.thumbnail_panel.selected_indices
            if indices:
                insert_pos = min(indices)
            else:
                insert_pos = 0
                
            # PDF Engine에 위임 (engine이 pno 지원해야 함)
            # self.pdf.create_blank_page 호출 시 width, height, pno 전달 필요
            # 현재 pdf_manager.create_blank_page는 arguments를 안 받는 단순 구현일 수 있음.
            # 직접 doc 접근 또는 engine 수정 필요.
            # v3.4 spirit: Logic in engine. But for quick fix I access doc if engine method is weak.
            # Let's check pdf_manager. Or just use doc here.
            
            self.pdf.push_undo_state()
            self.pdf.doc.new_page(pno=insert_pos, width=w, height=h)
            
            dialog.destroy()
            
            self.thumbnail_panel.refresh()
            self.status_bar.config(text=f"빈 페이지({page_size} {orientation}) 추가됨.")
            
            # 선택 업데이트 (삽입된 페이지 선택)
            self.thumbnail_panel.selected_indices = {insert_pos}
            self.thumbnail_panel.refresh_selection_visuals()
            self.preview_panel.show_page(insert_pos)
            
        except Exception as e:
            messagebox.showerror("오류", f"빈페이지 삽입 실패: {e}")
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
        # Calculate scale to fit window
        self.preview_panel.fit_to_window()
    def on_rotate(self, angle):
        if not self.pdf.doc: return
        indices = self.thumbnail_panel.selected_indices
        if not indices:
            messagebox.showinfo("알림", "회전할 페이지를 선택하세요.")
            return
            
        self.pdf.push_undo_state()
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
            self.pdf.push_undo_state()
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
            
        users.sort(key=lambda x: x['name'])
        
        win = tk.Toplevel(self)
        win.title("사용자 목록")
        win.geometry("400x500")
        
        text = tk.Text(win)
        text.pack(fill=BOTH, expand=YES)
        
        for u in users:
            text.insert(END, f"이름: {u['name']}\n직급: {u['role']}\nMAC: {u['mac']}\n{'-'*30}\n")
            
    def add_user_dialog(self):
        """Custom dialog for adding a user."""
        dialog = tk.Toplevel(self)
        dialog.title("사용자 추가")
        dialog.geometry("400x350")
        
        # Center
        x = self.winfo_x() + 50
        y = self.winfo_y() + 50
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="새 사용자 정보 입력", font=("맑은 고딕", 12, "bold")).pack(pady=20)
        
        form_frame = ttk.Frame(dialog, padding=10)
        form_frame.pack(fill=X)
        
        # Name
        ttk.Label(form_frame, text="이름:").grid(row=0, column=0, sticky=W, pady=5)
        name_entry = ttk.Entry(form_frame)
        name_entry.grid(row=0, column=1, sticky=EW, pady=5, padx=5)
        
        # Role
        ttk.Label(form_frame, text="직급:").grid(row=1, column=0, sticky=W, pady=5)
        role_combo = ttk.Combobox(form_frame, values=["user", "admin"], state="readonly")
        role_combo.set("user")
        role_combo.grid(row=1, column=1, sticky=EW, pady=5, padx=5)
        
        # MAC
        ttk.Label(form_frame, text="MAC:").grid(row=2, column=0, sticky=W, pady=5)
        mac_entry = ttk.Entry(form_frame)
        mac_entry.grid(row=2, column=1, sticky=EW, pady=5, padx=5)
        
        form_frame.columnconfigure(1, weight=1)
        
        def save():
            name = name_entry.get().strip()
            role = role_combo.get()
            mac = mac_entry.get().strip()
            
            if not name or not mac:
                messagebox.showwarning("경고", "이름과 MAC 주소는 필수입니다.")
                return
                
            success, msg = self.auth.add_user(name, role, mac)
            if success:
                messagebox.showinfo("성공", msg)
                dialog.destroy()
            else:
                messagebox.showerror("실패", msg)
                
        ttk.Button(dialog, text="추가", command=save, bootstyle="primary").pack(pady=20)
    def manage_users_dialog(self):
        """Custom dialog for managing users (Edit/Remove)."""
        users = self.auth.get_all_users()
        if not users:
             messagebox.showinfo("알림", "관리할 사용자가 없습니다.")
             return
             
        users.sort(key=lambda x: x['name'])
        
        dialog = tk.Toplevel(self)
        dialog.title("사용자 관리 (수정/삭제)")
        dialog.geometry("600x450")
        
        # Center
        x = self.winfo_x() + 50
        y = self.winfo_y() + 50
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="사용자 목록", font=("맑은 고딕", 12, "bold")).pack(pady=10)
        
        # Treeview
        columns = ("name", "role", "mac")
        tree = ttk.Treeview(dialog, columns=columns, show="headings", height=10)
        tree.heading("name", text="이름")
        tree.heading("role", text="직급")
        tree.heading("mac", text="MAC 주소")
        
        tree.column("name", width=100)
        tree.column("role", width=80)
        tree.column("mac", width=150)
        
        for u in users:
            tree.insert("", END, values=(u['name'], u['role'], u['mac']))
            
        tree.pack(fill=BOTH, expand=YES, padx=10, pady=5)
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10, fill=X)
        
        def refresh_list():
            for item in tree.get_children():
                tree.delete(item)
            sorted_users = sorted(self.auth.get_all_users(), key=lambda x: x['name'])
            for u in sorted_users:
                tree.insert("", END, values=(u['name'], u['role'], u['mac']))
        def edit_user():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("경고", "수정할 사용자를 선택하세요.")
                return
                
            item = tree.item(selected_item)
            values = item['values']
            original_mac = values[2]
            
            # Edit Dialog
            edit_win = tk.Toplevel(dialog)
            edit_win.title("사용자 수정")
            edit_win.geometry("400x350")
            edit_win.geometry(f"+{x+20}+{y+20}")
            
            ttk.Label(edit_win, text="사용자 정보 수정", font=("맑은 고딕", 11, "bold")).pack(pady=15)
            
            form = ttk.Frame(edit_win, padding=10)
            form.pack(fill=X)
            
            ttk.Label(form, text="이름:").grid(row=0, column=0, sticky=W, pady=5)
            name_entry = ttk.Entry(form)
            name_entry.insert(0, values[0])
            name_entry.grid(row=0, column=1, sticky=EW, pady=5)
            
            ttk.Label(form, text="직급:").grid(row=1, column=0, sticky=W, pady=5)
            role_combo = ttk.Combobox(form, values=["user", "admin"], state="readonly")
            role_combo.set(values[1])
            role_combo.grid(row=1, column=1, sticky=EW, pady=5)
            
            ttk.Label(form, text="MAC:").grid(row=2, column=0, sticky=W, pady=5)
            mac_entry = ttk.Entry(form)
            mac_entry.insert(0, values[2])
            mac_entry.grid(row=2, column=1, sticky=EW, pady=5)
            
            form.columnconfigure(1, weight=1)
            
            def save_edit():
                new_mac = mac_entry.get().strip()
                new_name = name_entry.get().strip()
                new_role = role_combo.get()
                
                if not new_name or not new_mac:
                    messagebox.showwarning("경고", "필수 항목을 입력하세요.")
                    return
                    
                success, msg = self.auth.update_user(original_mac, new_name, new_role, new_mac)
                if success:
                    messagebox.showinfo("성공", msg)
                    edit_win.destroy()
                    refresh_list()
                else:
                    messagebox.showerror("실패", msg)
            
            ttk.Button(edit_win, text="저장", command=save_edit, bootstyle="success").pack(pady=20)
        
        def delete_user():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("경고", "삭제할 사용자를 선택하세요.")
                return
                
            item = tree.item(selected_item)
            values = item['values']
            name = values[0]
            mac = values[2]
            
            if messagebox.askyesno("확인", f"정말 '{name}' ({mac}) 사용자를 삭제하시겠습니까?"):
                success, msg = self.auth.remove_user(mac)
                if success:
                    messagebox.showinfo("성공", msg)
                    refresh_list()
                else:
                    messagebox.showerror("실패", msg)
        # Buttons
        center_frame = ttk.Frame(btn_frame)
        center_frame.pack()
        
        ttk.Button(center_frame, text="수정", command=edit_user, bootstyle="info").pack(side=LEFT, padx=5)
        ttk.Button(center_frame, text="삭제", command=delete_user, bootstyle="danger").pack(side=LEFT, padx=5)
        ttk.Button(center_frame, text="닫기", command=dialog.destroy, bootstyle="secondary").pack(side=LEFT, padx=5)
    def edit_users_json(self):
         messagebox.showinfo("알림", "메모리 인증 모드에서는 JSON 직접 편집을 지원하지 않습니다.\n'사용자 추가/제거' 기능을 이용해주세요.")
    def export_gist_data(self):
        """Export plain JSON data for Gist from Memory."""
        json_content = self.auth.get_users_json_string()
        if not json_content:
             messagebox.showerror("오류", "데이터 로드 실패.")
             return
        dialog = tk.Toplevel(self)
        dialog.title("Gist용 데이터 내보내기")
        dialog.geometry("600x400")
        
        # Center
        x = self.winfo_x() + 50
        y = self.winfo_y() + 50
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="아래 내용을 복사하여 Gist의 'users.json' 파일에 붙여넣으세요.", bootstyle="info").pack(pady=10)
        
        text_widget = tk.Text(dialog, font=("Consolas", 10), wrap="char", height=10)
        text_widget.pack(fill=BOTH, expand=YES, padx=10)
        
        text_widget.insert("1.0", json_content)
        text_widget.config(state="disabled") # Read-only
        
        def open_gist():
            import webbrowser
            # Gist URL for the user
            gist_url = "https://gist.github.com/TaeW81/8c6597546e977140599d675c4760c298"
            webbrowser.open(gist_url)
            
        link_lbl = ttk.Label(dialog, text="👉 여기를 눌러 Gist 페이지 바로가기", bootstyle="info", cursor="hand2", font=("맑은 고딕", 10, "underline"))
        link_lbl.pack(pady=5)
        link_lbl.bind("<Button-1>", lambda e: open_gist())
        
        def copy():
            self.clipboard_clear()
            self.clipboard_append(json_content)
            self.update()
            messagebox.showinfo("복사 완료", "클립보드에 복사되었습니다.\nGist에 붙여넣으세요.")
            
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="모두 복사", command=copy, bootstyle="primary").pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="닫기", command=dialog.destroy, bootstyle="secondary").pack(side=LEFT, padx=5)
class ExportDialog(tk.Toplevel):
    def __init__(self, parent, total_pages, selected_count):
        super().__init__(parent)
        self.title("내보내기 옵션")
        self.geometry("350x450")
        self.resizable(False, False)
        
        # Center
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 175
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 225
        self.geometry(f"+{x}+{y}")
        
        self.result = None
        self.total_pages = total_pages
        self.selected_count = selected_count
        
        # Variables
        self.var_format = tk.StringVar(value="pdf")
        self.var_range = tk.StringVar(value="all")
        self.var_custom = tk.StringVar()
        
        self.create_widgets()
        
        # Modal
        self.transient(parent)
        self.grab_set()
        
    def create_widgets(self):
        pad = 10
        
        # 1. Format String
        lf_fmt = ttk.Labelframe(self, text="파일 형식", padding=pad)
        lf_fmt.pack(fill=X, padx=pad, pady=pad)
        
        ttk.Radiobutton(lf_fmt, text="PDF 문서 (*.pdf)", variable=self.var_format, value="pdf").pack(anchor=W, pady=2)
        ttk.Radiobutton(lf_fmt, text="JPEG 이미지 (*.jpg)", variable=self.var_format, value="jpg").pack(anchor=W, pady=2)
        ttk.Radiobutton(lf_fmt, text="PNG 이미지 (*.png)", variable=self.var_format, value="png").pack(anchor=W, pady=2)
        
        # 2. Page Range
        lf_range = ttk.Labelframe(self, text="페이지 범위", padding=pad)
        lf_range.pack(fill=X, padx=pad, pady=pad)
        
        ttk.Radiobutton(lf_range, text=f"전체 페이지 ({self.total_pages}장)", variable=self.var_range, value="all", command=self.toggle_custom).pack(anchor=W, pady=2)
        
        state_sel = "normal" if self.selected_count > 0 else "disabled"
        txt_sel = f"선택된 페이지 ({self.selected_count}장)"
        ttk.Radiobutton(lf_range, text=txt_sel, variable=self.var_range, value="selected", state=state_sel, command=self.toggle_custom).pack(anchor=W, pady=2)
        
        f_custom = ttk.Frame(lf_range)
        f_custom.pack(anchor=W, pady=2, fill=X)
        self.rb_custom = ttk.Radiobutton(f_custom, text="페이지 지정:", variable=self.var_range, value="custom", command=self.toggle_custom)
        self.rb_custom.pack(side=LEFT)
        
        self.ent_custom = ttk.Entry(f_custom, textvariable=self.var_custom, state="disabled")
        self.ent_custom.pack(side=LEFT, padx=5, fill=X, expand=YES)
        ttk.Label(lf_range, text="예: 1, 3-5, 8", font=("맑은 고딕", 8), bootstyle="secondary").pack(anchor=W, padx=25)
        
        # 3. Buttons
        f_btn = ttk.Frame(self, padding=pad)
        f_btn.pack(side=BOTTOM, fill=X)
        
        ttk.Button(f_btn, text="내보내기", command=self.on_ok, bootstyle="primary").pack(side=RIGHT, padx=5)
        ttk.Button(f_btn, text="취소", command=self.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
    def toggle_custom(self):
        if self.var_range.get() == "custom":
            self.ent_custom.config(state="normal")
            self.ent_custom.focus()
        else:
            self.ent_custom.config(state="disabled")
    def on_ok(self):
        # Validate Custom
        if self.var_range.get() == "custom" and not self.var_custom.get().strip():
            messagebox.showwarning("입력 확인", "페이지 번호를 입력해주세요.")
            self.ent_custom.focus()
            return
            
        self.result = {
            'format': self.var_format.get(),
            'range': self.var_range.get(),
            'custom_pages': self.var_custom.get()
        }
        self.destroy()


class MergeOrderingDialog(tk.Toplevel):
    def __init__(self, parent, file_paths):
        super().__init__(parent)
        self.title("파일 병합 순서 지정")
        self.geometry("400x400")
        
        # Center
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 200
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 200
        self.geometry(f"+{x}+{y}")
        
        self.result = None
        self.file_paths = list(file_paths)
        
        self.create_widgets()
        
        # Modal
        self.transient(parent)
        self.grab_set()
        
    def create_widgets(self):
        # Listbox Frame
        f_list = ttk.Frame(self, padding=10)
        f_list.pack(fill=BOTH, expand=YES)
        
        ttk.Label(f_list, text="병합할 파일 순서를 조정하세요:", font=("맑은 고딕", 10, "bold")).pack(anchor=W, pady=5)
        
        self.listbox = tk.Listbox(f_list, selectmode=SINGLE, font=("맑은 고딕", 9))
        self.listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        
        scroll = ttk.Scrollbar(f_list, orient=VERTICAL, command=self.listbox.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.listbox.config(yscrollcommand=scroll.set)
        
        # Populate
        for path in self.file_paths:
            self.listbox.insert(END, os.path.basename(path))
            
        # Control Buttons Frame
        f_ctrl = ttk.Frame(self, padding=10)
        f_ctrl.pack(fill=X)
        
        # Up/Down/Remove
        btn_frame = ttk.Frame(f_ctrl)
        btn_frame.pack(side=LEFT)
        
        ttk.Button(btn_frame, text="▲ 위로", command=self.move_up, bootstyle="secondary-outline").pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="▼ 아래로", command=self.move_down, bootstyle="secondary-outline").pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="삭제", command=self.remove_item, bootstyle="danger-outline").pack(side=LEFT, padx=5)
        
        # OK/Cancel
        ttk.Button(f_ctrl, text="병합 시작", command=self.on_ok, bootstyle="primary").pack(side=RIGHT, padx=5)
        ttk.Button(f_ctrl, text="취소", command=self.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
    def move_up(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx > 0:
            text = self.listbox.get(idx)
            path = self.file_paths.pop(idx)
            
            self.listbox.delete(idx)
            self.listbox.insert(idx-1, text)
            self.file_paths.insert(idx-1, path)
            
            self.listbox.selection_set(idx-1)
            
    def move_down(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx < self.listbox.size() - 1:
            text = self.listbox.get(idx)
            path = self.file_paths.pop(idx)
            
            self.listbox.delete(idx)
            self.listbox.insert(idx+1, text)
            self.file_paths.insert(idx+1, path)
            
            self.listbox.selection_set(idx+1)
            
    def remove_item(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        
        self.listbox.delete(idx)
        self.file_paths.pop(idx)
        
        # Select next if possible
        if self.listbox.size() > idx:
            self.listbox.selection_set(idx)
        elif self.listbox.size() > 0:
            self.listbox.selection_set(END)
            
    def on_ok(self):
        if not self.file_paths:
            messagebox.showwarning("경고", "병합할 파일이 없습니다.")
            return
        self.result = self.file_paths
        self.destroy()
