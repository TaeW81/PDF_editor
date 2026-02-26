import tkinter as tk

class WindowManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(WindowManager, cls).__new__(cls)
            cls._instance.windows = []
        return cls._instance

    def register(self, window):
        self.windows.append(window)

    def unregister(self, window):
        if window in self.windows:
            self.windows.remove(window)
        if not self.windows:
            # Last window closed, exit app? 
            # Ideally root.destroy() if using a hidden root.
            pass

    def get_windows(self):
        return self.windows

class ClipboardManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(ClipboardManager, cls).__new__(cls)
            cls._instance.data = None # Format: {"source_path": str, "pages": [int]}
        return cls._instance
    
    def copy(self, source_window, page_indices):
        self.data = {
            "source": source_window,
            "pages": sorted(list(page_indices))
        }
        print(f"Copied to clipboard: {len(page_indices)} pages")

    def get_data(self):
        return self.data
    
    def clear(self):
        self.data = None

class DragManager:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DragManager, cls).__new__(cls)
            cls._instance.dragging = False
            cls._instance.source_window = None
            cls._instance.source_indices = None
            cls._instance.drag_window = None # Visual feedback window
        return cls._instance

    def start_drag(self, source_window, indices, event):
        self.dragging = True
        self.source_window = source_window
        self.source_indices = indices
        
        # Create visual drag window (simple transparent window following mouse)
        self.drag_window = tk.Toplevel()
        self.drag_window.overrideredirect(True)
        self.drag_window.attributes("-alpha", 0.6)
        lbl = tk.Label(self.drag_window, text=f"{len(indices)} Pages", bg="blue", fg="white")
        lbl.pack()
        self.update_drag(event)

    def update_drag(self, event):
        if self.drag_window:
            x, y = event.x_root + 10, event.y_root + 10
            self.drag_window.geometry(f"+{x}+{y}")
            
        target_window = None
        for win in WindowManager().get_windows():
            try:
                wx = win.winfo_rootx()
                wy = win.winfo_rooty()
                ww = win.winfo_width()
                wh = win.winfo_height()
                
                if wx <= event.x_root <= wx + ww and wy <= event.y_root <= wy + wh:
                    target_window = win
                    break
            except:
                continue
                
        if target_window and hasattr(target_window, 'on_drag_hover'):
            print(f"Hovering over {target_window} at {event.x_root}, {event.y_root}")
            target_window.on_drag_hover(self.source_window, self.source_indices, event.x_root, event.y_root)
        else:
            # Clear all guides if hovering outside
            for win in WindowManager().get_windows():
                if hasattr(win, 'clear_drag_guide'):
                    win.clear_drag_guide()

    def stop_drag(self, event, x, y):
        self.dragging = False
        if self.drag_window:
            self.drag_window.destroy()
            self.drag_window = None
            
        for win in WindowManager().get_windows():
            if hasattr(win, 'clear_drag_guide'):
                win.clear_drag_guide()
        
        target_window = None
        # Check all registered windows
        for win in WindowManager().get_windows():
            try:
                wx = win.winfo_rootx()
                wy = win.winfo_rooty()
                ww = win.winfo_width()
                wh = win.winfo_height()
                
                if wx <= x <= wx + ww and wy <= y <= wy + wh:
                    target_window = win
                    break
            except:
                continue

        if target_window:
            # Check if dropped on thumbnail panel?
            # Ideally pass to window to handle specific drop logic
            target_window.on_external_drop(self.source_window, self.source_indices, x, y)
            
        self.source_window = None
        self.source_indices = None
