import sys
import os

# Ensure project root is in path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from ui.main_window import MainWindow, WindowManager

if __name__ == "__main__":
    # Create the root window but hide it
    root = tk.Tk()
    root.withdraw() # Hide root window
    
    # Initialize Manager
    manager = WindowManager()
    
    # Create first window
    app = MainWindow(master=root)
    
    # Start loop
    root.mainloop()
