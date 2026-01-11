#main.py
import sys
import tkinter as tk
import ctypes
from ui import DefectApp 


if __name__ == "__main__":
    if sys.platform.startswith('win'):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            ctypes.windll.user32.SetProcessDPIAware()

    root = tk.Tk()
    app = DefectApp(root)
    root.mainloop()