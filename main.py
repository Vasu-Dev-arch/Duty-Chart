# main.py
import logging; 
import tkinter as tk
from gui.app import DutyChartApp
from config.logging import setup_logging

if __name__ == "__main__":
    setup_logging()
    root = tk.Tk()
    app = DutyChartApp(root)
    root.mainloop()