import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES       # TkinterDnD import
from desoto.gui import ParcelTab, TaxTab, TitleTab


class App(TkinterDnD.Tk):                           # <─ inherit from TkinterDnD.Tk
    def __init__(self):
        super().__init__()                          # root now loads the Tk-DND extension

        # ── basic window ───────────────────────────────────────────
        self.title("DeSoto County Utility")
        self.geometry("650x460")
        self.resizable(False, False)

        # Hi-DPI on Windows
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        # ── nicer ttk look ─────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook.Tab", padding=(12, 6))
        style.configure("TButton", padding=(8, 3))
        self.option_add("*Font", ("Segoe UI", 10))

        # ── notebook with tabs ─────────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(expand=True, fill="both", padx=8, pady=8)

        nb.add(ParcelTab(nb), text="Parcel Finder")
        nb.add(TaxTab(nb),    text="Tax Calculator")
        nb.add(TitleTab(nb),  text="TitleDocs")      # Title tab with DND enabled


if __name__ == "__main__":
    App().mainloop()
