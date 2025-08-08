import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from desoto.data import SharedData
from desoto.gui import ParcelTab, TaxTab, TitleTab, ProcessingTab


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("DeSoto County Utility")
        self.geometry("800x600") # Increased size for new tab
        self.resizable(True, True)

        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook.Tab", padding=(12, 6))
        style.configure("TButton", padding=(8, 3))
        self.option_add("*Font", ("Segoe UI", 10))

        self.shared_data = SharedData()

        nb = ttk.Notebook(self)
        nb.pack(expand=True, fill="both", padx=8, pady=8)

        # Create tabs
        processing_tab = ProcessingTab(nb, self.shared_data)
        parcel_tab = ParcelTab(nb, self.shared_data, processing_tab)
        tax_tab = TaxTab(nb, self.shared_data, processing_tab)
        title_tab = TitleTab(nb, self.shared_data)

        # Add tabs to notebook
        nb.add(parcel_tab, text="Parcel Finder")
        nb.add(tax_tab, text="Tax Calculator")
        nb.add(title_tab, text="TitleDocs")
        nb.add(processing_tab, text="Processing")


if __name__ == "__main__":
    App().mainloop()
