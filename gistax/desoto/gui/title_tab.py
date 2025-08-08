import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinterdnd2 as tkdnd
from tkinterdnd2 import DND_FILES
import os
from desoto.services.title_chain import process_title_pdf


class TitleTab(ttk.Frame):
    """Title document processing tab with drag-and-drop PDF input."""

    def __init__(self, parent, shared_data):
        super().__init__(parent, padding=10)
        self.shared_data = shared_data

        # ── PDF Input Section ───────────────────────────────────
        pdf_frame = ttk.LabelFrame(self, text="PDF Input", padding=10)
        pdf_frame.pack(fill="x", pady=(0, 10))

        self.pdf_var = tk.StringVar()
        self.drop_frame = tk.Frame(pdf_frame, bg="#f0f0f0", relief="sunken", bd=2, height=80)
        self.drop_frame.pack(fill="x", pady=(0, 10))
        self.drop_frame.pack_propagate(False)

        tk.Label(
            self.drop_frame,
            text="Drop PDF file here or click Browse",
            bg="#f0f0f0",
            fg="#666",
            font=("Arial", 10),
        ).pack(expand=True)

        path_row = ttk.Frame(pdf_frame)
        path_row.pack(fill="x")
        ttk.Label(path_row, text="PDF File:").pack(side="left", padx=(0, 6))

        self.pdf_entry = ttk.Entry(path_row, textvariable=self.pdf_var, width=60)
        self.pdf_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(path_row, text="Browse", command=self.browse_pdf).pack(side="left", padx=(6, 0))

        # ── Output Section ──────────────────────────────────────
        output_frame = ttk.LabelFrame(self, text="Output", padding=10)
        output_frame.pack(fill="x", pady=(0, 10))

        output_row = ttk.Frame(output_frame)
        output_row.pack(fill="x")
        ttk.Label(output_row, text="Save As:").pack(side="left", padx=(0, 6))

        self.output_var = tk.StringVar()
        self.output_entry = ttk.Entry(output_row, textvariable=self.output_var, width=60)
        self.output_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(output_row, text="Browse", command=self.browse_output).pack(side="left", padx=(6, 0))

        ttk.Button(output_frame, text="Process Title Document", command=self.process_document).pack(
            pady=(10, 0)
        )

        # ── Status Section ──────────────────────────────────────
        status_frame = ttk.LabelFrame(self, text="Status", padding=10)
        status_frame.pack(fill="x", pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor="w")

        self.progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(5, 0))

        # ── Results Section ─────────────────────────────────────
        results_frame = ttk.LabelFrame(self, text="24-Month Chain Results", padding=10)
        results_frame.pack(fill="both", expand=True)

        cols = ("Date", "Grantor", "Grantee", "Instrument", "Book-Page")
        self.results_tree = ttk.Treeview(results_frame, columns=cols, show="headings", height=6)
        for col, w, anchor in (
            ("Date", 80, "center"),
            ("Grantor", 200, "w"),
            ("Grantee", 200, "w"),
            ("Instrument", 150, "w"),
            ("Book-Page", 100, "center"),
        ):
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=w, anchor=anchor)
        self.results_tree.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # ── Drag-and-drop setup ─────────────────────────────────
        self.setup_drag_drop()

        # ── Template path ───────────────────────────────────────
        self.template_path = self.get_template_path()

    # ------------------------------------------------------------------
    # Helper methods
    # ------------------------------------------------------------------
    def get_template_path(self):
        try:
            current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            path = os.path.join(current_dir, "templates", "td_tmplt2.docx")
            return path if os.path.exists(path) else None
        except Exception:
            return None

    def setup_drag_drop(self):
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.on_drop)
        self.drop_frame.bind("<Button-1>", lambda _e: self.browse_pdf())
        self.drop_frame.bind("<Enter>", lambda _e: self.drop_frame.config(cursor="hand2"))
        self.drop_frame.bind("<Leave>", lambda _e: self.drop_frame.config(cursor=""))

    def on_drop(self, event):
        """Handle a file being dropped on the frame."""
        # Properly parse the file list returned by tkdnd (handles braces & spaces)
        files = self.tk.splitlist(event.data)
        if not files:
            return

        path = files[0]        # first (and usually only) dropped file
        if path.lower().endswith(".pdf"):
            self.pdf_var.set(path)
            self.auto_set_output_path()
        else:
            messagebox.showerror("Error", "Please drop a PDF file.")

    # ------------------------------------------------------------------
    # File dialogs
    # ------------------------------------------------------------------
    def browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Select PDF File", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)
            self.auto_set_output_path()

    def auto_set_output_path(self):
        pdf_path = self.pdf_var.get()
        if pdf_path:
            name = os.path.splitext(os.path.basename(pdf_path))[0]
            self.output_var.set(os.path.join(os.path.dirname(pdf_path), f"{name}_TitleDoc.docx"))

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Title Document As",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.output_var.set(path)

    # ------------------------------------------------------------------
    # Processing
    # ------------------------------------------------------------------
    def process_document(self):
        pdf = self.pdf_var.get().strip()
        out = self.output_var.get().strip()
        if not pdf or not os.path.exists(pdf):
            messagebox.showerror("Error", "Please select a valid PDF file.")
            return
        if not out:
            messagebox.showerror("Error", "Please specify an output location.")
            return
        threading.Thread(target=self._do_process, args=(pdf, out), daemon=True).start()

    def _do_process(self, pdf, out):
        try:
            self._ui(lambda: self.status_var.set("Processing PDF..."))
            self._ui(self.progress.start)
            self._ui(lambda: self.results_tree.delete(*self.results_tree.get_children()))

            success, msg, deeds = process_title_pdf(pdf, out, self.template_path)

            self._ui(self.progress.stop)
            self._ui(lambda: self.status_var.set(msg))
            if success:
                self.shared_data.set_data("title_chain_results", deeds)
                self._ui(lambda: self.display_results(deeds))
                self._ui(lambda: messagebox.showinfo("Success", msg))
            else:
                self._ui(lambda: messagebox.showerror("Error", msg))
        except Exception as e:
            self._ui(self.progress.stop)
            self._ui(lambda: self.status_var.set(f"Error: {e}"))
            self._ui(lambda: messagebox.showerror("Error", f"Processing failed: {e}"))

    def _ui(self, fn):
        self.after(0, fn)

    # ------------------------------------------------------------------
    # Results
    # ------------------------------------------------------------------
    def display_results(self, deeds):
        self.results_tree.delete(*self.results_tree.get_children())
        if not deeds:
            self.results_tree.insert("", "end", values=("", "No vesting deeds found", "", "", ""))
            return
        for d in deeds:
            self.results_tree.insert(
                "",
                "end",
                values=(d.date_string, d.grantor, d.grantee, d.instrument, d.book_page),
            )

    def refresh_all(self):
        self.pdf_var.set("")
        self.output_var.set("")
        self.status_var.set("Ready")
        self.results_tree.delete(*self.results_tree.get_children())
        self.progress.stop()
