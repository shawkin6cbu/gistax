import threading, tkinter as tk
from tkinter import ttk, messagebox
from desoto.services import fetch_total, DISTRICT_OPTIONS


class TaxTab(ttk.Frame):
    def __init__(self, parent, shared_data, processing_tab):
        super().__init__(parent, padding=20)
        self.shared_data = shared_data
        self.processing_tab = processing_tab

        # helper for labels
        def lbl(text, r, c, **kw):
            ttk.Label(self, text=text, anchor="e")\
               .grid(row=r, column=c, sticky="e",
                     padx=(0, 8), pady=kw.get("pady", 4))

        # ── appraised value ──────────────────────────────────────
        lbl("Appraised value ($):", 0, 0)
        self.value_var = tk.StringVar()
        value_entry = ttk.Entry(self, textvariable=self.value_var, width=18)
        value_entry.grid(row=0, column=1, sticky="w")

        # hitting Enter inside the value field triggers Calculate
        value_entry.bind("<Return>", lambda e: self.calculate_tax())

        # ── district combo ───────────────────────────────────────
        lbl("Tax district:", 1, 0, pady=10)
        self.district_var = tk.StringVar(value=DISTRICT_OPTIONS[0])
        district_cmb = ttk.Combobox(self, textvariable=self.district_var,
                                    values=DISTRICT_OPTIONS, state="readonly",
                                    width=20)
        district_cmb.grid(row=1, column=1, sticky="w")

        # pressing Enter while combo has focus also triggers Calculate
        district_cmb.bind("<Return>", lambda e: self.calculate_tax())

        # ── calculate button & result ────────────────────────────
        self.btn_calc = ttk.Button(self, text="Calculate",
                                   command=self.calculate_tax)
        self.btn_calc.grid(row=2, column=1, sticky="w", pady=(0, 6))

        self.tax_result = tk.StringVar()
        ttk.Label(self, textvariable=self.tax_result,
                  font=("Segoe UI", 12, "bold"))\
           .grid(row=3, column=0, columnspan=2, pady=10)

        # tidy grid columns
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)

    # ── threaded fetch ───────────────────────────────────────────
    def calculate_tax(self):
        raw = self.value_var.get().replace(",", "").strip()
        if not raw.isdigit():
            messagebox.showerror("Input error", "Enter a numeric appraised value.")
            return

        assessed_val = str(round(int(raw) * 0.75))
        district = self.district_var.get()

        self.btn_calc.config(state="disabled")
        self.tax_result.set(f"Calculating on ${assessed_val} …")
        threading.Thread(target=self._thread,
                         args=(assessed_val, district), daemon=True).start()

    def _thread(self, val, district):
        try:
            total = fetch_total(val, district)
            msg = f"TOTAL: ${total}" if total else "Total not found."
        except Exception as e:
            msg = f"Lookup failed: {e}"
        self.after(0, self._done, msg)

    def _done(self, msg):
        self.tax_result.set(msg)
        self.btn_calc.config(state="normal")
        if "TOTAL: $" in msg:
            tax_amount = msg.replace("TOTAL: $", "")
            self.shared_data.set_data("tax_2025_estimated", tax_amount)
            self.processing_tab.load_from_tabs()
