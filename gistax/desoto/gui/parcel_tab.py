import threading, tkinter as tk
from tkinter import ttk
from desoto.services import query_parcels


class ParcelTab(ttk.Frame):
    """Parcel-lookup tab with Refresh button, aligned labels, and wider fields."""

    def __init__(self, parent, shared_data, processing_tab):
        super().__init__(parent, padding=10)
        self.shared_data = shared_data
        self.processing_tab = processing_tab
        win_bg = self.winfo_toplevel().cget("bg")   # read-only Entry background

        # ── search row ───────────────────────────────────────────
        row = ttk.Frame(self)
        row.pack(fill="x")

        ttk.Label(row, text="Address:").pack(side="left", padx=(0, 6))

        self.addr_var = tk.StringVar()
        self.addr_entry = ttk.Entry(row, textvariable=self.addr_var, width=50)
        self.addr_entry.pack(side="left", fill="x", expand=True)
        self.addr_entry.focus()

        ttk.Button(row, text="Refresh", command=self.on_refresh)\
           .pack(side="left", padx=(6, 0))

        self.addr_entry.bind("<KeyRelease>", self.on_type)
        self.addr_entry.bind("<Return>",     self.on_enter)

        # ── results list ─────────────────────────────────────────
        cols = ("Address", "PIN")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=8)
        self.tree.heading("Address", text="Address")
        self.tree.heading("PIN",     text="Parcel #")
        self.tree.column("Address", width=420)
        self.tree.column("PIN",     width=140, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=8)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.place(relx=1, rely=0, relheight=0.74, anchor="ne")

        self.tree.bind("<<TreeviewSelect>>", self.on_pick)

        # ── details grid (all rows aligned) ──────────────────────
        grid = ttk.Frame(self)
        grid.pack(anchor="w", pady=(6, 0))

        # helpers -------------------------------------------------
        def make_entry(var):
            return tk.Entry(grid, textvariable=var, state="readonly",
                            width=40, bd=0, highlightthickness=0,
                            readonlybackground=win_bg)

        def add_row(label_text, r):
            ttk.Label(grid, text=label_text, anchor="e")\
               .grid(row=r, column=0, sticky="e", padx=(0, 6), pady=2)
            var = tk.StringVar()
            make_entry(var).grid(row=r, column=1, sticky="w", pady=2)
            return var

        # rows ----------------------------------------------------
        ttk.Label(grid, text="Parcel:", anchor="e")\
           .grid(row=0, column=0, sticky="e", padx=(0, 6), pady=2)
        self.parcel_var = tk.StringVar()
        make_entry(self.parcel_var).grid(row=0, column=1, sticky="w", pady=2)

        self.address_var = add_row("Address:", 1)
        self.owner1_var = add_row("Owner 1:",          2)
        self.owner2_var = add_row("Owner 2:",          3)
        self.city_var   = add_row("City / State ZIP:", 4)
        self.subd_var   = add_row("Subdivision:",      5)
        self.lot_var    = add_row("Lot:",              6)

        self.results: list[dict] = []

    # ── autocomplete helpers ───────────────────────────────────
    def on_type(self, *_):
        txt = self.addr_var.get()
        if len(txt) < 3:
            self.tree.delete(*self.tree.get_children())
            return
        threading.Thread(target=self.populate, args=(txt,), daemon=True).start()

    def populate(self, text):
        self.results = query_parcels(text)
        self.tree.delete(*self.tree.get_children())
        for i, attr in enumerate(self.results):
            self.tree.insert("", "end", iid=str(i),
                             values=(attr["FULL_ADDR"], attr["PIN"]))

    # ── selection display ───────────────────────────────────────
    def on_pick(self, *_):
        sel = self.tree.selection()
        if not sel:
            return
        attr = self.results[int(sel[0])]
        self.addr_var.set(attr["FULL_ADDR"])
        self.parcel_var.set(attr["PIN"])
        self.address_var.set(attr["FULL_ADDR"])
        owner1 = attr.get("OWNER_NAME", "")
        owner2 = attr.get("SECOND_OWNER", "")
        self.owner1_var.set(owner1)
        self.owner2_var.set(owner2)
        self.city_var.set(f'{attr.get("CITY","")}, {attr.get("STATE","")} {attr.get("ZIP_CODE","")}')
        self.subd_var.set(attr.get("SUBD_NAME",      ""))
        self.lot_var.set(attr.get("LOT",             ""))

        full_owner = owner1
        if owner2:
            full_owner = f"{owner1} & {owner2}"

        self.shared_data.update_data({
            "parcel_pin": attr["PIN"],
            "parcel_address": attr["FULL_ADDR"],
            "parcel_owner": full_owner,
            "parcel_city_state_zip": f'{attr.get("CITY","")}, {attr.get("STATE","")} {attr.get("ZIP_CODE","")}',
            "parcel_legal_description": f'Lot {attr.get("LOT", "")}, {attr.get("SUBD_NAME", "")}',
        })

        self.processing_tab.load_from_tabs()

    # ── Enter selects first result ──────────────────────────────
    def on_enter(self, *_):
        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items[0])
            self.on_pick()

    # ── Refresh clears everything ───────────────────────────────
    def on_refresh(self, *_):
        self.addr_var.set("")
        self.tree.delete(*self.tree.get_children())
        self.parcel_var.set("")
        self.owner1_var.set("")
        self.owner2_var.set("")
        self.city_var.set("")
        self.subd_var.set("")
        self.lot_var.set("")
        self.addr_entry.focus()
