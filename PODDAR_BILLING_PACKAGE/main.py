"""
PODDAR STORES — Professional Billing Software
Full-featured GUI billing system matching Excel functionality.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os, sys, subprocess
from datetime import datetime
from database import (get_items, save_items, get_parties, save_parties,
                      get_invoices, add_invoice, next_invoice_no,
                      item_by_name, party_by_name, get_counter, save_counter)
from pdf_generator import generate_invoice_pdf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INVOICES_DIR = os.path.join(BASE_DIR, "invoices")
os.makedirs(INVOICES_DIR, exist_ok=True)

# ─── colour palette ──────────────────────────────────────────────────────────
C = {
    "bg":      "#F0F4F8",
    "header":  "#1F4E79",
    "accent":  "#2E75B6",
    "accent2": "#0070C0",
    "white":   "#FFFFFF",
    "row1":    "#FFFFFF",
    "row2":    "#DEEAF1",
    "red":     "#C00000",
    "green":   "#375623",
    "text":    "#1A1A1A",
    "border":  "#BDD7EE",
    "light":   "#E9F1F8",
}

MONTHS = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE",
          "JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]

def current_month():
    return datetime.now().strftime("%B").upper()

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════
class PoddarBilling(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PODDAR STORES — Billing Software")
        self.geometry("1200x750")
        self.minsize(1100, 650)
        self.configure(bg=C["bg"])
        self.current_items = []   # rows in active invoice
        self._build_ui()
        self.new_invoice()

    # ──────────────────────────── UI BUILD ──────────────────────────────────
    def _build_ui(self):
        # top header
        hdr = tk.Frame(self, bg=C["header"], height=55)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="PODDAR STORES", font=("Arial",20,"bold"),
                 bg=C["header"], fg="white").pack(side="left", padx=15, pady=10)
        tk.Label(hdr, text="MADHPUR, MABBI, DARBHANGA",
                 font=("Arial",9), bg=C["header"], fg="#BDD7EE").pack(side="left", padx=5)

        # nav buttons
        for (txt, cmd) in [("🧾 New Invoice", self.new_invoice),
                            ("📂 Load Invoice", self.load_invoice_dialog),
                            ("👥 Parties", self.open_parties),
                            ("📦 Items", self.open_items),
                            ("📊 Reports", self.open_reports)]:
            tk.Button(hdr, text=txt, command=cmd, font=("Arial",9,"bold"),
                      bg=C["accent"], fg="white", relief="flat",
                      padx=10, pady=4, cursor="hand2",
                      activebackground=C["accent2"]).pack(side="right", padx=4, pady=12)

        # main body
        body = tk.Frame(self, bg=C["bg"]); body.pack(fill="both", expand=True, padx=10, pady=6)

        # LEFT: invoice form
        left = tk.Frame(body, bg=C["bg"]); left.pack(side="left", fill="both", expand=True)
        self._build_invoice_panel(left)

        # RIGHT: category summary + actions
        right = tk.Frame(body, bg=C["bg"], width=220); right.pack(side="right", fill="y", padx=(8,0))
        right.pack_propagate(False)
        self._build_right_panel(right)

    def _build_invoice_panel(self, parent):
        # ── invoice header ──
        hdr_f = tk.LabelFrame(parent, text=" Invoice Details ", bg=C["bg"],
                              font=("Arial",9,"bold"), fg=C["header"],
                              relief="solid", bd=1)
        hdr_f.pack(fill="x", pady=(0,6))

        row1 = tk.Frame(hdr_f, bg=C["bg"]); row1.pack(fill="x", padx=8, pady=4)

        # Party
        tk.Label(row1, text="Party Name:", bg=C["bg"], font=("Arial",9,"bold"),
                 fg=C["text"]).grid(row=0, column=0, sticky="w", padx=4)
        self.party_var = tk.StringVar()
        party_names = [p["name"] for p in get_parties()]
        self.party_cb = ttk.Combobox(row1, textvariable=self.party_var,
                                     values=party_names, width=30, font=("Arial",9))
        self.party_cb.grid(row=0, column=1, padx=4)
        self.party_cb.bind("<<ComboboxSelected>>", self._on_party_select)

        # Month
        tk.Label(row1, text="Month:", bg=C["bg"], font=("Arial",9,"bold"),
                 fg=C["text"]).grid(row=0, column=2, sticky="w", padx=8)
        self.month_var = tk.StringVar(value=current_month())
        ttk.Combobox(row1, textvariable=self.month_var, values=MONTHS,
                     width=12, state="readonly").grid(row=0, column=3, padx=4)

        # Date
        tk.Label(row1, text="Date:", bg=C["bg"], font=("Arial",9,"bold"),
                 fg=C["text"]).grid(row=0, column=4, sticky="w", padx=8)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        tk.Entry(row1, textvariable=self.date_var, width=12,
                 font=("Arial",9)).grid(row=0, column=5, padx=4)

        # Invoice No (auto)
        tk.Label(row1, text="Invoice No:", bg=C["bg"], font=("Arial",9,"bold"),
                 fg=C["text"]).grid(row=0, column=6, sticky="w", padx=8)
        self.invno_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.invno_var, width=14, font=("Arial",9),
                 state="readonly", bg=C["light"]).grid(row=0, column=7, padx=4)

        # Party info display
        row2 = tk.Frame(hdr_f, bg=C["bg"]); row2.pack(fill="x", padx=8, pady=(0,4))
        self.party_info = tk.Label(row2, text="", bg=C["light"], fg=C["accent"],
                                   font=("Arial",8), anchor="w", relief="flat",
                                   padx=6, pady=2)
        self.party_info.pack(fill="x")

        # ── add item row ──
        add_f = tk.LabelFrame(parent, text=" Add Item ", bg=C["bg"],
                              font=("Arial",9,"bold"), fg=C["header"],
                              relief="solid", bd=1)
        add_f.pack(fill="x", pady=(0,6))

        af = tk.Frame(add_f, bg=C["bg"]); af.pack(fill="x", padx=8, pady=5)
        item_names = [i["name"] for i in get_items()]

        tk.Label(af, text="Item:", bg=C["bg"], font=("Arial",9)).grid(row=0,column=0,sticky="w")
        self.item_var = tk.StringVar()
        self.item_cb = ttk.Combobox(af, textvariable=self.item_var,
                                    values=item_names, width=24, font=("Arial",9))
        self.item_cb.grid(row=0, column=1, padx=4)
        self.item_cb.bind("<<ComboboxSelected>>", self._on_item_select)
        self.item_cb.bind("<Return>", self._on_item_select)

        tk.Label(af, text="Qty:", bg=C["bg"], font=("Arial",9)).grid(row=0,column=2,padx=(8,0))
        self.qty_var = tk.StringVar()
        self.qty_entry = tk.Entry(af, textvariable=self.qty_var, width=8, font=("Arial",9))
        self.qty_entry.grid(row=0, column=3, padx=4)
        self.qty_entry.bind("<Return>", lambda e: self._add_item())

        tk.Label(af, text="Rate:", bg=C["bg"], font=("Arial",9)).grid(row=0,column=4,padx=(8,0))
        self.rate_var = tk.StringVar()
        tk.Entry(af, textvariable=self.rate_var, width=8, font=("Arial",9)).grid(row=0,column=5,padx=4)

        tk.Label(af, text="Unit:", bg=C["bg"], font=("Arial",9)).grid(row=0,column=6,padx=(8,0))
        self.unit_var = tk.StringVar()
        tk.Entry(af, textvariable=self.unit_var, width=6, font=("Arial",9)).grid(row=0,column=7,padx=4)

        tk.Label(af, text="GST%:", bg=C["bg"], font=("Arial",9)).grid(row=0,column=8,padx=(8,0))
        self.gst_var = tk.StringVar()
        tk.Entry(af, textvariable=self.gst_var, width=5, font=("Arial",9)).grid(row=0,column=9,padx=4)

        tk.Button(af, text="➕ Add", command=self._add_item,
                  font=("Arial",9,"bold"), bg=C["accent"], fg="white",
                  relief="flat", padx=10, cursor="hand2").grid(row=0,column=10,padx=8)

        # ── items table ──
        tbl_f = tk.Frame(parent, bg=C["bg"]); tbl_f.pack(fill="both", expand=True)

        cols = ("sr","item","hsn","qty","unit","rate","gst","gst_amt","amount")
        col_labels = ("Sr","Item","HSN","Qty","Unit","Rate","GST%","GST Amt","Amount")
        col_w      = (35, 200, 60, 60, 55, 70, 50, 80, 90)

        self.tree = ttk.Treeview(tbl_f, columns=cols, show="headings", height=14)
        for c_id, lbl, w in zip(cols, col_labels, col_w):
            self.tree.heading(c_id, text=lbl)
            anchor = "w" if c_id == "item" else "center"
            self.tree.column(c_id, width=w, anchor=anchor, minwidth=30)

        style = ttk.Style()
        style.configure("Treeview", rowheight=22, font=("Arial",9))
        style.configure("Treeview.Heading", font=("Arial",9,"bold"),
                        background=C["accent"], foreground="white")
        style.map("Treeview.Heading", background=[("active", C["header"])])

        vsb = ttk.Scrollbar(tbl_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.tag_configure("odd",  background=C["row1"])
        self.tree.tag_configure("even", background=C["row2"])
        self.tree.bind("<Delete>", self._delete_selected)
        self.tree.bind("<Double-1>", self._edit_selected)

        # delete row button
        del_f = tk.Frame(parent, bg=C["bg"]); del_f.pack(fill="x", pady=3)
        tk.Button(del_f, text="🗑 Delete Selected Row", command=self._delete_selected,
                  font=("Arial",8), bg=C["red"], fg="white", relief="flat",
                  padx=8, cursor="hand2").pack(side="left")
        tk.Button(del_f, text="✏ Edit Selected Row", command=self._edit_selected,
                  font=("Arial",8), bg="#7F6000", fg="white", relief="flat",
                  padx=8, cursor="hand2").pack(side="left", padx=6)

        # totals bar
        tot_f = tk.Frame(parent, bg=C["header"], pady=5)
        tot_f.pack(fill="x", pady=(4,0))
        self.lbl_subtotal = tk.Label(tot_f, text="Subtotal: ₹0.00",
                                     font=("Arial",10,"bold"), bg=C["header"], fg="white")
        self.lbl_subtotal.pack(side="left", padx=15)
        self.lbl_gst = tk.Label(tot_f, text="GST: ₹0.00",
                                font=("Arial",10,"bold"), bg=C["header"], fg="#BDD7EE")
        self.lbl_gst.pack(side="left", padx=15)
        self.lbl_grand = tk.Label(tot_f, text="GRAND TOTAL: ₹0",
                                  font=("Arial",13,"bold"), bg=C["header"], fg="#FFD700")
        self.lbl_grand.pack(side="right", padx=15)

    def _build_right_panel(self, parent):
        tk.Label(parent, text="Category Summary", font=("Arial",10,"bold"),
                 bg=C["accent"], fg="white", pady=5).pack(fill="x")

        self.cat_frame = tk.Frame(parent, bg=C["white"], relief="solid", bd=1)
        self.cat_frame.pack(fill="x", pady=4)

        # action buttons
        btn_cfg = [
            ("💾 Save Invoice",  self.save_invoice,  C["green"]),
            ("🖨 Print / PDF",   self.print_invoice, C["accent"]),
            ("🆕 New Invoice",   self.new_invoice,   C["header"]),
            ("📂 Load Invoice",  self.load_invoice_dialog, "#5A5A5A"),
        ]
        for txt, cmd, bg in btn_cfg:
            tk.Button(parent, text=txt, command=cmd, font=("Arial",10,"bold"),
                      bg=bg, fg="white", relief="flat", pady=8,
                      cursor="hand2", activebackground=C["accent2"]
                      ).pack(fill="x", pady=3, padx=4)

        # GST breakdown box
        tk.Label(parent, text="GST Breakdown", font=("Arial",9,"bold"),
                 bg=C["light"], fg=C["header"], pady=3).pack(fill="x", pady=(12,0))
        self.gst_frame = tk.Frame(parent, bg=C["white"], relief="solid", bd=1)
        self.gst_frame.pack(fill="x")

    # ──────────────────────────── LOGIC ────────────────────────────────────
    def new_invoice(self):
        self.current_items.clear()
        for row in self.tree.get_children(): self.tree.delete(row)
        self.party_var.set("")
        self.party_info.config(text="")
        self.month_var.set(current_month())
        self.date_var.set(datetime.now().strftime("%d-%m-%Y"))
        self.invno_var.set("(auto on save)")
        self._update_totals()
        self.item_cb.set("")
        self.qty_var.set("")
        self.rate_var.set("")
        self.unit_var.set("")
        self.gst_var.set("")

    def _on_party_select(self, event=None):
        p = party_by_name(self.party_var.get())
        if p:
            self.party_info.config(
                text=f"Address: {p['address']}  |  Block: {p.get('block','')}  "
                     f"|  District: {p['district']}  |  State: {p.get('state','')}  "
                     f"|  State Code: {p.get('state_code','')}")

    def _on_item_select(self, event=None):
        it = item_by_name(self.item_var.get())
        if it:
            self.rate_var.set(str(it["rate"]))
            self.unit_var.set(it["unit"])
            self.gst_var.set(str(int(it["gst"]*100)))
            self.qty_entry.focus()

    def _add_item(self):
        name = self.item_var.get().strip()
        if not name: messagebox.showwarning("Warning","Please select an item"); return
        try:
            qty  = float(self.qty_var.get())
            rate = float(self.rate_var.get())
            gst  = float(self.gst_var.get()) / 100
        except:
            messagebox.showerror("Error","Invalid Qty/Rate/GST"); return
        unit = self.unit_var.get().strip() or "KGS"
        it_data = item_by_name(name) or {}
        hsn  = it_data.get("hsn","")

        amount     = qty * rate
        gst_amount = round(amount * gst, 4)

        row = {"name":name,"hsn":hsn,"qty":qty,"unit":unit,
               "rate":rate,"gst":gst,"gst_amount":gst_amount,"amount":amount}
        self.current_items.append(row)
        self._refresh_table()
        self.item_cb.set(""); self.qty_var.set("")
        self.rate_var.set(""); self.unit_var.set(""); self.gst_var.set("")
        self.item_cb.focus()

    def _refresh_table(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        for idx, it in enumerate(self.current_items, 1):
            tag = "odd" if idx % 2 else "even"
            gst_pct = f"{int(it['gst']*100)}%"
            self.tree.insert("", "end", iid=str(idx-1), tag=tag, values=(
                idx, it["name"], it.get("hsn",""),
                self._fmt(it["qty"]), it.get("unit",""),
                self._fmt(it["rate"]), gst_pct,
                self._fmt(it["gst_amount"]), self._fmt(it["amount"])
            ))
        self._update_totals()

    def _fmt(self, v):
        if v == int(v): return str(int(v))
        return f"{v:.4f}".rstrip('0').rstrip('.')

    def _delete_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        self.current_items.pop(idx)
        self._refresh_table()

    def _edit_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        it = self.current_items[idx]
        self.item_var.set(it["name"])
        self.qty_var.set(str(it["qty"]))
        self.rate_var.set(str(it["rate"]))
        self.unit_var.set(it.get("unit",""))
        self.gst_var.set(str(int(it["gst"]*100)))
        self.current_items.pop(idx)
        self._refresh_table()

    def _update_totals(self):
        subtotal  = sum(i["amount"] for i in self.current_items)
        total_gst = sum(i["gst_amount"] for i in self.current_items)
        grand     = round(subtotal + total_gst)
        self.lbl_subtotal.config(text=f"Subtotal: ₹{subtotal:,.2f}")
        self.lbl_gst.config(text=f"GST: ₹{total_gst:,.2f}")
        self.lbl_grand.config(text=f"GRAND TOTAL: ₹{grand:,}")
        self._update_category_summary()
        self._update_gst_breakdown(total_gst)

    def _update_category_summary(self):
        for w in self.cat_frame.winfo_children(): w.destroy()
        cats = {}
        for it in self.current_items:
            item_data = item_by_name(it["name"])
            cat = item_data.get("category","OTHER") if item_data else "OTHER"
            cats[cat] = cats.get(cat, 0) + it["amount"]
        for cat, amt in sorted(cats.items()):
            row = tk.Frame(self.cat_frame, bg=C["white"]); row.pack(fill="x", padx=4, pady=1)
            tk.Label(row, text=cat, font=("Arial",8), bg=C["white"],
                     fg=C["text"], anchor="w").pack(side="left")
            tk.Label(row, text=f"₹{amt:,.2f}", font=("Arial",8,"bold"),
                     bg=C["white"], fg=C["accent"], anchor="e").pack(side="right")

    def _update_gst_breakdown(self, total_gst):
        for w in self.gst_frame.winfo_children(): w.destroy()
        slabs = {}
        for it in self.current_items:
            pct = f"{int(it['gst']*100)}%"
            slabs[pct] = slabs.get(pct, 0) + it["gst_amount"]
        for pct, amt in sorted(slabs.items()):
            row = tk.Frame(self.gst_frame, bg=C["white"]); row.pack(fill="x", padx=4, pady=1)
            tk.Label(row, text=f"GST {pct}", font=("Arial",8), bg=C["white"]).pack(side="left")
            tk.Label(row, text=f"₹{amt:,.2f}", font=("Arial",8,"bold"),
                     bg=C["white"], fg=C["accent"]).pack(side="right")
        row = tk.Frame(self.gst_frame, bg=C["light"]); row.pack(fill="x", padx=4, pady=2)
        tk.Label(row, text="CGST", font=("Arial",8), bg=C["light"]).pack(side="left")
        tk.Label(row, text=f"₹{total_gst/2:,.2f}", font=("Arial",8,"bold"),
                 bg=C["light"], fg=C["header"]).pack(side="right")
        row2 = tk.Frame(self.gst_frame, bg=C["light"]); row2.pack(fill="x", padx=4)
        tk.Label(row2, text="SGST", font=("Arial",8), bg=C["light"]).pack(side="left")
        tk.Label(row2, text=f"₹{total_gst/2:,.2f}", font=("Arial",8,"bold"),
                 bg=C["light"], fg=C["header"]).pack(side="right")

    def _build_invoice_dict(self):
        party_name = self.party_var.get().strip()
        if not party_name:
            messagebox.showerror("Error","Please select a party"); return None
        if not self.current_items:
            messagebox.showerror("Error","Please add at least one item"); return None
        party = party_by_name(party_name) or {"name":party_name,"address":"","district":"","block":"","state":""}
        inv_no = self.invno_var.get()
        if inv_no == "(auto on save)":
            inv_no = next_invoice_no(self.month_var.get())
            self.invno_var.set(inv_no)
        return {
            "invoice_no": inv_no,
            "date":       self.date_var.get(),
            "month":      self.month_var.get(),
            "party":      party,
            "items":      [dict(i) for i in self.current_items]
        }

    def save_invoice(self):
        inv = self._build_invoice_dict()
        if not inv: return
        add_invoice(inv)
        messagebox.showinfo("Saved", f"Invoice {inv['invoice_no']} saved successfully!")

    def print_invoice(self):
        inv = self._build_invoice_dict()
        if not inv: return
        fname = f"Invoice_{inv['invoice_no'].replace('/','_')}.pdf"
        out_path = os.path.join(INVOICES_DIR, fname)
        try:
            generate_invoice_pdf(inv, out_path)
            if sys.platform == "win32":
                os.startfile(out_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", out_path])
            else:
                subprocess.run(["xdg-open", out_path])
            messagebox.showinfo("PDF Created", f"Invoice saved:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_invoice_dialog(self):
        invs = get_invoices()
        if not invs:
            messagebox.showinfo("No Invoices","No saved invoices found."); return
        win = tk.Toplevel(self); win.title("Load Invoice")
        win.geometry("700x400"); win.configure(bg=C["bg"])
        tk.Label(win, text="Select Invoice to Load", font=("Arial",11,"bold"),
                 bg=C["accent"], fg="white", pady=6).pack(fill="x")

        cols = ("inv_no","party","date","grand")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
        tree.heading("inv_no", text="Invoice No")
        tree.heading("party",  text="Party")
        tree.heading("date",   text="Date")
        tree.heading("grand",  text="Grand Total")
        for c in cols: tree.column(c, width=150, anchor="center")
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        for inv in reversed(invs):
            items = inv.get("items",[])
            gt = round(sum(i["amount"]+i.get("gst_amount",0) for i in items))
            tree.insert("","end", values=(inv["invoice_no"],
                        inv["party"]["name"], inv["date"], f"₹{gt:,}"))

        def load_sel():
            sel = tree.selection()
            if not sel: return
            idx_vals = tree.item(sel[0])["values"]
            inv_no   = idx_vals[0]
            for inv in invs:
                if inv["invoice_no"] == inv_no:
                    self._load_invoice(inv)
                    win.destroy()
                    return

        tk.Button(win, text="Load Selected", command=load_sel,
                  font=("Arial",10,"bold"), bg=C["accent"], fg="white",
                  relief="flat", padx=15, pady=6).pack(pady=6)

    def _load_invoice(self, inv):
        self.new_invoice()
        self.party_var.set(inv["party"]["name"])
        self._on_party_select()
        self.month_var.set(inv.get("month", current_month()))
        self.date_var.set(inv["date"])
        self.invno_var.set(inv["invoice_no"])
        for it in inv["items"]:
            if "gst_amount" not in it:
                it["gst_amount"] = round(it["amount"] * it["gst"], 4)
            self.current_items.append(it)
        self._refresh_table()

    # ──────────────────────────── MASTERS ──────────────────────────────────
    def open_parties(self):
        MasterWindow(self, "Parties", get_parties, save_parties,
                     ["name","address","block","district","state","state_code"],
                     ["Party Name","Address","Block","District","State","State Code"])

    def open_items(self):
        MasterWindow(self, "Items", get_items, save_items,
                     ["name","hsn","unit","rate","gst","category"],
                     ["Item Name","HSN","Unit","Rate","GST (0-1)","Category"])

    def open_reports(self):
        ReportsWindow(self)


# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER DATA WINDOW
# ═══════════════════════════════════════════════════════════════════════════════
class MasterWindow(tk.Toplevel):
    def __init__(self, parent, title, get_fn, save_fn, fields, labels):
        super().__init__(parent)
        self.title(f"Manage {title}"); self.geometry("900x500")
        self.configure(bg=C["bg"])
        self.get_fn = get_fn; self.save_fn = save_fn
        self.fields = fields; self.labels = labels
        self._build(title)

    def _build(self, title):
        tk.Label(self, text=f"Manage {title}", font=("Arial",12,"bold"),
                 bg=C["accent"], fg="white", pady=6).pack(fill="x")

        main = tk.Frame(self, bg=C["bg"]); main.pack(fill="both", expand=True, padx=8, pady=6)

        # table
        self.tree = ttk.Treeview(main, columns=self.fields, show="headings", height=15)
        for f, l in zip(self.fields, self.labels):
            self.tree.heading(f, text=l)
            self.tree.column(f, width=120, anchor="w")
        vsb = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        # form
        form_f = tk.Frame(main, bg=C["bg"], width=260); form_f.pack(side="right", fill="y", padx=8)
        form_f.pack_propagate(False)
        tk.Label(form_f, text="Add / Edit", font=("Arial",10,"bold"),
                 bg=C["light"], fg=C["header"], pady=4).pack(fill="x")
        self.entries = {}
        for f, l in zip(self.fields, self.labels):
            tk.Label(form_f, text=l+":", font=("Arial",9), bg=C["bg"],
                     fg=C["text"]).pack(anchor="w", pady=(4,0))
            e = tk.Entry(form_f, font=("Arial",9), width=28)
            e.pack(fill="x")
            self.entries[f] = e

        btn_f = tk.Frame(form_f, bg=C["bg"]); btn_f.pack(fill="x", pady=8)
        tk.Button(btn_f, text="➕ Add", command=self._add, font=("Arial",9,"bold"),
                  bg=C["green"], fg="white", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_f, text="✏ Update", command=self._update, font=("Arial",9,"bold"),
                  bg="#7F6000", fg="white", relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(btn_f, text="🗑 Delete", command=self._delete, font=("Arial",9,"bold"),
                  bg=C["red"], fg="white", relief="flat", padx=8).pack(side="left", padx=2)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self._refresh()

    def _refresh(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        for rec in self.get_fn():
            self.tree.insert("", "end", values=[rec.get(f,"") for f in self.fields])

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]
        for f, v in zip(self.fields, vals):
            self.entries[f].delete(0,"end")
            self.entries[f].insert(0, str(v))

    def _get_form(self):
        rec = {}
        for f in self.fields:
            rec[f] = self.entries[f].get().strip()
        return rec

    def _add(self):
        rec = self._get_form()
        if not rec.get("name"):
            messagebox.showerror("Error","Name is required"); return
        # convert numeric fields
        for nf in ["rate","gst"]:
            if nf in rec:
                try: rec[nf] = float(rec[nf])
                except: rec[nf] = 0
        data = self.get_fn(); data.append(rec); self.save_fn(data); self._refresh()

    def _update(self):
        sel = self.tree.selection()
        if not sel: return
        idx  = self.tree.index(sel[0])
        rec  = self._get_form()
        for nf in ["rate","gst"]:
            if nf in rec:
                try: rec[nf] = float(rec[nf])
                except: rec[nf] = 0
        data = self.get_fn(); data[idx] = rec; self.save_fn(data); self._refresh()

    def _delete(self):
        sel = self.tree.selection()
        if not sel: return
        if not messagebox.askyesno("Confirm","Delete this record?"): return
        idx  = self.tree.index(sel[0])
        data = self.get_fn(); data.pop(idx); self.save_fn(data); self._refresh()


# ═══════════════════════════════════════════════════════════════════════════════
#  REPORTS WINDOW
# ═══════════════════════════════════════════════════════════════════════════════
class ReportsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Reports & Analytics"); self.geometry("900x550")
        self.configure(bg=C["bg"])
        self._build()

    def _build(self):
        tk.Label(self, text="Invoice Reports", font=("Arial",12,"bold"),
                 bg=C["accent"], fg="white", pady=6).pack(fill="x")

        # filter row
        flt = tk.Frame(self, bg=C["bg"]); flt.pack(fill="x", padx=8, pady=6)
        tk.Label(flt, text="Filter by Month:", bg=C["bg"], font=("Arial",9)).pack(side="left")
        self.filter_month = tk.StringVar(value="ALL")
        months_opt = ["ALL"] + MONTHS
        ttk.Combobox(flt, textvariable=self.filter_month, values=months_opt,
                     width=12, state="readonly").pack(side="left", padx=6)
        tk.Button(flt, text="🔍 Apply Filter", command=self._load,
                  font=("Arial",9), bg=C["accent"], fg="white", relief="flat",
                  padx=8).pack(side="left")

        # table
        cols = ("inv_no","party","month","date","items","subtotal","gst","grand")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        hdrs = ("Invoice No","Party","Month","Date","Items","Subtotal","GST","Grand Total")
        widths = (110, 180, 90, 90, 50, 90, 80, 100)
        for c,h,w in zip(cols, hdrs, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="center")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8,0), pady=4)
        vsb.pack(side="right", fill="y", pady=4)

        self._load()

    def _load(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        invs = get_invoices()
        fm = self.filter_month.get()
        total_grand = 0
        for inv in invs:
            if fm != "ALL" and inv.get("month","").upper() != fm: continue
            items = inv.get("items",[])
            sub   = sum(i["amount"] for i in items)
            gst   = sum(i.get("gst_amount",0) for i in items)
            grand = round(sub+gst)
            total_grand += grand
            self.tree.insert("","end", values=(
                inv["invoice_no"], inv["party"]["name"],
                inv.get("month",""), inv["date"],
                len(items), f"₹{sub:,.2f}", f"₹{gst:,.2f}", f"₹{grand:,}"
            ))
        # summary row
        self.tree.insert("","end", values=("","","TOTAL","","","","",f"₹{total_grand:,}"),
                         tags=("total",))
        self.tree.tag_configure("total", background=C["accent"], foreground="white",
                                font=("Arial",9,"bold"))


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = PoddarBilling()
    app.mainloop()
