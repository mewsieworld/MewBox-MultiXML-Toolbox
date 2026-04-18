"""
SCRIPT 4 – NCash / Ticket Cost Updater  (Parent-Box CSV variant)
─────────────────────────────────────────────────────────────────
Same as Script 3 but the CSV comes directly from the parent-box
sheet, which looks like:

  ID , Level , Dragon     , ID , Level , Sheep      , ID , Level , Fox ...
  432002 , 190 , [JP]Hanyu , 432045 , 140 , Goodie Bag , ...

Any number of repeating groups; any columns between them.
Only columns whose header is exactly "ID" (case-insensitive) are
read as item IDs. Every other column (Level, Rate, names, etc.) is
ignored — including any header that matches an ItemParam XML tag.

ItemParam: pick any one of the 4 files; siblings auto-loaded.
Each file processed independently. Only files with matches exported.
Formula: NCash = round(tickets × 133)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, os

# ═══════════════════════════════════════════════════════════════════════════════
# XML HELPERS  (identical to Script 3)
# ═══════════════════════════════════════════════════════════════════════════════
ROW_RE   = re.compile(r'<ROW>.*?</ROW>', re.DOTALL)
CDATA_RE = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)

def _get_tag(block, tag):
    m = re.search(rf'<{re.escape(tag)}>(.*?)</{re.escape(tag)}>', block, re.DOTALL)
    if not m: return ""
    cd = CDATA_RE.search(m.group(1))
    return cd.group(1).strip() if cd else m.group(1).strip()

def build_item_lib(files):
    """files: [(fname, text), ...]. Returns {id_str: name_str}."""
    lib = {}
    for _, text in files:
        for row in ROW_RE.findall(text):
            rid  = _get_tag(row, "ID")
            name = _get_tag(row, "Name")
            if rid.isdigit() and name:
                lib[rid] = name
    return lib

def bulk_update_ncash(xml_text, updates):
    """
    updates: {id_str: ncash_int}
    Single pass. Returns (modified_text, {id_str: bool_found}).
    """
    found = {k: False for k in updates}
    def replace_row(m):
        block = m.group(0)
        rid   = _get_tag(block, "ID")
        if rid not in updates:
            return block
        found[rid] = True
        return re.sub(r'<Ncash>\d+</Ncash>',
                      f'<Ncash>{updates[rid]}</Ncash>', block)
    return ROW_RE.sub(replace_row, xml_text), found

# ═══════════════════════════════════════════════════════════════════════════════
# CSV PARSER  — parent-box CSV specific
# ═══════════════════════════════════════════════════════════════════════════════

# Every ItemParam XML tag name lowercased, MINUS "id" and "ncash" which are
# the two fields this script actually cares about.  Any CSV column whose
# header matches one of these is silently ignored when scanning for IDs.
_NON_ID_HEADERS = {
    "acplus","ap","applus","bundlenum","cardgengrade","cardgenparam","cardnum",
    "chrftypeflag","chrgender","chrtypeflags","class","cmtbundlenum","cmtfilename",
    "comment","comment_eng","compoundslot","daplus","dpplus","dxplus","dailygencnt",
    "delay","depth","effect","effectflags2","equipfilename","existtype","famcm",
    "filename","groundflags","groupid","hp","hpcon","hprecoveryrate","hvplus",
    "hidehat","invbundlenum","invfilename","itemftype","lkplus","life","maplus",
    "mdplus","mp","mpcon","mprecoveryrate","maxhpplus","maxmpplus","maxwtplus",
    "minlevel","minstatLv","minstattype","money","name","name_eng","newcm",
    "options","optionsex","paletteid","partfilename","pivotid","refineindex",
    "refinetype","reformcount","selrange","setitemid","shopbundlenum","shopfilename",
    "subtype","summary","systemflags","type","use","value","weight",
    # common spreadsheet column words that are never IDs
    "level","rate","lv","luck","lvl","chance","prob","drop","qty",
    "quantity","count","amount","row",
}

def _find_value_col(raw_headers, id_pos):
    """
    Given the position of an ID column, scan to the right for the nearest
    column whose header is exactly "tickets", "ticket", "ncash", or "ncash_val"
    (case-insensitive), stopping before the next "id" column.
    Returns (col_index, value_type) where value_type is "tickets" or "ncash",
    or (None, None) if not found.
    """
    TICKET_NAMES = {"tickets", "ticket"}
    NCASH_NAMES  = {"ncash", "ncash_val", "ncashval"}
    next_id = next((i for i in range(id_pos + 1, len(raw_headers))
                    if raw_headers[i].lower() == "id"), len(raw_headers))
    for i in range(id_pos + 1, next_id):
        h = raw_headers[i].lower()
        if h in TICKET_NAMES: return i, "tickets"
        if h in NCASH_NAMES:  return i, "ncash"
    return None, None

def parse_parentbox_csv(text):
    """
    Reads a parent-box CSV.
    - Collects every cell from columns headed exactly "ID" (case-insensitive).
    - For each ID column, looks rightward (before the next ID column) for a
      "Tickets"/"Ticket" or "NCash" column and reads that value.
      Tickets are stored as ticket_cost (converted ×133 at process time).
      NCash values are stored as ncash_direct (used as-is).
    - All other columns are ignored entirely.
    Returns list of {"id", "ticket_cost", "ncash_direct", "name"}.
    """
    stripped = text.strip()
    if not stripped:
        return []

    all_rows = list(csv.reader(io.StringIO(stripped)))
    if not all_rows:
        return []

    raw_headers = [h.strip() for h in all_rows[0]]
    data_rows   = all_rows[1:]

    id_positions = [i for i, h in enumerate(raw_headers) if h.lower() == "id"]

    # For each ID column, find its paired value column if any
    # val_map: {id_col_pos: (val_col_pos, "tickets"|"ncash")}
    val_map = {}
    for id_pos in id_positions:
        vcol, vtype = _find_value_col(raw_headers, id_pos)
        if vcol is not None:
            val_map[id_pos] = (vcol, vtype)

    items, seen = [], set()

    def add(id_str, ticket_cost, ncash_direct):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str)
            items.append({
                "id":           id_str,
                "ticket_cost":  ticket_cost,    # float or None
                "ncash_direct": ncash_direct,   # int or None
                "name":         "",
            })

    if id_positions:
        for row in data_rows:
            for id_pos in id_positions:
                if id_pos >= len(row):
                    continue
                id_val = row[id_pos].strip()
                if not (id_val and id_val.isdigit()):
                    continue
                ticket_cost  = None
                ncash_direct = None
                if id_pos in val_map:
                    vcol, vtype = val_map[id_pos]
                    if vcol < len(row):
                        raw = row[vcol].strip()
                        try:
                            num = float(raw)
                            if vtype == "tickets":
                                ticket_cost  = num
                            else:
                                ncash_direct = int(round(num))
                        except ValueError:
                            pass
                add(id_val, ticket_cost, ncash_direct)
        return items

    # Fallback: no "ID" header — grab any numeric cell not in a known non-ID col
    for row in data_rows:
        for i, cell in enumerate(row):
            hdr = raw_headers[i].lower() if i < len(raw_headers) else ""
            if hdr not in _NON_ID_HEADERS:
                add(cell, None, None)
    return items

# ═══════════════════════════════════════════════════════════════════════════════
# APP
# ═══════════════════════════════════════════════════════════════════════════════
TARGET_FILES = {"itemparam2.xml", "itemparamcm2.xml", "itemparamex2.xml", "itemparamex.xml"}

class NCashUpdaterParentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 4 – NCash Updater (Parent-Box CSV)")
        self.geometry("980x700")
        self.configure(bg="#1e1e2e")
        self.csv_items = []
        self.xml_files = []
        self.item_lib  = {}
        self.mode_var  = tk.StringVar(value="uniform")
        self._build_load_screen()

    # ─────────────────────────────────────────────────────────────────────────
    # LOAD SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="NCASH UPDATER — PARENT-BOX CSV",
                 font=("Consolas", 16, "bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(24, 2))
        tk.Label(self,
                 text="Reads IDs from every column headed  \"ID\"  in your parent-box CSV.\n"
                      "All other columns (Level, Rate, names, tags…) are ignored.\n"
                      "Formula: NCash = round(tickets × 133)",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 9),
                 justify="center").pack(pady=(0, 10))

        csv_status = tk.StringVar(value="No file loaded")
        xml_status = tk.StringVar(value="No file loaded")

        def section(title):
            f = tk.LabelFrame(self, text=f"  {title}  ", bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=30, pady=5)
            return f

        # ── CSV ──────────────────────────────────────────────────────────────
        csv_frm = section("Step 1 — Parent-Box CSV  (ID, Tickets or NCash — all other values filtered out)")
        tk.Label(csv_frm, textvariable=csv_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)

        def load_csv():
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
            items = parse_parentbox_csv(text)
            if not items:
                messagebox.showerror("Error",
                    "No item IDs found.\n\n"
                    "Make sure your CSV has at least one column headed exactly \"ID\"."); return
            self.csv_items = items
            if self.item_lib:
                for it in self.csv_items:
                    it["name"] = self.item_lib.get(it["id"], "")
            csv_status.set(f"✓  {os.path.basename(path)}  —  {len(items)} unique IDs")

        tk.Button(csv_frm, text="📂 Load", command=load_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── ItemParam ────────────────────────────────────────────────────────
        xml_frm = section("Step 2 — ItemParam XML  (pick any one of the 4 files)")
        tk.Label(xml_frm, textvariable=xml_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)

        def load_xml():
            path = filedialog.askopenfilename(
                title="Select any one of the 4 ItemParam XML files",
                filetypes=[("XML","*.xml"),("All","*.*")])
            if not path: return
            folder = os.path.dirname(path)
            loaded = []
            for fname in os.listdir(folder):
                if fname.lower() in TARGET_FILES:
                    try:
                        with open(os.path.join(folder, fname),
                                  encoding="utf-8-sig", errors="replace") as f:
                            loaded.append((fname, f.read()))
                    except Exception:
                        pass
            if not loaded:
                messagebox.showerror("Error", "None of the 4 ItemParam files found."); return
            self.xml_files = loaded
            self.item_lib  = build_item_lib(loaded)
            for it in self.csv_items:
                it["name"] = self.item_lib.get(it["id"], "")
            fnames = [fn for fn, _ in loaded]
            xml_status.set(
                f"✓  {len(loaded)}/4 files  |  {len(self.item_lib)} items indexed  "
                f"({', '.join(fnames)})")

        tk.Button(xml_frm, text="📂 Load", command=load_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── Mode ─────────────────────────────────────────────────────────────
        mode_frm = section("Step 3 — Mode")
        mf = tk.Frame(mode_frm, bg="#1e1e2e"); mf.pack(anchor="w", padx=10, pady=6)
        tk.Radiobutton(mf, text="Uniform  —  one ticket cost applied to every ID",
                       variable=self.mode_var, value="uniform",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(anchor="w", pady=2)
        tk.Radiobutton(mf, text="Manual   —  set ticket cost per ID individually",
                       variable=self.mode_var, value="manual",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(anchor="w", pady=2)

        def proceed():
            if not self.csv_items:
                messagebox.showwarning("Missing", "Load a CSV first."); return
            if not self.xml_files:
                messagebox.showwarning("Missing", "Load ItemParam XML first."); return
            if self.mode_var.get() == "uniform":
                self._build_uniform_screen()
            else:
                self._build_manual_screen()

        tk.Button(self, text="▶  Continue →", command=proceed,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 12, "bold"),
                  relief="flat", padx=20, pady=8).pack(pady=18)

    # ─────────────────────────────────────────────────────────────────────────
    # UNIFORM SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_uniform_screen(self):
        self._clear()
        tk.Label(self, text="Uniform Value",
                 font=("Consolas", 14, "bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(20, 4))
        tk.Label(self,
                 text=f"Applied to all {len(self.csv_items)} IDs.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack(pady=(0, 8))

        # Detect if CSV pre-loaded a uniform value
        # Prefer ncash_direct if present, else ticket_cost
        pre_type  = None
        pre_value = None
        ticket_costs  = [it["ticket_cost"]  for it in self.csv_items if it.get("ticket_cost")  is not None]
        ncash_directs = [it["ncash_direct"] for it in self.csv_items if it.get("ncash_direct") is not None]
        if ncash_directs and len(set(ncash_directs)) == 1:
            pre_type  = "ncash"
            pre_value = str(ncash_directs[0])
        elif ticket_costs and len(set(ticket_costs)) == 1:
            pre_type  = "tickets"
            pre_value = str(ticket_costs[0])

        val_type_var = tk.StringVar(value=pre_type or "tickets")

        type_frm = tk.Frame(self, bg="#1e1e2e"); type_frm.pack(pady=(0,6))
        tk.Radiobutton(type_frm, text="Tickets  (× 133 = NCash)",
                       variable=val_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10),
                       command=lambda: _refresh_preview()).pack(side="left", padx=10)
        tk.Radiobutton(type_frm, text="NCash  (exact)",
                       variable=val_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10),
                       command=lambda: _refresh_preview()).pack(side="left", padx=10)

        frm = tk.Frame(self, bg="#1e1e2e"); frm.pack()
        val_lbl = tk.Label(frm, text="Tickets:", bg="#1e1e2e", fg="#cdd6f4",
                           font=("Consolas", 12))
        val_lbl.pack(side="left", padx=8)
        tv = tk.StringVar(value=pre_value or "")
        ent = tk.Entry(frm, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                       insertbackground="#cdd6f4", font=("Consolas", 12), relief="flat")
        ent.pack(side="left", padx=8)
        ent.focus()

        preview_var = tk.StringVar(value="NCash: —")
        tk.Label(self, textvariable=preview_var, bg="#1e1e2e", fg="#a6e3a1",
                 font=("Consolas", 12, "bold")).pack(pady=8)

        def _refresh_preview(*_):
            is_tickets = val_type_var.get() == "tickets"
            val_lbl.config(text="Tickets:" if is_tickets else "NCash:")
            try:
                num = float(tv.get())
                if is_tickets:
                    preview_var.set(f"NCash: {round(num * 133)}")
                else:
                    preview_var.set(f"NCash: {int(round(num))}  (exact)")
            except:
                preview_var.set("NCash: —")

        tv.trace_add("write", _refresh_preview)
        _refresh_preview()

        def apply_uniform():
            try:    num = float(tv.get())
            except:
                messagebox.showwarning("Invalid", "Enter a valid number."); return
            is_tickets = val_type_var.get() == "tickets"
            for it in self.csv_items:
                if is_tickets:
                    it["ticket_cost"]  = num
                    it["ncash_direct"] = None
                else:
                    it["ncash_direct"] = int(round(num))
                    it["ticket_cost"]  = None
            self._process_and_show()

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(pady=16)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=8)
        tk.Button(bot, text="✓  Apply to All & Update XML", command=apply_uniform,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", padx=16, pady=8).pack(side="left", padx=8)

    # ─────────────────────────────────────────────────────────────────────────
    # MANUAL SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_manual_screen(self):
        self._clear()
        tk.Label(self, text="Manual Ticket Costs",
                 font=("Consolas", 14, "bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(12, 2))
        tk.Label(self, text="Leave blank to skip an ID.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 9)).pack(pady=(0, 4))

        outer  = tk.Frame(self, bg="#1e1e2e"); outer.pack(fill="both", expand=True, padx=20, pady=4)
        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        cont = tk.Frame(canvas, bg="#1e1e2e")
        wid  = canvas.create_window((0, 0), window=cont, anchor="nw")
        cont.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        # ── Column type toggle ────────────────────────────────────────────
        type_row = tk.Frame(cont, bg="#1e1e2e"); type_row.pack(anchor="w", padx=6, pady=(4,2))
        col_type_var = tk.StringVar(value="tickets")
        # Detect pre-loaded type from CSV items
        has_ncash   = any(it.get("ncash_direct")  is not None for it in self.csv_items)
        has_tickets = any(it.get("ticket_cost")    is not None for it in self.csv_items)
        if has_ncash and not has_tickets:
            col_type_var.set("ncash")
        tk.Label(type_row, text="Input type:", bg="#1e1e2e", fg="#a6adc8",
                 font=("Consolas",9)).pack(side="left", padx=(0,6))
        tk.Radiobutton(type_row, text="Tickets (×133)",
                       variable=col_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",9),
                       command=lambda: _refresh_labels()).pack(side="left", padx=4)
        tk.Radiobutton(type_row, text="NCash (exact)",
                       variable=col_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",9),
                       command=lambda: _refresh_labels()).pack(side="left", padx=4)

        hdr = tk.Frame(cont, bg="#181825"); hdr.pack(fill="x", pady=2)
        for txt, w in [("Item ID",12),("Item Name",30),("Value",14),("NCash (result)",16)]:
            tk.Label(hdr, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas",9,"bold"), anchor="w").pack(side="left", padx=6, pady=4)

        ticket_vars  = []
        ncash_labels = []
        val_header_labels = []  # so _refresh_labels can update them

        for i, item in enumerate(self.csv_items):
            bg  = "#1e1e2e" if i % 2 == 0 else "#181825"
            row = tk.Frame(cont, bg=bg); row.pack(fill="x")
            tk.Label(row, text=item["id"], width=12, bg=bg, fg="#585b70",
                     font=("Consolas",9), anchor="w").pack(side="left", padx=6, pady=2)
            name = item.get("name") or self.item_lib.get(item["id"], "—")
            tk.Label(row, text=name[:32], width=30, bg=bg, fg="#a6adc8",
                     font=("Consolas",9), anchor="w").pack(side="left", padx=6, pady=2)
            tv = tk.StringVar()
            # Pre-populate: ncash_direct takes priority, else ticket_cost
            if item.get("ncash_direct") is not None and col_type_var.get() == "ncash":
                tv.set(str(item["ncash_direct"]))
            elif item.get("ticket_cost") is not None:
                tv.set(str(item["ticket_cost"]))
            ticket_vars.append(tv)
            tk.Entry(row, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas",9),
                     relief="flat").pack(side="left", padx=6, pady=2)
            ncash_lbl = tk.Label(row, text="—", width=16, bg=bg, fg="#a6e3a1",
                                 font=("Consolas",9), anchor="w")
            ncash_lbl.pack(side="left", padx=6)
            ncash_labels.append(ncash_lbl)
            def make_trace(var, lbl):
                def cb(*_):
                    try:
                        num = float(var.get())
                        if col_type_var.get() == "tickets":
                            lbl.config(text=str(round(num * 133)))
                        else:
                            lbl.config(text=str(int(round(num))) + " (exact)")
                    except: lbl.config(text="—")
                var.trace_add("write", cb); cb()
            make_trace(tv, ncash_lbl)

        def _refresh_labels():
            """Recalculate all ncash preview labels when type radio changes."""
            for i, (var, lbl) in enumerate(zip(ticket_vars, ncash_labels)):
                try:
                    num = float(var.get())
                    if col_type_var.get() == "tickets":
                        lbl.config(text=str(round(num * 133)))
                    else:
                        lbl.config(text=str(int(round(num))) + " (exact)")
                except:
                    lbl.config(text="—")

        def confirm():
            blanks = []
            is_tickets = col_type_var.get() == "tickets"
            for i, item in enumerate(self.csv_items):
                raw = ticket_vars[i].get().strip()
                try:
                    num = float(raw)
                    if is_tickets:
                        item["ticket_cost"]  = num
                        item["ncash_direct"] = None
                    else:
                        item["ncash_direct"] = int(round(num))
                        item["ticket_cost"]  = None
                except:
                    item["ticket_cost"]  = None
                    item["ncash_direct"] = None
                    blanks.append(item["id"])
            if blanks:
                ans = messagebox.askyesno(
                    "Missed a spot",
                    f"{len(blanks)} ID(s) have no cost and will be SKIPPED:\n\n"
                    + ", ".join(blanks[:20]) + ("…" if len(blanks) > 20 else "")
                    + "\n\nContinue anyway?")
                if not ans: return
            self._process_and_show()

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x", pady=6)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=14)
        tk.Button(bot, text="✓  Apply & Update XML", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=16, pady=8).pack(side="right", padx=14)

    # ─────────────────────────────────────────────────────────────────────────
    # PROCESS
    # ─────────────────────────────────────────────────────────────────────────
    def _process_and_show(self):
        def resolve_ncash(it):
            """ticket_cost → ×133; ncash_direct → as-is; neither → None."""
            if it.get("ticket_cost") is not None:
                return round(it["ticket_cost"] * 133)
            if it.get("ncash_direct") is not None:
                return int(it["ncash_direct"])
            return None
        updates  = {it["id"]: resolve_ncash(it)
                    for it in self.csv_items if resolve_ncash(it) is not None}
        name_map = {it["id"]: it.get("name", "") for it in self.csv_items}

        file_results = []
        for fname, text in self.xml_files:
            modified, found_map = bulk_update_ncash(text, updates)
            file_results.append((fname, modified, found_map))

        found_in = {}
        for fname, _, found_map in file_results:
            for iid, hit in found_map.items():
                if hit and iid not in found_in:
                    found_in[iid] = fname

        results = []
        for item in self.csv_items:
            iid  = item["id"]
            name = name_map.get(iid, "")
            if item["ticket_cost"] is None:
                results.append((iid, name, None, None))
            else:
                results.append((iid, name, updates[iid], found_in.get(iid)))

        self._build_output_screen(file_results, results, updates)

    # ─────────────────────────────────────────────────────────────────────────
    # OUTPUT SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_output_screen(self, file_results, results, updates):
        self._clear()

        updated_count = sum(1 for _,_,n,f in results if n is not None and f is not None)
        skipped_count = sum(1 for _,_,n,_ in results if n is None)
        missing_count = sum(1 for _,_,n,f in results if n is not None and f is None)
        summary = (f"✓ Updated: {updated_count}    "
                   f"⚠ Not found in any file: {missing_count}    "
                   f"— Skipped: {skipped_count}")
        tk.Label(self, text=summary, font=("Consolas",10,"bold"),
                 bg="#1e1e2e", fg="#a6e3a1").pack(pady=8)

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=12, pady=4)

        def make_tab(title, content, fname):
            frm = tk.Frame(nb, bg="#1e1e2e"); nb.add(frm, text=title)
            br  = tk.Frame(frm, bg="#1e1e2e"); br.pack(side="bottom", fill="x")
            tk.Button(br, text="📋 Copy",
                      command=lambda c=content: (self.clipboard_clear(),
                          self.clipboard_append(c),
                          messagebox.showinfo("Copied","Copied to clipboard.")),
                      bg="#313244", fg="#cdd6f4", font=("Consolas",9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            tk.Button(br, text="💾 Save As…",
                      command=lambda c=content, f=fname: self._save(c, f),
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            txt = scrolledtext.ScrolledText(frm, font=("Consolas",9),
                                            bg="#181825", fg="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", content); txt.config(state="disabled")

        col_hdr = f"{'ID':<15}{'Name':<34}{'NCash':<13}Status"
        col_sep = "─" * 74
        exports = []

        for fname, modified_text, found_map in file_results:
            if not any(hit for hit in found_map.values()):
                continue
            exports.append((fname, modified_text))
            make_tab(os.path.splitext(fname)[0], modified_text, fname)

        # ── Update Log ───────────────────────────────────────────────────────
        log_parts = []
        for fname, _, found_map in file_results:
            file_rows = [(iid, name, ncash, ff)
                         for iid, name, ncash, ff in results if ff == fname]
            if not file_rows:
                log_parts.append(f"{fname}  →  No matching IDs — Skipped file!\n")
                continue
            log_parts.append(f"{fname}  →  {len(file_rows)} ID(s) matching CSV")
            log_parts.append("  " + ", ".join(r[0] for r in file_rows))
            log_parts.append(f"  {col_hdr}")
            log_parts.append(f"  {col_sep}")
            for iid, name, ncash, _ in file_rows:
                log_parts.append(
                    f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}✓ Updated")
            log_parts.append("")

        unassigned   = [(iid, name, ncash, ff)
                        for iid, name, ncash, ff in results if ff is None]
        skipped_rows = [(iid, name) for iid, name, ncash, _ in unassigned if ncash is None]
        missing_rows = [(iid, name, ncash) for iid, name, ncash, _ in unassigned if ncash is not None]

        log_parts.append("── Unassigned / Skipped ──────────────────────────────────────────────────")
        if missing_rows:
            log_parts.append(f"  ⚠ Not found in any file: {len(missing_rows)} ID(s)")
            log_parts.append("  " + ", ".join(r[0] for r in missing_rows))
            log_parts.append(f"  {col_hdr}")
            log_parts.append(f"  {col_sep}")
            for iid, name, ncash in missing_rows:
                log_parts.append(
                    f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}⚠ Not found")
            log_parts.append("")
        if skipped_rows:
            log_parts.append(f"  — Skipped (no cost): {len(skipped_rows)} ID(s)")
            log_parts.append("  " + ", ".join(r[0] for r in skipped_rows))
            log_parts.append(f"  {col_hdr}")
            log_parts.append(f"  {col_sep}")
            for iid, name in skipped_rows:
                log_parts.append(
                    f"  {iid:<15}{(name or '—')[:32]:<34}{'—':<13}SKIPPED")
        if not missing_rows and not skipped_rows:
            log_parts.append("  (none)")

        log_content = "\n".join(log_parts)
        exports.append(("ncash_update_log.txt", log_content))
        make_tab("Update Log", log_content, "ncash_update_log.txt")
        nb.select(0)

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x", pady=6)

        def export_all():
            folder = filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            saved = []
            for efname, content in exports:
                with open(os.path.join(folder, efname), "w", encoding="utf-8") as f:
                    f.write(content)
                saved.append(efname)
            messagebox.showinfo("Export Complete",
                f"Saved to:\n{folder}\n\n" + "\n".join(saved))

        tk.Button(bot, text="💾  Export All Files", command=export_all,
                  bg="#cba6f7", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=20, pady=8).pack(side="left", padx=14)
        tk.Button(bot, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=4)

    def _save(self, content, fname):
        p = filedialog.asksaveasfilename(initialfile=fname,
                filetypes=[("XML","*.xml"),("Text","*.txt"),("All","*.*")])
        if p:
            with open(p, "w", encoding="utf-8") as f: f.write(content)
            messagebox.showinfo("Saved", f"Saved to {p}")

    def _clear(self):
        for w in self.winfo_children(): w.destroy()


if __name__ == "__main__":
    NCashUpdaterParentApp().mainloop()
