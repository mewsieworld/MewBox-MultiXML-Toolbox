"""
SCRIPT 4 – NCash / Ticket Cost Updater  (Parent-Box CSV + PresentItemParam2 sub-box mode)
──────────────────────────────────────────────────────────────────────────────────────────
Load screen steps:
  1. Parent-Box CSV  (ID, Tickets or NCash — all other values filtered out)
     Optionally also has "Tickets of Box Contents" / "Box Contents Tickets" columns
     to trigger the sub-box flow.
  2. ItemParam XML   (pick any one; siblings auto-loaded; PresentItemParam2.xml
     silently loaded from same folder if found — never crashes if absent)
  3. Mode for parent-box IDs:  Uniform / Manual
  4. [Optional] PresentItemParam2 sub-box mode checkbox  (only shown if file loaded)
  5. [Optional] Sub-box NCash mode:  Uniform / Manual

Flow:
  Continue → parent-box Uniform or Manual screen
           → after Apply, if sub-box mode active → sub-box configure screen
           → Output screen (per-file tabs + Update Log + Export All)

Formula: NCash = round(tickets × 133)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, os

# ═══════════════════════════════════════════════════════════════════════════════
# XML HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
ROW_RE   = re.compile(r'<ROW>.*?</ROW>', re.DOTALL)
CDATA_RE = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)

def _get_tag(block, tag):
    m = re.search(rf'<{re.escape(tag)}>(.*?)</{re.escape(tag)}>', block, re.DOTALL)
    if not m: return ""
    cd = CDATA_RE.search(m.group(1))
    return cd.group(1).strip() if cd else m.group(1).strip()

def build_item_lib(files):
    """files: [(fname, text),...]. Returns {id_str: name_str}."""
    lib = {}
    for _, text in files:
        for row in ROW_RE.findall(text):
            rid  = _get_tag(row, "ID")
            name = _get_tag(row, "Name")
            if rid.isdigit() and name:
                lib[rid] = name
    return lib

def bulk_update_ncash(xml_text, updates):
    """updates: {id_str: ncash_int}. Single pass. Returns (modified, {id: found})."""
    found = {k: False for k in updates}
    def replace_row(m):
        block = m.group(0)
        rid   = _get_tag(block, "ID")
        if rid not in updates: return block
        found[rid] = True
        return re.sub(r'<Ncash>\d+</Ncash>', f'<Ncash>{updates[rid]}</Ncash>', block)
    return ROW_RE.sub(replace_row, xml_text), found

def extract_drop_ids_from_present(present_text, box_ids):
    """
    Given PresentItemParam2 XML text and a set of box_ids (strings),
    returns {box_id: [drop_id, ...]} — the non-zero DropId_# values
    for every matching <Id> row.
    """
    result = {}
    for row in ROW_RE.findall(present_text):
        bid = _get_tag(row, "Id")          # PresentItemParam2 uses <Id> not <ID>
        if bid not in box_ids:
            continue
        drops = []
        for i in range(20):
            did = _get_tag(row, f"DropId_{i}")
            if did and did.isdigit() and did != "0":
                drops.append(did)
        result[bid] = drops
    return result

# ═══════════════════════════════════════════════════════════════════════════════
# CSV PARSER
# ═══════════════════════════════════════════════════════════════════════════════
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
    "level","rate","lv","luck","lvl","chance","prob","drop","qty",
    "quantity","count","amount","row",
}

_TICKET_NAMES        = {"tickets", "ticket"}
_NCASH_NAMES         = {"ncash", "ncash_val", "ncashval"}
_BOX_TICKET_NAMES    = {
    "tickets of box contents", "box contents tickets",
    "box content tickets", "tickets of box content",
    "sub-box tickets", "subbox tickets",
}

def _find_value_col(raw_headers, id_pos):
    """Scan right from id_pos (before next ID col) for Tickets/NCash column."""
    next_id = next((i for i in range(id_pos+1, len(raw_headers))
                    if raw_headers[i].lower() == "id"), len(raw_headers))
    for i in range(id_pos+1, next_id):
        h = raw_headers[i].lower()
        if h in _TICKET_NAMES: return i, "tickets"
        if h in _NCASH_NAMES:  return i, "ncash"
    return None, None

def _find_box_ticket_col(raw_headers, id_pos):
    """Scan right from id_pos (before next ID col) for 'Tickets of Box Contents' column."""
    next_id = next((i for i in range(id_pos+1, len(raw_headers))
                    if raw_headers[i].lower() == "id"), len(raw_headers))
    for i in range(id_pos+1, next_id):
        if raw_headers[i].lower().strip() in _BOX_TICKET_NAMES:
            return i
    return None

def parse_parentbox_csv(text):
    """
    Returns list of {id, ticket_cost, ncash_direct, box_ticket_cost, name}.
      ticket_cost     — parent-box tickets (×133 later)
      ncash_direct    — parent-box NCash exact
      box_ticket_cost — sub-box contents tickets (×133 later), or None
    """
    stripped = text.strip()
    if not stripped: return []
    all_rows = list(csv.reader(io.StringIO(stripped)))
    if not all_rows: return []
    raw_headers = [h.strip() for h in all_rows[0]]
    data_rows   = all_rows[1:]

    id_positions = [i for i, h in enumerate(raw_headers) if h.lower() == "id"]
    val_map      = {}   # id_pos → (vcol, "tickets"|"ncash")
    box_tick_map = {}   # id_pos → col_index  for box-contents tickets
    for id_pos in id_positions:
        vcol, vtype = _find_value_col(raw_headers, id_pos)
        if vcol is not None:
            val_map[id_pos] = (vcol, vtype)
        btcol = _find_box_ticket_col(raw_headers, id_pos)
        if btcol is not None:
            box_tick_map[id_pos] = btcol

    items, seen = [], set()

    def add(id_str, ticket_cost, ncash_direct, box_ticket_cost):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str)
            items.append({
                "id":             id_str,
                "ticket_cost":    ticket_cost,
                "ncash_direct":   ncash_direct,
                "box_ticket_cost": box_ticket_cost,
                "name":           "",
            })

    def _parse_num(row, col):
        if col is None or col >= len(row): return None
        try:    return float(row[col].strip())
        except: return None

    if id_positions:
        for row in data_rows:
            for id_pos in id_positions:
                if id_pos >= len(row): continue
                id_val = row[id_pos].strip()
                if not (id_val and id_val.isdigit()): continue
                ticket_cost  = None
                ncash_direct = None
                if id_pos in val_map:
                    vcol, vtype = val_map[id_pos]
                    num = _parse_num(row, vcol)
                    if num is not None:
                        if vtype == "tickets": ticket_cost  = num
                        else:                  ncash_direct = int(round(num))
                btcol = box_tick_map.get(id_pos)
                box_ticket_cost = _parse_num(row, btcol) if btcol is not None else None
                add(id_val, ticket_cost, ncash_direct, box_ticket_cost)
        return items

    # Fallback: no ID header
    for row in data_rows:
        for i, cell in enumerate(row):
            hdr = raw_headers[i].lower() if i < len(raw_headers) else ""
            if hdr not in _NON_ID_HEADERS:
                add(cell, None, None, None)
    return items

# ═══════════════════════════════════════════════════════════════════════════════
# APP
# ═══════════════════════════════════════════════════════════════════════════════
TARGET_FILES    = {"itemparam2.xml","itemparamcm2.xml","itemparamex2.xml","itemparamex.xml"}
PRESENT_FILE    = "presentitemparam2.xml"

class NCashUpdaterParentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 4 – NCash Updater (Parent-Box CSV)")
        self.geometry("1000x760")
        self.configure(bg="#1e1e2e")
        self.csv_items        = []
        self.xml_files        = []       # [(fname, text)]  ItemParam files
        self.item_lib         = {}       # {id: name}
        self.present_text     = None     # str or None
        self.mode_var         = tk.StringVar(value="uniform")
        self.use_present_var  = tk.BooleanVar(value=False)
        self.sub_mode_var     = tk.StringVar(value="uniform")
        # Resolved after parent-box configure step:
        self.sub_items        = []       # [{id, ticket_cost, ncash_direct, name}]
        self._build_load_screen()

    # ─────────────────────────────────────────────────────────────────────────
    # HELPERS
    # ─────────────────────────────────────────────────────────────────────────
    def _resolve_ncash(self, it):
        if it.get("ticket_cost")  is not None: return round(it["ticket_cost"] * 133)
        if it.get("ncash_direct") is not None: return int(it["ncash_direct"])
        return None

    def _run_bulk(self, items):
        """Run bulk_update_ncash across all ItemParam files. Returns file_results, results."""
        updates  = {it["id"]: self._resolve_ncash(it) for it in items
                    if self._resolve_ncash(it) is not None}
        name_map = {it["id"]: it.get("name","") for it in items}
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
        for it in items:
            iid  = it["id"]
            name = name_map.get(iid,"")
            ncash = self._resolve_ncash(it)
            if ncash is None:
                results.append((iid, name, None, None))
            else:
                results.append((iid, name, ncash, found_in.get(iid)))
        return file_results, results

    # ─────────────────────────────────────────────────────────────────────────
    # LOAD SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="NCASH UPDATER — PARENT-BOX CSV",
                 font=("Consolas",16,"bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(20,2))
        tk.Label(self,
                 text="IDs read from every \"ID\" column. Tickets/NCash columns auto-detected.\n"
                      "\"Tickets of Box Contents\" column triggers optional sub-box mode.\n"
                      "Formula: NCash = round(tickets × 133)",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",9),
                 justify="center").pack(pady=(0,8))

        csv_status     = tk.StringVar(value="No file loaded")
        xml_status     = tk.StringVar(value="No file loaded")
        present_status = tk.StringVar(value="No file loaded")

        def section(title, color="#89b4fa"):
            f = tk.LabelFrame(self, text=f"  {title}  ", bg="#1e1e2e", fg=color,
                              font=("Consolas",10,"bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=28, pady=4)
            return f

        # ── Step 1: CSV ───────────────────────────────────────────────────────
        csv_frm = section("Step 1 — Parent-Box CSV  (ID, Tickets or NCash — all other values filtered out)")
        tk.Label(csv_frm, textvariable=csv_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas",9)).pack(side="left", padx=10)
        def load_csv():
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f: text = f.read()
            items = parse_parentbox_csv(text)
            if not items:
                messagebox.showerror("Error",
                    "No item IDs found.\nMake sure at least one column is headed \"ID\"."); return
            self.csv_items = items
            if self.item_lib:
                for it in self.csv_items: it["name"] = self.item_lib.get(it["id"],"")
            has_box_tick = any(it.get("box_ticket_cost") is not None for it in items)
            suffix = "  ✦ Box-contents tickets detected" if has_box_tick else ""
            csv_status.set(f"✓  {os.path.basename(path)}  —  {len(items)} IDs{suffix}")
        tk.Button(csv_frm, text="📂 Load", command=load_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=5)

        # ── Step 2: ItemParam ─────────────────────────────────────────────────
        xml_frm = section("Step 2 — ItemParam XML  (pick any one of the 4 files)")
        tk.Label(xml_frm, textvariable=xml_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas",9)).pack(side="left", padx=10)
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
                        with open(os.path.join(folder,fname),
                                  encoding="utf-8-sig", errors="replace") as f:
                            loaded.append((fname, f.read()))
                    except: pass
            if not loaded:
                messagebox.showerror("Error","None of the 4 ItemParam files found."); return
            self.xml_files = loaded
            self.item_lib  = build_item_lib(loaded)
            for it in self.csv_items: it["name"] = self.item_lib.get(it["id"],"")
            # Silently try to load PresentItemParam2.xml from same folder
            self.present_text = None
            present_path = os.path.join(folder, PRESENT_FILE)
            if not os.path.exists(present_path):
                # case-insensitive scan
                for fn in os.listdir(folder):
                    if fn.lower() == PRESENT_FILE:
                        present_path = os.path.join(folder, fn); break
                else:
                    present_path = None
            if present_path:
                try:
                    with open(present_path, encoding="utf-8-sig", errors="replace") as f:
                        self.present_text = f.read()
                    present_status.set(f"✓  {os.path.basename(present_path)}  (auto-loaded)")
                except:
                    present_status.set("⚠  Found but could not read PresentItemParam2.xml")
            else:
                present_status.set("Not found in folder — sub-box mode unavailable")
                self.use_present_var.set(False)
            fnames = [fn for fn,_ in loaded]
            xml_status.set(
                f"✓  {len(loaded)}/4 files  |  {len(self.item_lib)} items indexed  "
                f"({', '.join(fnames)})")
            _refresh_present_section()
        tk.Button(xml_frm, text="📂 Load", command=load_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=5)

        # ── Step 3: Mode ──────────────────────────────────────────────────────
        mode_frm = section("Step 3 — Mode  (parent-box IDs)")
        mf = tk.Frame(mode_frm, bg="#1e1e2e"); mf.pack(anchor="w", padx=10, pady=5)
        tk.Radiobutton(mf, text="Uniform  —  one value applied to every parent-box ID",
                       variable=self.mode_var, value="uniform",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10)).pack(anchor="w",pady=2)
        tk.Radiobutton(mf, text="Manual   —  set value per parent-box ID individually",
                       variable=self.mode_var, value="manual",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10)).pack(anchor="w",pady=2)

        # ── Step 4: PresentItemParam2 optional ────────────────────────────────
        pres_frm = section("Optional Step 4 — PresentItemParam2.xml — Modify sub-box content Recycle Values",
                           color="#fab387")
        pres_inner = tk.Frame(pres_frm, bg="#1e1e2e"); pres_inner.pack(fill="x", padx=10, pady=4)
        tk.Label(pres_inner, textvariable=present_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas",9)).pack(side="left")
        self._use_present_cb = tk.Checkbutton(
            pres_inner,
            text="Enable sub-box NCash update via PresentItemParam2",
            variable=self.use_present_var,
            bg="#1e1e2e", fg="#fab387", selectcolor="#313244",
            activebackground="#1e1e2e", font=("Consolas",9),
            command=lambda: _refresh_sub_mode())
        self._use_present_cb.pack(side="left", padx=14)

        # ── Step 5: Sub-box mode ──────────────────────────────────────────────
        self._sub_mode_frm = section("Optional Step 5 — PresentItemParam2.xml — Sub-box Mode",
                                     color="#fab387")
        sf = tk.Frame(self._sub_mode_frm, bg="#1e1e2e"); sf.pack(anchor="w",padx=10,pady=5)
        tk.Radiobutton(sf, text="Uniform  —  one value for all sub-box drop IDs",
                       variable=self.sub_mode_var, value="uniform",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10)).pack(anchor="w",pady=2)
        tk.Radiobutton(sf, text="Manual   —  configure each sub-box drop ID individually",
                       variable=self.sub_mode_var, value="manual",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas",10)).pack(anchor="w",pady=2)

        def _refresh_present_section():
            has_present = self.present_text is not None
            if has_present:
                self._use_present_cb.config(state="normal")
            else:
                self.use_present_var.set(False)
                self._use_present_cb.config(state="disabled")
            _refresh_sub_mode()

        def _refresh_sub_mode():
            active = self.use_present_var.get() and self.present_text is not None
            state  = "normal" if active else "disabled"
            for w in self._sub_mode_frm.winfo_children():
                try: w.config(state=state)
                except: pass
                for ww in w.winfo_children():
                    try: ww.config(state=state)
                    except: pass

        _refresh_present_section()

        def proceed():
            if not self.csv_items:
                messagebox.showwarning("Missing","Load a CSV first."); return
            if not self.xml_files:
                messagebox.showwarning("Missing","Load ItemParam XML first."); return
            if self.mode_var.get() == "uniform":
                self._build_uniform_screen()
            else:
                self._build_manual_screen()

        tk.Button(self, text="▶  Continue →", command=proceed,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",12,"bold"),
                  relief="flat", padx=20, pady=8).pack(pady=14)

    # ─────────────────────────────────────────────────────────────────────────
    # UNIFORM SCREEN  (parent-box)
    # ─────────────────────────────────────────────────────────────────────────
    def _build_uniform_screen(self):
        self._clear()
        tk.Label(self, text="Uniform Value — Parent-Box IDs",
                 font=("Consolas",14,"bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(20,4))
        tk.Label(self, text=f"Applied to all {len(self.csv_items)} IDs.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",10)).pack(pady=(0,8))

        ticket_costs  = [it["ticket_cost"]  for it in self.csv_items if it.get("ticket_cost")  is not None]
        ncash_directs = [it["ncash_direct"] for it in self.csv_items if it.get("ncash_direct") is not None]
        pre_type  = None; pre_value = None
        if ncash_directs and len(set(ncash_directs)) == 1:
            pre_type = "ncash";    pre_value = str(ncash_directs[0])
        elif ticket_costs and len(set(ticket_costs)) == 1:
            pre_type = "tickets";  pre_value = str(int(ticket_costs[0]) if ticket_costs[0]==int(ticket_costs[0]) else ticket_costs[0])

        val_type_var = tk.StringVar(value=pre_type or "tickets")
        type_frm = tk.Frame(self, bg="#1e1e2e"); type_frm.pack(pady=(0,6))
        tk.Radiobutton(type_frm, text="Tickets  (× 133 = NCash)", variable=val_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",10), command=lambda:_rp()).pack(side="left",padx=10)
        tk.Radiobutton(type_frm, text="NCash  (exact)", variable=val_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",10), command=lambda:_rp()).pack(side="left",padx=10)

        frm = tk.Frame(self, bg="#1e1e2e"); frm.pack()
        val_lbl = tk.Label(frm, text="Tickets:", bg="#1e1e2e", fg="#cdd6f4", font=("Consolas",12))
        val_lbl.pack(side="left", padx=8)
        tv = tk.StringVar(value=pre_value or "")
        tk.Entry(frm, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                 insertbackground="#cdd6f4", font=("Consolas",12), relief="flat").pack(side="left",padx=8)
        preview_var = tk.StringVar(value="NCash: —")
        tk.Label(self, textvariable=preview_var, bg="#1e1e2e", fg="#a6e3a1",
                 font=("Consolas",12,"bold")).pack(pady=8)

        def _rp(*_):
            is_t = val_type_var.get()=="tickets"
            val_lbl.config(text="Tickets:" if is_t else "NCash:")
            try:
                n = float(tv.get())
                preview_var.set(f"NCash: {round(n*133)}" if is_t else f"NCash: {int(round(n))}  (exact)")
            except: preview_var.set("NCash: —")
        tv.trace_add("write", _rp); _rp()

        def apply():
            try: num = float(tv.get())
            except: messagebox.showwarning("Invalid","Enter a valid number."); return
            is_t = val_type_var.get()=="tickets"
            for it in self.csv_items:
                if is_t: it["ticket_cost"]=num;    it["ncash_direct"]=None
                else:    it["ncash_direct"]=int(round(num)); it["ticket_cost"]=None
            self._after_parent_configured()

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(pady=16)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=8)
        tk.Button(bot, text="✓  Apply to All & Continue", command=apply,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=16, pady=8).pack(side="left", padx=8)

    # ─────────────────────────────────────────────────────────────────────────
    # MANUAL SCREEN  (parent-box)
    # ─────────────────────────────────────────────────────────────────────────
    def _build_manual_screen(self):
        self._clear()
        tk.Label(self, text="Manual Values — Parent-Box IDs",
                 font=("Consolas",14,"bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(12,2))
        tk.Label(self, text="Leave blank to skip an ID.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",9)).pack(pady=(0,4))

        outer  = tk.Frame(self, bg="#1e1e2e"); outer.pack(fill="both",expand=True,padx=20,pady=4)
        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right",fill="y"); canvas.pack(side="left",fill="both",expand=True)
        cont = tk.Frame(canvas, bg="#1e1e2e")
        wid  = canvas.create_window((0,0), window=cont, anchor="nw")
        cont.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120),"units"))

        col_type_var = tk.StringVar(value="tickets")
        has_ncash   = any(it.get("ncash_direct")  is not None for it in self.csv_items)
        has_tickets = any(it.get("ticket_cost")   is not None for it in self.csv_items)
        if has_ncash and not has_tickets: col_type_var.set("ncash")

        type_row = tk.Frame(cont, bg="#1e1e2e"); type_row.pack(anchor="w",padx=6,pady=(4,2))
        tk.Label(type_row, text="Input type:", bg="#1e1e2e", fg="#a6adc8",
                 font=("Consolas",9)).pack(side="left",padx=(0,6))
        tk.Radiobutton(type_row, text="Tickets (×133)", variable=col_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",9), command=lambda:_refresh_all()).pack(side="left",padx=4)
        tk.Radiobutton(type_row, text="NCash (exact)", variable=col_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",9), command=lambda:_refresh_all()).pack(side="left",padx=4)

        hdr = tk.Frame(cont, bg="#181825"); hdr.pack(fill="x",pady=2)
        for txt,w in [("Item ID",12),("Item Name",30),("Value",14),("NCash (result)",16)]:
            tk.Label(hdr, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas",9,"bold"), anchor="w").pack(side="left",padx=6,pady=4)

        ticket_vars  = []
        ncash_labels = []
        for i, item in enumerate(self.csv_items):
            bg  = "#1e1e2e" if i%2==0 else "#181825"
            row = tk.Frame(cont, bg=bg); row.pack(fill="x")
            tk.Label(row, text=item["id"], width=12, bg=bg, fg="#585b70",
                     font=("Consolas",9), anchor="w").pack(side="left",padx=6,pady=2)
            name = item.get("name") or self.item_lib.get(item["id"],"—")
            tk.Label(row, text=name[:32], width=30, bg=bg, fg="#a6adc8",
                     font=("Consolas",9), anchor="w").pack(side="left",padx=6,pady=2)
            tv = tk.StringVar()
            if item.get("ncash_direct") is not None and col_type_var.get()=="ncash":
                tv.set(str(item["ncash_direct"]))
            elif item.get("ticket_cost") is not None:
                tv.set(str(item["ticket_cost"]))
            ticket_vars.append(tv)
            tk.Entry(row, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas",9),
                     relief="flat").pack(side="left",padx=6,pady=2)
            ncash_lbl = tk.Label(row, text="—", width=16, bg=bg, fg="#a6e3a1",
                                 font=("Consolas",9), anchor="w")
            ncash_lbl.pack(side="left",padx=6)
            ncash_labels.append(ncash_lbl)
            def make_trace(var, lbl):
                def cb(*_):
                    try:
                        n = float(var.get())
                        if col_type_var.get()=="tickets": lbl.config(text=str(round(n*133)))
                        else: lbl.config(text=str(int(round(n)))+" (exact)")
                    except: lbl.config(text="—")
                var.trace_add("write",cb); cb()
            make_trace(tv, ncash_lbl)

        def _refresh_all():
            for var, lbl in zip(ticket_vars, ncash_labels):
                try:
                    n = float(var.get())
                    if col_type_var.get()=="tickets": lbl.config(text=str(round(n*133)))
                    else: lbl.config(text=str(int(round(n)))+" (exact)")
                except: lbl.config(text="—")

        def confirm():
            blanks = []; is_t = col_type_var.get()=="tickets"
            for i, item in enumerate(self.csv_items):
                raw = ticket_vars[i].get().strip()
                try:
                    n = float(raw)
                    if is_t: item["ticket_cost"]=n;      item["ncash_direct"]=None
                    else:    item["ncash_direct"]=int(round(n)); item["ticket_cost"]=None
                except:
                    item["ticket_cost"]=None; item["ncash_direct"]=None
                    blanks.append(item["id"])
            if blanks:
                ans = messagebox.askyesno("Missed a spot",
                    f"{len(blanks)} ID(s) have no value and will be SKIPPED:\n\n"
                    + ", ".join(blanks[:20]) + ("…" if len(blanks)>20 else "")
                    + "\n\nContinue anyway?")
                if not ans: return
            self._after_parent_configured()

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x",pady=6)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left",padx=14)
        tk.Button(bot, text="✓  Apply & Continue", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=16, pady=8).pack(side="right",padx=14)

    # ─────────────────────────────────────────────────────────────────────────
    # AFTER PARENT CONFIGURED — decide whether to do sub-box step
    # ─────────────────────────────────────────────────────────────────────────
    def _after_parent_configured(self):
        if self.use_present_var.get() and self.present_text:
            self._build_sub_configure_screen()
        else:
            self._process_and_show(self.csv_items, [])

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-BOX CONFIGURE SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_sub_configure_screen(self):
        """
        Look up every parent-box ID in PresentItemParam2, collect all DropId_#
        values, build sub_items list, then either auto-apply (uniform with
        box_ticket_cost) or show manual screen.
        """
        box_ids = {it["id"] for it in self.csv_items}
        drop_map = extract_drop_ids_from_present(self.present_text, box_ids)

        # Flatten to unique sub-item IDs, preserving first occurrence order
        # Also carry the box_ticket_cost from the parent so uniform can pre-fill
        parent_tick = {it["id"]: it.get("box_ticket_cost") for it in self.csv_items}
        seen = set(); sub_items = []
        for it in self.csv_items:
            bid  = it["id"]
            btc  = it.get("box_ticket_cost")
            for did in drop_map.get(bid, []):
                if did not in seen:
                    seen.add(did)
                    sub_items.append({
                        "id":           did,
                        "ticket_cost":  btc,    # pre-fill if CSV had box-contents tickets
                        "ncash_direct": None,
                        "name":         self.item_lib.get(did,""),
                    })
        self.sub_items = sub_items

        if not sub_items:
            messagebox.showwarning("Sub-box",
                "No DropId entries found in PresentItemParam2 for the loaded box IDs.\n"
                "Proceeding with parent-box updates only.")
            self._process_and_show(self.csv_items, [])
            return

        if self.sub_mode_var.get() == "uniform":
            self._build_sub_uniform_screen()
        else:
            self._build_sub_manual_screen()

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-BOX UNIFORM SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_sub_uniform_screen(self):
        self._clear()
        tk.Label(self, text="Uniform Value — Sub-Box Drop IDs",
                 font=("Consolas",14,"bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(20,4))
        tk.Label(self, text=f"{len(self.sub_items)} unique drop IDs found across all matched boxes.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",10)).pack(pady=(0,8))

        ticket_costs = [it["ticket_cost"] for it in self.sub_items if it.get("ticket_cost") is not None]
        pre_type = "tickets" if ticket_costs else None
        pre_value = str(ticket_costs[0]) if (ticket_costs and len(set(ticket_costs))==1) else None

        val_type_var = tk.StringVar(value=pre_type or "tickets")
        type_frm = tk.Frame(self, bg="#1e1e2e"); type_frm.pack(pady=(0,6))
        tk.Radiobutton(type_frm, text="Tickets  (× 133 = NCash)", variable=val_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",10), command=lambda:_rp()).pack(side="left",padx=10)
        tk.Radiobutton(type_frm, text="NCash  (exact)", variable=val_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",10), command=lambda:_rp()).pack(side="left",padx=10)

        frm = tk.Frame(self, bg="#1e1e2e"); frm.pack()
        val_lbl = tk.Label(frm, text="Tickets:", bg="#1e1e2e", fg="#cdd6f4", font=("Consolas",12))
        val_lbl.pack(side="left",padx=8)
        tv = tk.StringVar(value=pre_value or "")
        tk.Entry(frm, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                 insertbackground="#cdd6f4", font=("Consolas",12), relief="flat").pack(side="left",padx=8)
        preview_var = tk.StringVar(value="NCash: —")
        tk.Label(self, textvariable=preview_var, bg="#1e1e2e", fg="#a6e3a1",
                 font=("Consolas",12,"bold")).pack(pady=8)

        def _rp(*_):
            is_t = val_type_var.get()=="tickets"
            val_lbl.config(text="Tickets:" if is_t else "NCash:")
            try:
                n = float(tv.get())
                preview_var.set(f"NCash: {round(n*133)}" if is_t else f"NCash: {int(round(n))}  (exact)")
            except: preview_var.set("NCash: —")
        tv.trace_add("write",_rp); _rp()

        def apply():
            try: num = float(tv.get())
            except: messagebox.showwarning("Invalid","Enter a valid number."); return
            is_t = val_type_var.get()=="tickets"
            for it in self.sub_items:
                if is_t: it["ticket_cost"]=num;    it["ncash_direct"]=None
                else:    it["ncash_direct"]=int(round(num)); it["ticket_cost"]=None
            self._process_and_show(self.csv_items, self.sub_items)

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(pady=16)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left",padx=8)
        tk.Button(bot, text="✓  Apply to All Sub-Boxes & Update XML", command=apply,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=16, pady=8).pack(side="left",padx=8)

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-BOX MANUAL SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_sub_manual_screen(self):
        self._clear()
        tk.Label(self, text="Manual Values — Sub-Box Drop IDs",
                 font=("Consolas",14,"bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(12,2))
        tk.Label(self, text="Pre-populated from 'Tickets of Box Contents' column where available.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",9)).pack(pady=(0,4))

        outer  = tk.Frame(self, bg="#1e1e2e"); outer.pack(fill="both",expand=True,padx=20,pady=4)
        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right",fill="y"); canvas.pack(side="left",fill="both",expand=True)
        cont = tk.Frame(canvas, bg="#1e1e2e")
        wid  = canvas.create_window((0,0), window=cont, anchor="nw")
        cont.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120),"units"))

        has_ncash   = any(it.get("ncash_direct")  is not None for it in self.sub_items)
        has_tickets = any(it.get("ticket_cost")   is not None for it in self.sub_items)
        col_type_var = tk.StringVar(value="ncash" if (has_ncash and not has_tickets) else "tickets")

        type_row = tk.Frame(cont, bg="#1e1e2e"); type_row.pack(anchor="w",padx=6,pady=(4,2))
        tk.Label(type_row, text="Input type:", bg="#1e1e2e", fg="#a6adc8",
                 font=("Consolas",9)).pack(side="left",padx=(0,6))
        tk.Radiobutton(type_row, text="Tickets (×133)", variable=col_type_var, value="tickets",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",9), command=lambda:_refresh_all()).pack(side="left",padx=4)
        tk.Radiobutton(type_row, text="NCash (exact)", variable=col_type_var, value="ncash",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",9), command=lambda:_refresh_all()).pack(side="left",padx=4)

        hdr = tk.Frame(cont, bg="#181825"); hdr.pack(fill="x",pady=2)
        for txt,w in [("Drop ID",12),("Item Name",30),("Value",14),("NCash (result)",16)]:
            tk.Label(hdr, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas",9,"bold"), anchor="w").pack(side="left",padx=6,pady=4)

        ticket_vars  = []
        ncash_labels = []
        for i, item in enumerate(self.sub_items):
            bg  = "#1e1e2e" if i%2==0 else "#181825"
            row = tk.Frame(cont, bg=bg); row.pack(fill="x")
            tk.Label(row, text=item["id"], width=12, bg=bg, fg="#585b70",
                     font=("Consolas",9), anchor="w").pack(side="left",padx=6,pady=2)
            name = item.get("name") or "—"
            tk.Label(row, text=name[:32], width=30, bg=bg, fg="#a6adc8",
                     font=("Consolas",9), anchor="w").pack(side="left",padx=6,pady=2)
            tv = tk.StringVar()
            if item.get("ncash_direct") is not None and col_type_var.get()=="ncash":
                tv.set(str(item["ncash_direct"]))
            elif item.get("ticket_cost") is not None:
                tv.set(str(item["ticket_cost"]))
            ticket_vars.append(tv)
            tk.Entry(row, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas",9),
                     relief="flat").pack(side="left",padx=6,pady=2)
            ncash_lbl = tk.Label(row, text="—", width=16, bg=bg, fg="#a6e3a1",
                                 font=("Consolas",9), anchor="w")
            ncash_lbl.pack(side="left",padx=6)
            ncash_labels.append(ncash_lbl)
            def make_trace(var,lbl):
                def cb(*_):
                    try:
                        n = float(var.get())
                        if col_type_var.get()=="tickets": lbl.config(text=str(round(n*133)))
                        else: lbl.config(text=str(int(round(n)))+" (exact)")
                    except: lbl.config(text="—")
                var.trace_add("write",cb); cb()
            make_trace(tv, ncash_lbl)

        def _refresh_all():
            for var,lbl in zip(ticket_vars, ncash_labels):
                try:
                    n = float(var.get())
                    if col_type_var.get()=="tickets": lbl.config(text=str(round(n*133)))
                    else: lbl.config(text=str(int(round(n)))+" (exact)")
                except: lbl.config(text="—")

        def confirm():
            blanks=[]; is_t = col_type_var.get()=="tickets"
            for i, item in enumerate(self.sub_items):
                raw = ticket_vars[i].get().strip()
                try:
                    n = float(raw)
                    if is_t: item["ticket_cost"]=n;      item["ncash_direct"]=None
                    else:    item["ncash_direct"]=int(round(n)); item["ticket_cost"]=None
                except:
                    item["ticket_cost"]=None; item["ncash_direct"]=None
                    blanks.append(item["id"])
            if blanks:
                ans = messagebox.askyesno("Missed a spot",
                    f"{len(blanks)} drop ID(s) have no value and will be SKIPPED:\n\n"
                    + ", ".join(blanks[:20]) + ("…" if len(blanks)>20 else "")
                    + "\n\nContinue anyway?")
                if not ans: return
            self._process_and_show(self.csv_items, self.sub_items)

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x",pady=6)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left",padx=14)
        tk.Button(bot, text="✓  Apply & Update XML", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=16, pady=8).pack(side="right",padx=14)

    # ─────────────────────────────────────────────────────────────────────────
    # PROCESS — merge parent + sub items, run per file
    # ─────────────────────────────────────────────────────────────────────────
    def _process_and_show(self, parent_items, sub_items):
        # Merge: sub_items override parent if same ID (shouldn't overlap but safe)
        combined = {it["id"]: it for it in parent_items}
        for it in sub_items: combined[it["id"]] = it
        all_items = list(combined.values())

        file_results, results = self._run_bulk(all_items)

        # Annotate each result with whether it was a parent or sub item
        parent_ids = {it["id"] for it in parent_items}
        sub_ids    = {it["id"] for it in sub_items}

        self._build_output_screen(file_results, results, parent_ids, sub_ids)

    # ─────────────────────────────────────────────────────────────────────────
    # OUTPUT SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_output_screen(self, file_results, results, parent_ids, sub_ids):
        self._clear()

        updated = sum(1 for _,_,n,f in results if n is not None and f is not None)
        skipped = sum(1 for _,_,n,_ in results if n is None)
        missing = sum(1 for _,_,n,f in results if n is not None and f is None)
        p_upd   = sum(1 for iid,_,n,f in results if iid in parent_ids and n is not None and f is not None)
        s_upd   = sum(1 for iid,_,n,f in results if iid in sub_ids   and n is not None and f is not None)

        summary = f"✓ Updated: {updated}  (parent: {p_upd}, sub-box drops: {s_upd})    ⚠ Not found: {missing}    — Skipped: {skipped}"
        tk.Label(self, text=summary, font=("Consolas",9,"bold"),
                 bg="#1e1e2e", fg="#a6e3a1").pack(pady=8)

        nb = ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=12,pady=4)

        def make_tab(title, content, fname):
            frm = tk.Frame(nb, bg="#1e1e2e"); nb.add(frm, text=title)
            br  = tk.Frame(frm, bg="#1e1e2e"); br.pack(side="bottom",fill="x")
            tk.Button(br, text="📋 Copy",
                      command=lambda c=content:(self.clipboard_clear(),
                          self.clipboard_append(c),
                          messagebox.showinfo("Copied","Copied to clipboard.")),
                      bg="#313244", fg="#cdd6f4", font=("Consolas",9),
                      relief="flat", padx=10, pady=4).pack(side="left",padx=6,pady=4)
            tk.Button(br, text="💾 Save As…",
                      command=lambda c=content,f=fname:self._save(c,f),
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",9),
                      relief="flat", padx=10, pady=4).pack(side="left",padx=6,pady=4)
            txt = scrolledtext.ScrolledText(frm, font=("Consolas",9), bg="#181825", fg="#cdd6f4")
            txt.pack(fill="both",expand=True,padx=4,pady=4)
            txt.insert("1.0",content); txt.config(state="disabled")

        col_hdr = f"{'ID':<15}{'Name':<34}{'NCash':<13}{'Type':<10}Status"
        col_sep = "─" * 82
        exports = []

        for fname, modified_text, found_map in file_results:
            if not any(hit for hit in found_map.values()): continue
            exports.append((fname, modified_text))
            make_tab(os.path.splitext(fname)[0], modified_text, fname)

        # ── Update Log ────────────────────────────────────────────────────────
        log_parts = []
        def _row_type(iid):
            if iid in sub_ids:    return "sub-drop"
            if iid in parent_ids: return "parent"
            return "?"

        for fname, _, found_map in file_results:
            file_rows = [(iid,name,ncash,ff)
                         for iid,name,ncash,ff in results if ff==fname]
            if not file_rows:
                log_parts.append(f"{fname}  →  No matching IDs — Skipped file!\n"); continue
            log_parts.append(f"{fname}  →  {len(file_rows)} ID(s)")
            log_parts.append("  " + ", ".join(r[0] for r in file_rows))
            log_parts.append(f"  {col_hdr}")
            log_parts.append(f"  {col_sep}")
            for iid,name,ncash,_ in file_rows:
                log_parts.append(
                    f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}{_row_type(iid):<10}✓ Updated")
            log_parts.append("")

        unassigned   = [(iid,name,ncash,ff) for iid,name,ncash,ff in results if ff is None]
        skipped_rows = [(iid,name) for iid,name,ncash,_ in unassigned if ncash is None]
        missing_rows = [(iid,name,ncash) for iid,name,ncash,_ in unassigned if ncash is not None]

        log_parts.append("── Unassigned / Skipped ──────────────────────────────────────────────────────")
        if missing_rows:
            log_parts.append(f"  ⚠ Not found in any file: {len(missing_rows)} ID(s)")
            log_parts.append("  " + ", ".join(r[0] for r in missing_rows))
            log_parts.append(f"  {col_hdr}"); log_parts.append(f"  {col_sep}")
            for iid,name,ncash in missing_rows:
                log_parts.append(f"  {iid:<15}{(name or '—')[:32]:<34}{ncash:<13}{_row_type(iid):<10}⚠ Not found")
            log_parts.append("")
        if skipped_rows:
            log_parts.append(f"  — Skipped (no value): {len(skipped_rows)} ID(s)")
            log_parts.append("  " + ", ".join(r[0] for r in skipped_rows))
            log_parts.append(f"  {col_hdr}"); log_parts.append(f"  {col_sep}")
            for iid,name in skipped_rows:
                log_parts.append(f"  {iid:<15}{(name or '—')[:32]:<34}{'—':<13}{_row_type(iid):<10}SKIPPED")
        if not missing_rows and not skipped_rows:
            log_parts.append("  (none)")

        log_content = "\n".join(log_parts)
        exports.append(("ncash_update_log.txt", log_content))
        make_tab("Update Log", log_content, "ncash_update_log.txt")
        nb.select(0)

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x",pady=6)
        def export_all():
            folder = filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            saved = []
            for efname, content in exports:
                with open(os.path.join(folder,efname),"w",encoding="utf-8") as f:
                    f.write(content)
                saved.append(efname)
            messagebox.showinfo("Export Complete", f"Saved to:\n{folder}\n\n"+"\n".join(saved))
        tk.Button(bot, text="💾  Export All Files", command=export_all,
                  bg="#cba6f7", fg="#1e1e2e", font=("Consolas",11,"bold"),
                  relief="flat", padx=20, pady=8).pack(side="left",padx=14)
        tk.Button(bot, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left",padx=4)

    def _save(self, content, fname):
        p = filedialog.asksaveasfilename(initialfile=fname,
                filetypes=[("XML","*.xml"),("Text","*.txt"),("All","*.*")])
        if p:
            with open(p,"w",encoding="utf-8") as f: f.write(content)
            messagebox.showinfo("Saved", f"Saved to {p}")

    def _clear(self):
        for w in self.winfo_children(): w.destroy()


if __name__ == "__main__":
    NCashUpdaterParentApp().mainloop()
