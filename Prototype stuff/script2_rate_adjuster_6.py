"""
SCRIPT 2 – PresentItemParam2 Rate Adjuster  (v6)
─────────────────────────────────────────────────
The CSV this script expects contains the IDs of the BOXES THEMSELVES
(the <Id> field in PresentItemParam2), NOT the item IDs inside them.

Use the "Box ID List" CSV exported by Script 1, OR any CSV where:
  • Column 1 = Box ID  (matches <Id> in PresentItemParam2)
  • Column 2 = Box Name  (label only, optional)
  • Repeating groups allowed (multiple parent-box columns)

Two modes:
  AUTOMATIC – set Type=2, DropCnt=real item count, and apply uniform
              Rate + Count values to every used DropId slot in each matched box.
  MANUAL    – review each box one at a time; set Type, DropCnt, Rate, ItemCnt
              per slot individually with optional ItemParam library for names.

Requirements: Python 3.x  (standard library only)
Run: python script2_rate_adjuster.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, os, copy

# ═══════════════════════════════════════════════════════════════════════════════
# CSV PARSER  — reads box IDs (column 1 of each group)
# ═══════════════════════════════════════════════════════════════════════════════
_SKIP = {"id", "level", "rate", "lv", "luck", "lvl", "chance", "prob"}

def _is_name_col(h):
    return bool(h) and h.strip().lower() not in _SKIP and not h.strip().isdigit()

def parse_box_id_csv(text):
    """
    Returns {box_id_str: box_name_str}.

    Accepts:
      1. Simple 2-col:  BoxID , BoxName
      2. Repeating groups where each group starts with an ID column.
         The FIRST ID value of each group on EACH DATA ROW is a box to modify.
         (i.e. column 1 of each group = box ID, last col = name)

    The box IDs are the <Id> values in PresentItemParam2 — NOT item/drop IDs.
    """
    reader  = csv.reader(io.StringIO(text))
    rows    = list(reader)
    if not rows:
        return {}
    headers = [h.strip() for h in rows[0]]

    # Find all ID column positions
    id_positions = [i for i, h in enumerate(headers) if h.strip().lower() == "id"]
    if not id_positions:
        # Fallback: treat col 0 as ID, col 1 as name
        id_positions = [0]

    box_map = {}
    for g, id_pos in enumerate(id_positions):
        next_id = id_positions[g+1] if g+1 < len(id_positions) else len(headers)
        gcols   = list(range(id_pos, next_id))
        ghdrs   = [headers[c] for c in gcols]
        name_local = next((i for i,h in enumerate(ghdrs) if _is_name_col(h)), None)

        for row in rows[1:]:
            id_val = row[id_pos].strip() if id_pos < len(row) else ""
            if not id_val or not id_val.isdigit():
                continue
            if name_local is not None:
                nc       = gcols[name_local]
                name_val = row[nc].strip() if nc < len(row) else ""
            else:
                name_val = ""
            box_map[id_val] = name_val   # last writer wins for dupes

    return box_map

# ═══════════════════════════════════════════════════════════════════════════════
# XML HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
ROW_RE   = re.compile(r'<ROW>.*?</ROW>', re.DOTALL)
CDATA_RE = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)

def get_tag(block, tag):
    m = re.search(rf'<{re.escape(tag)}>(.*?)</{re.escape(tag)}>', block, re.DOTALL)
    if not m: return ""
    inner = m.group(1)
    cd    = CDATA_RE.search(inner)
    return cd.group(1).strip() if cd else inner.strip()

def set_tag(block, tag, val):
    # Use word-boundary anchor so DropRate_1 never matches DropRate_10
    return re.sub(
        rf'<{re.escape(tag)}>.*?</{re.escape(tag)}>',
        f'<{tag}>{val}</{tag}>',
        block, flags=re.DOTALL
    )

def real_drop_slots(block):
    """[(slot_index_int, drop_id_str), ...] for every non-zero DropId_#, sorted."""
    pairs = re.findall(r'<DropId_(\d+)>(\d+)</DropId_\d+>', block)
    return [(int(i), v) for i, v in sorted(pairs, key=lambda x: int(x[0])) if v != "0"]

def apply_cfg(block, cfg):
    """
    cfg = {"type": int, "drop_cnt": int, "slots": [{"rate":int,"count":int}, ...]}
    Writes Type, DropCnt, and per-slot DropRate_# / ItemCnt_# for every real slot.
    Empty slots (DropId=0) are not touched.
    """
    block = set_tag(block, "Type",    str(cfg["type"]))
    block = set_tag(block, "DropCnt", str(cfg["drop_cnt"]))
    for pos, (idx, _) in enumerate(real_drop_slots(block)):
        sc    = cfg["slots"][pos] if pos < len(cfg["slots"]) else {"rate": 100, "count": 1}
        block = set_tag(block, f"DropRate_{idx}", str(sc["rate"]))
        block = set_tag(block, f"ItemCnt_{idx}",  str(sc["count"]))
    return block

# ═══════════════════════════════════════════════════════════════════════════════
# ITEMPARAM LIBRARY  — reads name from <n><![CDATA[...]]></n>
# ═══════════════════════════════════════════════════════════════════════════════
def load_itemparam_folder(folder):
    lib = {}
    for fname in os.listdir(folder):
        if not fname.lower().endswith(".xml"):
            continue
        try:
            with open(os.path.join(folder, fname), encoding="utf-8-sig", errors="replace") as f:
                text = f.read()
            for row in ROW_RE.findall(text):
                rid  = get_tag(row, "ID")
                name = get_tag(row, "n")
                if rid.isdigit() and name:
                    lib[rid] = name
        except Exception:
            pass
    return lib

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 2 – PresentItemParam2 Rate Adjuster")
        self.geometry("920x720")
        self.configure(bg="#1e1e2e")
        self.csv_text    = ""
        self.xml_text    = ""
        self.item_lib    = {}
        self.mode_var    = tk.StringVar(value="automatic")
        self._rate_var   = tk.StringVar(value="100")
        self._count_var  = tk.StringVar(value="1")
        self._lib_status = tk.StringVar(value="No library loaded  (item names won't appear)")
        self._mode_panel_frame = None
        self._build_load_screen()

    # ─────────────────────────────────────────────────────────────────────────
    # LOAD SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="PRESENTITEMPARAM2 RATE ADJUSTER",
                 font=("Consolas", 16, "bold"), bg="#1e1e2e", fg="#89dceb").pack(pady=(18, 2))
        tk.Label(self,
                 text="CSV must contain the BOX IDs  (<Id> in PresentItemParam2).\n"
                      "Use the 'Box ID List' exported by Script 1, or any 2-col ID / Name CSV.",
                 bg="#1e1e2e", fg="#6c7086", font=("Consolas", 8), justify="center").pack(pady=(0, 6))

        def section(title):
            f = tk.LabelFrame(self, text=title, bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=30, pady=4)
            return f

        # ── CSV ──────────────────────────────────────────────────────────────
        csv_frm    = section("  Step 1 — Box ID CSV  ")
        csv_status = tk.StringVar(value="No file loaded")
        tk.Label(csv_frm, textvariable=csv_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)
        def load_csv():
            p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("All", "*.*")])
            if not p: return
            with open(p, encoding="utf-8-sig") as f:
                self.csv_text = f.read()
            bm = parse_box_id_csv(self.csv_text)
            csv_status.set(f"✓  {os.path.basename(p)}  ({len(bm)} box IDs found)")
        tk.Button(csv_frm, text="📂 Load CSV", command=load_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── XML ──────────────────────────────────────────────────────────────
        xml_frm    = section("  Step 2 — PresentItemParam2.xml  ")
        xml_status = tk.StringVar(value="No file loaded")
        tk.Label(xml_frm, textvariable=xml_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)
        def load_xml():
            p = filedialog.askopenfilename(filetypes=[("XML", "*.xml"), ("All", "*.*")])
            if not p: return
            with open(p, encoding="utf-8-sig") as f:
                self.xml_text = f.read()
            xml_status.set(f"✓  {os.path.basename(p)}")
        tk.Button(xml_frm, text="📂 Load XML", command=load_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── Mode ─────────────────────────────────────────────────────────────
        mode_frm = section("  Step 3 — Mode  ")
        mf = tk.Frame(mode_frm, bg="#1e1e2e")
        mf.pack(anchor="w", padx=10, pady=6)
        tk.Radiobutton(mf, text="Manual     — review and configure each box individually",
                       variable=self.mode_var, value="manual",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10),
                       command=self._refresh_mode_panel).pack(anchor="w", pady=2)
        tk.Radiobutton(mf, text="Automatic  — apply the same values to every matched box",
                       variable=self.mode_var, value="automatic",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10),
                       command=self._refresh_mode_panel).pack(anchor="w", pady=2)

        # ── Mode panel ───────────────────────────────────────────────────────
        self._mode_panel_frame = tk.Frame(self, bg="#1e1e2e")
        self._mode_panel_frame.pack(fill="x", padx=30, pady=2)
        self._refresh_mode_panel()

        tk.Button(self, text="▶  Continue →", command=self._on_continue,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 12, "bold"),
                  relief="flat", padx=20, pady=8).pack(pady=14)

    def _refresh_mode_panel(self):
        if not self._mode_panel_frame:
            return
        for w in self._mode_panel_frame.winfo_children():
            w.destroy()
        if self.mode_var.get() == "automatic":
            frm = tk.LabelFrame(self._mode_panel_frame, text="  Adjustment Values  ",
                                bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
            frm.pack(fill="x")
            def num_row(lbl, var, note=""):
                r = tk.Frame(frm, bg="#1e1e2e"); r.pack(fill="x", padx=10, pady=4)
                tk.Label(r, text=lbl, width=18, anchor="w", bg="#1e1e2e",
                         fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
                tk.Entry(r, textvariable=var, width=8, bg="#313244", fg="#cdd6f4",
                         insertbackground="#cdd6f4", font=("Consolas", 9),
                         relief="flat").pack(side="left", padx=6)
                tk.Label(r, text=note, bg="#1e1e2e", fg="#6c7086",
                         font=("Consolas", 8)).pack(side="left")
            num_row("Adjust Rate:",  self._rate_var,  "(1–32766)  applied to every used DropRate_# slot")
            num_row("Adjust Count:", self._count_var, "(1–32766)  applied to every used ItemCnt_# slot")
            tk.Label(frm,
                     text="  Type will be set to 2.  DropCnt will be set to the number of real items.",
                     bg="#1e1e2e", fg="#6c7086", font=("Consolas", 8)).pack(anchor="w", padx=10, pady=(0,6))
        else:
            frm = tk.LabelFrame(self._mode_panel_frame, text="  ItemParam Library (optional)  ",
                                bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
            frm.pack(fill="x")
            tk.Label(frm, textvariable=self._lib_status, bg="#1e1e2e",
                     fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10, pady=6)
            def load_lib():
                folder = filedialog.askdirectory(title="Select folder containing ItemParam XML files")
                if not folder: return
                self.item_lib = load_itemparam_folder(folder)
                self._lib_status.set(f"✓  {len(self.item_lib)} items from {os.path.basename(folder)}")
            tk.Button(frm, text="📂 Load Folder", command=load_lib,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

    def _on_continue(self):
        if not self.csv_text:
            messagebox.showwarning("Missing", "Please load a CSV first."); return
        if not self.xml_text:
            messagebox.showwarning("Missing", "Please load PresentItemParam2.xml first."); return

        box_map = parse_box_id_csv(self.csv_text)
        if not box_map:
            messagebox.showerror("Error", "No box IDs found in CSV."); return

        # Find matching ROW blocks by <Id> tag
        matched = []
        seen    = set()
        for row in ROW_RE.findall(self.xml_text):
            rid = get_tag(row, "Id")
            if rid in box_map and rid not in seen:
                matched.append((rid, box_map[rid], row))
                seen.add(rid)

        if not matched:
            ids_preview = ", ".join(list(box_map.keys())[:6])
            xml_preview = ", ".join([get_tag(r, "Id") for r in ROW_RE.findall(self.xml_text)[:6]])
            messagebox.showwarning("No Matches",
                f"None of the CSV box IDs matched any <Id> in the XML.\n\n"
                f"CSV IDs (first 6):  {ids_preview}\n"
                f"XML <Id>s (first 6): {xml_preview}\n\n"
                f"Make sure you're using the Box ID CSV from Script 1\n"
                f"(the IDs Script 1 assigned to each box, not the item IDs inside them).")
            return

        if self.mode_var.get() == "automatic":
            try:
                rate  = int(self._rate_var.get())
                count = int(self._count_var.get())
                if not (1 <= rate  <= 32766): raise ValueError
                if not (1 <= count <= 32766): raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Rate and Count must be integers 1–32766."); return
            self._run_automatic(matched, rate, count)
        else:
            self._run_manual(matched)

    # ─────────────────────────────────────────────────────────────────────────
    # AUTOMATIC MODE
    # ─────────────────────────────────────────────────────────────────────────
    def _run_automatic(self, matched, rate, count):
        matched_ids = {rid: row for rid, _, row in matched}
        csv_rows    = []

        def replace_row(m):
            row = m.group(0)
            rid = get_tag(row, "Id")
            if rid not in matched_ids:
                return row
            slots   = real_drop_slots(row)
            cfg     = {
                "type":     2,
                "drop_cnt": len(slots),
                "slots":    [{"rate": rate, "count": count} for _ in slots],
            }
            new_row  = apply_cfg(row, cfg)
            drop_ids = [v for _, v in real_drop_slots(new_row)]
            name     = next((n for r, n, _ in matched if r == rid), "")
            csv_rows.append([rid, name, *drop_ids])
            return new_row

        full_out = ROW_RE.sub(replace_row, self.xml_text)
        self._build_output_screen(full_out, csv_rows, len(matched))

    # ─────────────────────────────────────────────────────────────────────────
    # MANUAL MODE
    # ─────────────────────────────────────────────────────────────────────────
    def _run_manual(self, matched):
        self.manual_matched       = matched   # [(rid, name, row_block), ...]
        self.manual_idx           = 0
        self.manual_configs       = {}        # rid -> cfg
        self.manual_saved         = None
        self.manual_continue_mode = None
        self._build_manual_screen()

    def _build_manual_screen(self):
        self._clear()
        idx                      = self.manual_idx
        total                    = len(self.manual_matched)
        rid, csv_name, row_block = self.manual_matched[idx]
        slots                    = real_drop_slots(row_block)

        s          = self.manual_saved or {}
        last_type  = s.get("type",     2)
        last_dc    = s.get("drop_cnt", len(slots))
        last_slots = s.get("slots",    [])

        # ── Scroll canvas ─────────────────────────────────────────────────
        outer = tk.Frame(self, bg="#1e1e2e"); outer.pack(fill="both", expand=True)
        hdr   = tk.Frame(outer, bg="#181825"); hdr.pack(fill="x")

        hdr_txt = f"  Box {idx+1} / {total}   ID: {rid}"
        if csv_name:
            hdr_txt += f"   —   {csv_name}"
        tk.Label(hdr, text=hdr_txt, font=("Consolas", 12, "bold"),
                 bg="#181825", fg="#89dceb", pady=8).pack(side="left", padx=10)
        if self.manual_continue_mode:
            ml = "🤖 AUTO" if self.manual_continue_mode == "automate" else "👁 MONITOR"
            tk.Label(hdr, text=ml, font=("Consolas", 10),
                     bg="#181825", fg="#fab387").pack(side="right", padx=15)

        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        sb     = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        cont = tk.Frame(canvas, bg="#1e1e2e")
        wid  = canvas.create_window((0, 0), window=cont, anchor="nw")
        cont.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        def section(title):
            f = tk.LabelFrame(cont, text=title, bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=12, pady=5); return f

        # ── Drop Type ─────────────────────────────────────────────────────
        sec_type = section("  Drop Type  ")
        type_var = tk.IntVar(value=last_type)
        tf = tk.Frame(sec_type, bg="#1e1e2e"); tf.pack(anchor="w", padx=8, pady=6)
        tk.Radiobutton(tf, text="Random      (Type = 0)", variable=type_var, value=0,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(side="left", padx=8)
        tk.Radiobutton(tf, text="Egalitarian (Type = 2)", variable=type_var, value=2,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(side="left", padx=8)

        dc_frame = tk.Frame(sec_type, bg="#1e1e2e"); dc_frame.pack(anchor="w", padx=8, pady=(0,6))
        dc_lbl   = tk.Label(dc_frame, text="Types of Items (DropCnt):", bg="#1e1e2e",
                            fg="#cdd6f4", font=("Consolas", 9))
        dc_var   = tk.StringVar(value=str(last_dc))
        dc_ent   = tk.Entry(dc_frame, textvariable=dc_var, width=6, bg="#313244",
                            fg="#cdd6f4", insertbackground="#cdd6f4",
                            font=("Consolas", 9), relief="flat")
        dc_auto  = tk.Label(dc_frame,
                            text=f"DropCnt auto-set to {len(slots)}  (all items in this box)",
                            bg="#1e1e2e", fg="#6c7086", font=("Consolas", 9))

        def toggle_dc(*_):
            if type_var.get() == 0:
                dc_auto.pack_forget()
                dc_lbl.pack(side="left"); dc_ent.pack(side="left", padx=6)
            else:
                dc_lbl.pack_forget(); dc_ent.pack_forget()
                dc_auto.pack(anchor="w")
        type_var.trace_add("write", toggle_dc); toggle_dc()

        # ── Slots ─────────────────────────────────────────────────────────
        sec_slots = section(f"  Drop Slots  ({len(slots)} items)  ")

        # Header row
        hrow = tk.Frame(sec_slots, bg="#181825"); hrow.pack(fill="x", padx=8, pady=2)
        for txt, w in [("Slot #", 6), ("Drop ID", 12), ("Item Name", 34), ("Rate %", 9), ("Item Count", 10)]:
            tk.Label(hrow, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas", 9, "bold"), anchor="w").pack(side="left", padx=3)

        slot_rate_vars  = []
        slot_count_vars = []

        for pos, (sidx, drop_id) in enumerate(slots):
            bg   = "#1e1e2e" if pos % 2 == 0 else "#181825"
            srow = tk.Frame(sec_slots, bg=bg); srow.pack(fill="x", padx=8, pady=1)

            prev_r = last_slots[pos]["rate"]  if pos < len(last_slots) else 100
            prev_c = last_slots[pos]["count"] if pos < len(last_slots) else 1

            # Slot index (greyed)
            tk.Label(srow, text=str(sidx), width=6, bg=bg, fg="#6c7086",
                     font=("Consolas", 9)).pack(side="left", padx=3)
            # Drop ID (greyed)
            tk.Label(srow, text=drop_id, width=12, bg=bg, fg="#585b70",
                     font=("Consolas", 9)).pack(side="left", padx=3)
            # Item name from library
            name = self.item_lib.get(drop_id, "—")
            tk.Label(srow, text=name[:36], width=34, bg=bg, fg="#a6adc8",
                     font=("Consolas", 9), anchor="w").pack(side="left", padx=3)

            rv = tk.StringVar(value=str(prev_r))
            tk.Entry(srow, textvariable=rv, width=7, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas", 9),
                     relief="flat").pack(side="left", padx=3)
            slot_rate_vars.append(rv)

            cv = tk.StringVar(value=str(prev_c))
            tk.Entry(srow, textvariable=cv, width=7, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas", 9),
                     relief="flat").pack(side="left", padx=3)
            slot_count_vars.append(cv)

        # ── Gather ────────────────────────────────────────────────────────
        def gather():
            t  = type_var.get()
            dc = len(slots) if t == 2 else max(1, int(dc_var.get() or 1))
            sl = []
            for rv, cv in zip(slot_rate_vars, slot_count_vars):
                try:    r = max(1, min(32766, int(rv.get())))
                except: r = 100
                try:    c = max(1, min(32766, int(cv.get())))
                except: c = 1
                sl.append({"rate": r, "count": c})
            return {"type": t, "drop_cnt": dc, "slots": sl}

        def save_and_advance():
            cfg = gather()
            self.manual_configs[rid] = cfg
            self.manual_saved        = cfg
            self.manual_idx         += 1
            if self.manual_idx >= total:
                self._finish_manual()
            elif self.manual_continue_mode == "automate":
                self._automate_manual_remaining(cfg)
            elif self.manual_continue_mode == "monitor":
                self._build_manual_screen()
            else:
                self._ask_manual_mode(cfg)

        # ── Nav ───────────────────────────────────────────────────────────
        nav = tk.Frame(cont, bg="#181825"); nav.pack(fill="x", pady=8)

        tk.Button(nav, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=8, pady=6)

        if idx > 0:
            def go_prev():
                self.manual_idx -= 1
                self._build_manual_screen()
            tk.Button(nav, text="◀  Prev", command=go_prev,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                      relief="flat", padx=12, pady=6).pack(side="left", padx=4, pady=6)

        if self.manual_continue_mode:
            def change_mode():
                self.manual_continue_mode = None
                self._build_manual_screen()
            tk.Button(nav, text="⚙ Change Mode", command=change_mode,
                      bg="#45475a", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=8, pady=6).pack(side="left", padx=4, pady=6)

        next_lbl = "Finish ✓" if idx == total - 1 else "Next ▶"
        tk.Button(nav, text=next_lbl, command=save_and_advance,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 10, "bold"),
                  relief="flat", padx=12, pady=6).pack(side="right", padx=8, pady=6)

    # ─────────────────────────────────────────────────────────────────────────
    def _ask_manual_mode(self, last_cfg):
        remaining = len(self.manual_matched) - self.manual_idx
        win = tk.Toplevel(self); win.title("Continue?")
        win.geometry("520x240"); win.configure(bg="#1e1e2e"); win.grab_set()
        tk.Label(win, text=f"{remaining} box(es) remaining.",
                 bg="#1e1e2e", fg="#cdd6f4", font=("Consolas", 13, "bold")).pack(pady=12)
        remember = tk.BooleanVar(value=False)
        bf = tk.Frame(win, bg="#1e1e2e"); bf.pack(pady=8)
        def choose(mode):
            if remember.get(): self.manual_continue_mode = mode
            win.destroy()
            if mode == "automate": self._automate_manual_remaining(last_cfg)
            else: self._build_manual_screen()
        tk.Button(bf, text="🤖  Automate  —  copy settings to all remaining boxes",
                  command=lambda: choose("automate"), bg="#cba6f7", fg="#1e1e2e",
                  font=("Consolas", 10), relief="flat", padx=10, pady=8).pack(pady=4)
        tk.Button(bf, text="👁  Monitor  —  review each box",
                  command=lambda: choose("monitor"), bg="#89b4fa", fg="#1e1e2e",
                  font=("Consolas", 10), relief="flat", padx=10, pady=8).pack(pady=4)
        tk.Checkbutton(win, text="Remember for rest of session", variable=remember,
                       bg="#1e1e2e", fg="#fab387", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(pady=4)

    def _automate_manual_remaining(self, last_cfg):
        while self.manual_idx < len(self.manual_matched):
            rid, _, row_block = self.manual_matched[self.manual_idx]
            slots = real_drop_slots(row_block)
            cfg   = copy.deepcopy(last_cfg)
            # Extend/trim slot list to match this box
            while len(cfg["slots"]) < len(slots):
                cfg["slots"].append(cfg["slots"][-1] if cfg["slots"] else {"rate": 100, "count": 1})
            cfg["slots"] = cfg["slots"][:len(slots)]
            if cfg["type"] == 2:
                cfg["drop_cnt"] = len(slots)
            self.manual_configs[rid] = cfg
            self.manual_idx += 1
        self._finish_manual()

    def _finish_manual(self):
        csv_rows = []
        def replace_row(m):
            row = m.group(0)
            rid = get_tag(row, "Id")
            if rid not in self.manual_configs:
                return row
            new_row  = apply_cfg(row, self.manual_configs[rid])
            drop_ids = [v for _, v in real_drop_slots(new_row)]
            name     = next((n for r, n, _ in self.manual_matched if r == rid), "")
            csv_rows.append([rid, name, *drop_ids])
            return new_row
        full_out = ROW_RE.sub(replace_row, self.xml_text)
        self._build_output_screen(full_out, csv_rows, len(self.manual_configs))

    # ─────────────────────────────────────────────────────────────────────────
    # OUTPUT SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_output_screen(self, full_xml, csv_rows, count):
        self._clear()
        tk.Label(self, text=f"Done — {count} box(es) modified",
                 font=("Consolas", 13, "bold"), bg="#1e1e2e", fg="#a6e3a1").pack(pady=12)
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=12, pady=4)

        max_items = max((len(r) - 2 for r in csv_rows), default=0)
        header    = ["BoxID", "BoxName"] + [f"Item{i+1}_ID" for i in range(max_items)]
        csv_cont  = "\n".join([",".join(header)] +
                              [",".join(str(x) for x in r) for r in csv_rows])

        def make_tab(title, content, fname):
            frm = tk.Frame(nb, bg="#1e1e2e"); nb.add(frm, text=title)
            br  = tk.Frame(frm, bg="#1e1e2e"); br.pack(side="bottom", fill="x")
            tk.Button(br, text="📋 Copy All",
                      command=lambda c=content: (self.clipboard_clear(),
                          self.clipboard_append(c),
                          messagebox.showinfo("Copied", "Copied to clipboard.")),
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            tk.Button(br, text="💾 Save As…",
                      command=lambda c=content, f=fname: self._save(c, f),
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            txt = scrolledtext.ScrolledText(frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", content); txt.config(state="disabled")

        make_tab("Full PresentItemParam2.xml (modified)", full_xml,  "PresentItemParam2_modified.xml")
        make_tab("Box Contents CSV (for Script 3)",       csv_cont,  "box_contents_for_script3.csv")
        nb.select(0)

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _save(self, content, fname):
        p = filedialog.asksaveasfilename(initialfile=fname, defaultextension=".xml",
                filetypes=[("XML", "*.xml"), ("CSV", "*.csv"), ("All", "*.*")])
        if p:
            with open(p, "w", encoding="utf-8") as f: f.write(content)
            messagebox.showinfo("Saved", f"Saved to {p}")

    def _clear(self):
        for w in self.winfo_children(): w.destroy()


if __name__ == "__main__":
    App().mainloop()
