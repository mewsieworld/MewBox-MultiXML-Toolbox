"""
SCRIPT 1 - Box XML Generator
Reads a CSV (groups of 3 cols: ID, Level, ParentBoxName), lets you configure
each box's ItemParam and PresentItemParam2 settings, then exports XML rows.

Requirements: Python 3.x (tkinter is included in standard Python installs)
Run: python script1_box_generator.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, os, math, copy

# ─── Character ChrTypeFlag table ────────────────────────────────────────────
CHR_FLAGS = [
    ("Bunny 1st",   1),       ("Buffalo 1st", 2),    ("Sheep 1st",   4),
    ("Dragon 1st",  8),       ("Fox 1st",     16),   ("Lion 1st",    32),
    ("Cat 1st",     64),      ("Raccoon 1st", 124),  ("Paula 1st",   256),
    ("Bunny 2nd",   512),     ("Buffalo 2nd", 1024), ("Sheep 2nd",   2048),
    ("Dragon 2nd",  4096),    ("Fox 2nd",     8192), ("Lion 2nd",    16384),
    ("Cat 2nd",     32768),   ("Raccoon 2nd", 65536),("Paula 2nd",   131072),
    ("Bunny 3rd",   262144),  ("Buffalo 3rd", 524288),("Sheep 3rd",  1048576),
    ("Dragon 3rd",  2097152), ("Fox 3rd",     4194304),("Lion 3rd",  8388608),
    ("Cat 3rd",     16777216),("Raccoon 3rd", 33554432),("Paula 3rd",67108864),
]

OPTIONS_CHECKS = [
    ("Not Buyable",      256),
    ("Not Sellable",     512),
    ("Not Exchangeable", 1024),
    ("Not Pickable",     2048),
    ("Not Droppable",    4096),
    ("Not Vanishable",   8192),
    ("No Angelina Bank", 65536),
    ("No Lisa Bank",     131072),
]

# ─── CSV Parser ──────────────────────────────────────────────────────────────
def parse_csv_text(text):
    """Parse CSV with groups of 3 cols: ID, Level/Rate, BoxName (header=parent box)."""
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        return []
    headers = rows[0]
    groups = []
    i = 0
    while i < len(headers):
        box_name = headers[i + 2].strip() if i + 2 < len(headers) else ""
        if box_name:
            items = []
            for row in rows[1:]:
                id_val  = row[i].strip()     if i     < len(row) else ""
                lv_val  = row[i+1].strip()   if i+1   < len(row) else ""
                # Check for a rate column label vs level — but store whatever's there
                if id_val and id_val.isdigit():
                    # strip any % or non-numeric from level/rate column
                    rate_clean = "".join(c for c in lv_val if c.isdigit())
                    items.append({
                        "id":    id_val,
                        "level": lv_val,
                        "rate":  int(rate_clean) if rate_clean else None,
                        "name":  row[i+2].strip() if i+2 < len(row) else "",
                    })
            groups.append({"box_name": box_name, "items": items})
        i += 3
    return groups

# ─── XML Builders ────────────────────────────────────────────────────────────
def build_options_str(check_states, recycle_val):
    base = [2, 32]
    checked = [v for (_, v), on in zip(OPTIONS_CHECKS, check_states) if on]
    rec = [recycle_val] if recycle_val > 0 else []
    all_vals = sorted(set(base + checked + rec))
    return "/".join(str(x) for x in all_vals)

def build_itemparam_row(cfg):
    opts = build_options_str(cfg["opt_checks"], cfg["opt_recycle"])
    chr_flags = sum(cfg["chr_type_flags"])
    ncash = cfg["ncash"]
    return f"""<ROW>
<ID>{cfg['id']}</ID>
<Class>1</Class>
<Type>15</Type>
<SubType>0</SubType>
<ItemFType>0</ItemFType>
<Name><![CDATA[{cfg['name']}]]></Name>
<Comment><![CDATA[{cfg['comment']}]]></Comment>
<Use><![CDATA[{cfg['use']}]]></Use>
<Name_Eng><![CDATA[ ]]></Name_Eng>
<Comment_Eng><![CDATA[ ]]></Comment_Eng>
<FileName><![CDATA[{cfg['file_name']}]]></FileName>
<BundleNum>0</BundleNum>
<InvFileName><![CDATA[{cfg['inv_file_name']}]]></InvFileName>
<InvBundleNum>0</InvBundleNum>
<CmtFileName><![CDATA[{cfg['cmt_file_name']}]]></CmtFileName>
<CmtBundleNum>0</CmtBundleNum>
<EquipFileName><![CDATA[ ]]></EquipFileName>
<PivotID>0</PivotID>
<PaletteId>0</PaletteId>
<Options>{opts}</Options>
<HideHat>0</HideHat>
<ChrTypeFlags>{chr_flags}</ChrTypeFlags>
<GroundFlags>0</GroundFlags>
<SystemFlags>0</SystemFlags>
<OptionsEx>0</OptionsEx>
<Weight>{cfg['weight']}</Weight>
<Value>{cfg['value']}</Value>
<MinLevel>{cfg['min_level']}</MinLevel>
<Effect>22</Effect>
<EffectFlags2>0</EffectFlags2>
<SelRange>0</SelRange>
<Life>0</Life>
<Depth>0</Depth>
<Delay>0.000000</Delay>
<AP>0</AP>
<HP>0</HP>
<HPCon>0</HPCon>
<MP>0</MP>
<MPCon>0</MPCon>
<Money>{cfg['money']}</Money>
<APPlus>0</APPlus>
<ACPlus>0</ACPlus>
<DXPlus>0</DXPlus>
<MaxMPPlus>0</MaxMPPlus>
<MAPlus>0</MAPlus>
<MDPlus>0</MDPlus>
<MaxWTPlus>0</MaxWTPlus>
<DAPlus>0</DAPlus>
<LKPlus>0</LKPlus>
<MaxHPPlus>0</MaxHPPlus>
<DPPlus>0</DPPlus>
<HVPlus>0</HVPlus>
<HPRecoveryRate>0.000000</HPRecoveryRate>
<MPRecoveryRate>0.000000</MPRecoveryRate>
<CardNum>0</CardNum>
<CardGenGrade>0</CardGenGrade>
<CardGenParam>0.000000</CardGenParam>
<DailyGenCnt>0</DailyGenCnt>
<PartFileName><![CDATA[ ]]></PartFileName>
<ChrFTypeFlag>0</ChrFTypeFlag>
<ChrGender>0</ChrGender>
<ExistType>0</ExistType>
<Ncash>{ncash}</Ncash>
<NewCM>0</NewCM>
<FamCM>0</FamCM>
<Summary><![CDATA[ ]]></Summary>
<ShopFileName><![CDATA[ ]]></ShopFileName>
<ShopBundleNum>0</ShopBundleNum>
<MinStatType>0</MinStatType>
<MinStatLv>0</MinStatLv>
<RefineIndex>0</RefineIndex>
<RefineType>0</RefineType>
<CompoundSlot>0</CompoundSlot>
<SetItemID>0</SetItemID>
<ReformCount>0</ReformCount>
<GroupId>0</GroupId>
</ROW>"""

def build_presentparam_row(box_id, items, ptype, drop_cnt, default_rate):
    is_distrib = (ptype == 2)
    actual_drop_cnt = len(items) if is_distrib else drop_cnt
    lines = [f"<ROW>", f"<Id>{box_id}</Id>", f"<Type>{ptype}</Type>",
             f"<DropCnt>{actual_drop_cnt}</DropCnt>"]
    for i in range(20):
        if i < len(items):
            did   = items[i]["id"]
            drate = 100 if is_distrib else (items[i]["rate"] if items[i]["rate"] is not None else default_rate)
        else:
            did, drate = 0, 0
        lines += [f"<DropId_{i}>{did}</DropId_{i}>",
                  f"<DropRate_{i}>{drate}</DropRate_{i}>",
                  f"<ItemCnt_{i}>0</ItemCnt_{i}>"]
    lines.append("</ROW>")
    return "\n".join(lines)

# ─── Main Application ────────────────────────────────────────────────────────
class BoxGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 1 – Box XML Generator")
        self.geometry("950x750")
        self.configure(bg="#1e1e2e")

        # Saved settings carried across boxes
        self.saved_settings = None
        self.groups = []          # parsed CSV groups
        self.current_group_idx = 0
        self.box_configs = []     # filled configs per group

        # Last-used filepath memory per field
        self.last_file_name     = tk.StringVar(value=r"data\item\itm_pre_107.nri")
        self.last_inv_file_name = tk.StringVar(value=r"data\item\itm_pre_107.nri")
        self.last_cmt_file_name = tk.StringVar(value=r"data\item\itm_pre_illu_107.nri")

        self._build_load_screen()

    # ── Screen 0: Load CSV ────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        frm = tk.Frame(self, bg="#1e1e2e")
        frm.pack(expand=True)

        tk.Label(frm, text="BOX XML GENERATOR", font=("Consolas", 20, "bold"),
                 bg="#1e1e2e", fg="#cba6f7").pack(pady=(30, 5))
        tk.Label(frm, text="Load a CSV with groups of 3 columns: ID | Level/Rate | Parent Box Name",
                 bg="#1e1e2e", fg="#a6adc8").pack(pady=5)

        btn_frm = tk.Frame(frm, bg="#1e1e2e")
        btn_frm.pack(pady=15)
        tk.Button(btn_frm, text="📂  Load CSV File", command=self._load_csv_file,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 11), relief="flat",
                  padx=16, pady=8).pack(side="left", padx=8)
        tk.Button(btn_frm, text="📋  Paste CSV Text", command=self._paste_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 11), relief="flat",
                  padx=16, pady=8).pack(side="left", padx=8)

    def _load_csv_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"),("All","*.*")])
        if not path:
            return
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()
        self._process_csv(text)

    def _paste_csv(self):
        win = tk.Toplevel(self)
        win.title("Paste CSV")
        win.geometry("600x400")
        win.configure(bg="#1e1e2e")
        tk.Label(win, text="Paste CSV content below:", bg="#1e1e2e",
                 fg="#cdd6f4", font=("Consolas", 10)).pack(anchor="w", padx=10, pady=5)
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9))
        txt.pack(fill="both", expand=True, padx=10, pady=5)
        def confirm():
            self._process_csv(txt.get("1.0", "end"))
            win.destroy()
        tk.Button(win, text="Confirm", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _process_csv(self, text):
        groups = parse_csv_text(text)
        if not groups:
            messagebox.showerror("Error", "No valid box groups found in CSV.")
            return
        self.groups = groups
        self.current_group_idx = 0
        self.box_configs = []
        self._build_config_screen()

    # ── Screen 1: Configure a box ─────────────────────────────────────────
    def _build_config_screen(self):
        self._clear()
        idx = self.current_group_idx
        grp = self.groups[idx]
        box_name = grp["box_name"]
        total    = len(self.groups)

        # ── Outer scroll canvas ──
        outer = tk.Frame(self, bg="#1e1e2e")
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg="#181825")
        header.pack(fill="x")
        tk.Label(header,
                 text=f"Box {idx+1}/{total}:  {box_name}",
                 font=("Consolas", 14, "bold"), bg="#181825", fg="#cba6f7",
                 pady=8).pack(side="left", padx=15)

        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        container = tk.Frame(canvas, bg="#1e1e2e")
        win_id = canvas.create_window((0, 0), window=container, anchor="nw")

        def on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win_id, width=canvas.winfo_width())
        container.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120),"units"))

        # ── Section helper ──
        def section(parent, title):
            f = tk.LabelFrame(parent, text=title, bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=12, pady=6)
            return f

        def lbl_entry(parent, label, var, width=40):
            r = tk.Frame(parent, bg="#1e1e2e")
            r.pack(fill="x", padx=8, pady=2)
            tk.Label(r, text=label, width=22, anchor="w", bg="#1e1e2e",
                     fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
            tk.Entry(r, textvariable=var, width=width,
                     bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                     font=("Consolas", 9), relief="flat").pack(side="left", padx=4)
            return r

        # ── Variables ──
        v = {}
        v["id"]       = tk.StringVar(value="")
        v["name"]     = tk.StringVar(value=box_name)
        v["comment"]  = tk.StringVar(value="Special dice that contains amazing items.")
        v["use"]      = tk.StringVar(value="Event Box.")
        v["file_name"]     = tk.StringVar(value=self.last_file_name.get())
        v["inv_file_name"] = tk.StringVar(value=self.last_inv_file_name.get())
        v["cmt_file_name"] = tk.StringVar(value=self.last_cmt_file_name.get())
        v["weight"]    = tk.StringVar(value="1")
        v["value"]     = tk.StringVar(value="0")
        v["min_level"] = tk.StringVar(value="1")
        v["money"]     = tk.StringVar(value="0")
        v["ticket"]    = tk.StringVar(value="0")  # user inputs ticket value

        # Options checkboxes
        opt_check_vars = [tk.BooleanVar() for _ in OPTIONS_CHECKS]
        opt_recycle_var = tk.IntVar(value=0)

        # ChrTypeFlags
        chr_selected = []  # list of vals

        # Present param vars
        present_type_var    = tk.IntVar(value=0)
        drop_cnt_var        = tk.StringVar(value="1")
        default_rate_var    = tk.StringVar(value="50")
        remember_present    = tk.BooleanVar(value=False)

        # Per-item rate overrides (list of StringVar, length = items)
        item_rate_vars = [tk.StringVar(value=str(it["rate"]) if it["rate"] else "50")
                          for it in grp["items"]]

        # ── If saved settings exist, prefill ──
        if self.saved_settings:
            s = self.saved_settings
            v["comment"].set(s.get("comment", v["comment"].get()))
            v["use"].set(s.get("use", v["use"].get()))
            v["file_name"].set(s.get("file_name", v["file_name"].get()))
            v["inv_file_name"].set(s.get("inv_file_name", v["inv_file_name"].get()))
            v["cmt_file_name"].set(s.get("cmt_file_name", v["cmt_file_name"].get()))
            v["weight"].set(s.get("weight", "1"))
            v["value"].set(s.get("value", "0"))
            v["min_level"].set(s.get("min_level", "1"))
            v["money"].set(s.get("money", "0"))
            v["ticket"].set(s.get("ticket", "0"))
            for i, bv in enumerate(opt_check_vars):
                bv.set(s.get("opt_checks", [False]*8)[i])
            opt_recycle_var.set(s.get("opt_recycle", 0))
            chr_selected = list(s.get("chr_type_flags", []))
            if s.get("remember_present"):
                present_type_var.set(s.get("present_type", 0))
                drop_cnt_var.set(s.get("drop_cnt", "1"))
                default_rate_var.set(s.get("default_rate", "50"))
                remember_present.set(True)

        # ── Build UI sections ──

        # Basic Info
        sec_basic = section(container, "  ItemParam – Basic Info  ")
        lbl_entry(sec_basic, "Box ID (itemparam):", v["id"], 20)
        tk.Label(sec_basic, text="  (Name auto-filled from CSV header)",
                 bg="#1e1e2e", fg="#6c7086", font=("Consolas", 8)).pack(anchor="w", padx=8)
        lbl_entry(sec_basic, "Name (CDATA):",    v["name"])
        lbl_entry(sec_basic, "Comment (CDATA):", v["comment"])
        lbl_entry(sec_basic, "Use (CDATA):",     v["use"])

        # Filepaths
        sec_files = section(container, "  Filepaths (CDATA, data\\...\\*.nri)  ")
        lbl_entry(sec_files, "FileName:",    v["file_name"])
        lbl_entry(sec_files, "InvFileName:", v["inv_file_name"])
        lbl_entry(sec_files, "CmtFileName:", v["cmt_file_name"])

        # Options
        sec_opts = section(container, "  Options  ")
        chk_frm = tk.Frame(sec_opts, bg="#1e1e2e")
        chk_frm.pack(anchor="w", padx=8, pady=4)
        for i, (lbl, _) in enumerate(OPTIONS_CHECKS):
            col = i % 4
            row = i // 4
            tk.Checkbutton(chk_frm, text=lbl, variable=opt_check_vars[i],
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e", font=("Consolas", 9)
                           ).grid(row=row, column=col, sticky="w", padx=6, pady=2)

        rec_frm = tk.Frame(sec_opts, bg="#1e1e2e")
        rec_frm.pack(anchor="w", padx=8, pady=4)
        tk.Label(rec_frm, text="Recycle:", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Consolas", 9)).pack(side="left", padx=(0, 8))
        for lbl, val in [("None", 0), ("Recyclable", 262144), ("Non-Recyclable", 8388608)]:
            tk.Radiobutton(rec_frm, text=lbl, variable=opt_recycle_var, value=val,
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e", font=("Consolas", 9)
                           ).pack(side="left", padx=6)

        # ChrTypeFlags
        sec_chr = section(container, "  ChrTypeFlags (up to 24 entries)  ")
        chr_display_var = tk.StringVar(value="None")
        chr_listbox_items = []

        chr_ctrl = tk.Frame(sec_chr, bg="#1e1e2e")
        chr_ctrl.pack(fill="x", padx=8, pady=4)

        chr_lb_frame = tk.Frame(chr_ctrl, bg="#1e1e2e")
        chr_lb_frame.pack(side="left")
        tk.Label(chr_lb_frame, text="Added flags:", bg="#1e1e2e",
                 fg="#89b4fa", font=("Consolas", 9)).pack(anchor="w")
        chr_lb = tk.Listbox(chr_lb_frame, height=5, width=30,
                            bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                            selectbackground="#45475a")
        chr_lb.pack()

        def refresh_chr_lb():
            chr_lb.delete(0, "end")
            for val in chr_selected:
                name = next((n for n, v2 in CHR_FLAGS if v2 == val), str(val))
                chr_lb.insert("end", name)

        refresh_chr_lb()

        add_frm = tk.Frame(chr_ctrl, bg="#1e1e2e")
        add_frm.pack(side="left", padx=16)
        tk.Label(add_frm, text="Add character flag:", bg="#1e1e2e",
                 fg="#89b4fa", font=("Consolas", 9)).pack(anchor="w")
        chr_combo = ttk.Combobox(add_frm, values=[n for n, _ in CHR_FLAGS],
                                 state="readonly", width=22, font=("Consolas", 9))
        chr_combo.pack(pady=2)

        def add_chr():
            sel = chr_combo.get()
            if not sel:
                return
            val = next((v2 for n, v2 in CHR_FLAGS if n == sel), None)
            if val and val not in chr_selected and len(chr_selected) < 24:
                chr_selected.append(val)
                refresh_chr_lb()

        def rem_chr():
            sel = chr_lb.curselection()
            if sel:
                chr_selected.pop(sel[0])
                refresh_chr_lb()

        tk.Button(add_frm, text="Add", command=add_chr,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                  relief="flat", padx=8).pack(side="left", pady=4, padx=2)
        tk.Button(add_frm, text="Remove Selected", command=rem_chr,
                  bg="#f38ba8", fg="#1e1e2e", font=("Consolas", 9),
                  relief="flat", padx=8).pack(side="left", pady=4, padx=2)

        # Numeric fields
        sec_nums = section(container, "  Numeric Fields  ")
        rf = tk.Frame(sec_nums, bg="#1e1e2e")
        rf.pack(fill="x", padx=8, pady=4)
        for i, (lbl, var) in enumerate([("Weight:", v["weight"]), ("Value:", v["value"]),
                                         ("MinLevel:", v["min_level"]), ("Money:", v["money"])]):
            tk.Label(rf, text=lbl, bg="#1e1e2e", fg="#cdd6f4",
                     font=("Consolas", 9), width=10, anchor="w").grid(row=0, column=i*2, padx=4)
            tk.Entry(rf, textvariable=var, width=10,
                     bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                     font=("Consolas", 9), relief="flat").grid(row=0, column=i*2+1, padx=4)

        # NCash ticket
        sec_nc = section(container, "  NCash / Ticket Cost  ")
        nc_frm = tk.Frame(sec_nc, bg="#1e1e2e")
        nc_frm.pack(fill="x", padx=8, pady=4)
        tk.Label(nc_frm, text="Ticket value (NCash = tickets × 133, rounded):",
                 bg="#1e1e2e", fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
        tk.Entry(nc_frm, textvariable=v["ticket"], width=10,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 font=("Consolas", 9), relief="flat").pack(side="left", padx=8)
        ncash_lbl = tk.Label(nc_frm, text="→ NCash: 0", bg="#1e1e2e",
                             fg="#a6e3a1", font=("Consolas", 9))
        ncash_lbl.pack(side="left")

        def update_ncash(*_):
            try:
                tickets = float(v["ticket"].get())
                ncash = round(tickets * 133)
                ncash_lbl.config(text=f"→ NCash: {ncash}")
            except:
                ncash_lbl.config(text="→ NCash: ?")
        v["ticket"].trace_add("write", update_ncash)
        update_ncash()

        # ── PresentItemParam2 section ──
        sec_pres = section(container, "  PresentItemParam2 Settings  ")
        pk = tk.Frame(sec_pres, bg="#1e1e2e")
        pk.pack(fill="x", padx=8, pady=4)

        id_disp = tk.Label(pk, text=f"Box ID: (fill above)  |  {box_name}",
                           bg="#1e1e2e", fg="#6c7086", font=("Consolas", 9))
        id_disp.pack(anchor="w")

        # Update ID display when ID var changes
        def update_id_disp(*_):
            id_disp.config(text=f"Box ID: {v['id'].get()}  |  {box_name}")
        v["id"].trace_add("write", update_id_disp)

        type_frm = tk.Frame(sec_pres, bg="#1e1e2e")
        type_frm.pack(anchor="w", padx=8, pady=2)
        tk.Label(type_frm, text="Drop Type:", bg="#1e1e2e",
                 fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
        tk.Radiobutton(type_frm, text="Random",      variable=present_type_var, value=0,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=8)
        tk.Radiobutton(type_frm, text="Distributive", variable=present_type_var, value=2,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=8)
        tk.Checkbutton(type_frm, text="Remember this setting", variable=remember_present,
                       bg="#1e1e2e", fg="#fab387", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=12)

        # DropCnt (only visible for Random)
        dc_frm = tk.Frame(sec_pres, bg="#1e1e2e")
        dc_frm.pack(anchor="w", padx=8, pady=2)
        dc_lbl = tk.Label(dc_frm, text="DropCnt:", bg="#1e1e2e",
                          fg="#cdd6f4", font=("Consolas", 9))
        dc_ent = tk.Entry(dc_frm, textvariable=drop_cnt_var, width=6,
                          bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                          font=("Consolas", 9), relief="flat")
        dr_lbl = tk.Label(dc_frm, text="Default DropRate:", bg="#1e1e2e",
                          fg="#cdd6f4", font=("Consolas", 9))
        dr_ent = tk.Entry(dc_frm, textvariable=default_rate_var, width=6,
                          bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                          font=("Consolas", 9), relief="flat")

        def toggle_drop_fields(*_):
            if present_type_var.get() == 0:
                dc_lbl.pack(side="left", padx=4)
                dc_ent.pack(side="left", padx=4)
                dr_lbl.pack(side="left", padx=4)
                dr_ent.pack(side="left", padx=4)
            else:
                dc_lbl.pack_forget()
                dc_ent.pack_forget()
                dr_lbl.pack_forget()
                dr_ent.pack_forget()
        present_type_var.trace_add("write", toggle_drop_fields)
        toggle_drop_fields()

        # Per-item rate table
        sec_items = section(container, f"  Box Contents ({len(grp['items'])} items)  ")
        items_frm = tk.Frame(sec_items, bg="#1e1e2e")
        items_frm.pack(fill="x", padx=8, pady=4)

        tk.Label(items_frm, text="#",    width=4,  bg="#181825", fg="#89b4fa",
                 font=("Consolas", 9, "bold")).grid(row=0, column=0, padx=2)
        tk.Label(items_frm, text="ID",   width=10, bg="#181825", fg="#89b4fa",
                 font=("Consolas", 9, "bold")).grid(row=0, column=1, padx=2)
        tk.Label(items_frm, text="Name", width=40, bg="#181825", fg="#89b4fa",
                 font=("Consolas", 9, "bold"), anchor="w").grid(row=0, column=2, padx=2, sticky="w")
        tk.Label(items_frm, text="DropRate",width=10, bg="#181825", fg="#89b4fa",
                 font=("Consolas", 9, "bold")).grid(row=0, column=3, padx=2)
        tk.Label(items_frm, text="(hidden if Distrib)", bg="#1e1e2e", fg="#6c7086",
                 font=("Consolas", 8)).grid(row=0, column=4, padx=4)

        for i, item in enumerate(grp["items"]):
            bg = "#1e1e2e" if i % 2 == 0 else "#181825"
            tk.Label(items_frm, text=str(i), width=4, bg=bg, fg="#6c7086",
                     font=("Consolas", 9)).grid(row=i+1, column=0, padx=2, pady=1)
            tk.Label(items_frm, text=item["id"], width=10, bg=bg, fg="#cdd6f4",
                     font=("Consolas", 9)).grid(row=i+1, column=1, padx=2)
            tk.Label(items_frm, text=item["name"][:48], width=40, bg=bg, fg="#a6adc8",
                     font=("Consolas", 9), anchor="w").grid(row=i+1, column=2, padx=2, sticky="w")
            tk.Entry(items_frm, textvariable=item_rate_vars[i], width=8,
                     bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                     font=("Consolas", 9), relief="flat").grid(row=i+1, column=3, padx=2)

        # ── Navigation buttons ──
        nav = tk.Frame(container, bg="#181825")
        nav.pack(fill="x", pady=10)

        def gather_config():
            """Collect all field values into a dict."""
            try:
                tickets = float(v["ticket"].get() or "0")
                ncash   = round(tickets * 133)
            except:
                ncash = 0

            item_rates = []
            for i, item in enumerate(grp["items"]):
                try:
                    rate = int(item_rate_vars[i].get())
                except:
                    rate = 50
                item_rates.append(rate)

            cfg = {
                "id":           v["id"].get().strip(),
                "name":         v["name"].get(),
                "comment":      v["comment"].get(),
                "use":          v["use"].get(),
                "file_name":    v["file_name"].get(),
                "inv_file_name":v["inv_file_name"].get(),
                "cmt_file_name":v["cmt_file_name"].get(),
                "weight":       v["weight"].get() or "1",
                "value":        v["value"].get() or "0",
                "min_level":    v["min_level"].get() or "1",
                "money":        v["money"].get() or "0",
                "ncash":        ncash,
                "ticket":       v["ticket"].get() or "0",
                "opt_checks":   [bv.get() for bv in opt_check_vars],
                "opt_recycle":  opt_recycle_var.get(),
                "chr_type_flags": list(chr_selected),
                "present_type": present_type_var.get(),
                "drop_cnt":     drop_cnt_var.get() or "1",
                "default_rate": default_rate_var.get() or "50",
                "remember_present": remember_present.get(),
                "item_rates":   item_rates,
                "box_name":     box_name,
                "items":        grp["items"],
            }
            # Update last filepath memory
            self.last_file_name.set(cfg["file_name"])
            self.last_inv_file_name.set(cfg["inv_file_name"])
            self.last_cmt_file_name.set(cfg["cmt_file_name"])
            return cfg

        def save_settings_and_continue(cfg):
            """Persist settings for future boxes."""
            self.saved_settings = {
                "comment":      cfg["comment"],
                "use":          cfg["use"],
                "file_name":    cfg["file_name"],
                "inv_file_name":cfg["inv_file_name"],
                "cmt_file_name":cfg["cmt_file_name"],
                "weight":       cfg["weight"],
                "value":        cfg["value"],
                "min_level":    cfg["min_level"],
                "money":        cfg["money"],
                "ticket":       cfg["ticket"],
                "opt_checks":   cfg["opt_checks"],
                "opt_recycle":  cfg["opt_recycle"],
                "chr_type_flags": cfg["chr_type_flags"],
                "remember_present": cfg["remember_present"],
            }
            if cfg["remember_present"]:
                self.saved_settings["present_type"]  = cfg["present_type"]
                self.saved_settings["drop_cnt"]       = cfg["drop_cnt"]
                self.saved_settings["default_rate"]   = cfg["default_rate"]

        def go_next():
            cfg = gather_config()
            if not cfg["id"]:
                messagebox.showwarning("Missing ID", "Please enter a Box ID before continuing.")
                return
            self.box_configs.append(cfg)
            save_settings_and_continue(cfg)
            self.current_group_idx += 1
            if self.current_group_idx < len(self.groups):
                self._ask_automate_or_monitor(cfg)
            else:
                self._build_output_screen()

        tk.Button(nav, text="◀  Back to CSV", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=10, pady=8)

        if idx > 0:
            def go_prev():
                cfg = gather_config()
                self.box_configs.append(cfg)  # partial save
                self.current_group_idx -= 1
                self.box_configs = self.box_configs[:-2]  # remove this and prev
                self._build_config_screen()
            tk.Button(nav, text="◀  Previous Box", command=go_prev,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                      relief="flat", padx=12, pady=6).pack(side="left", padx=4, pady=8)

        next_lbl = "Generate XML ✓" if idx == len(self.groups) - 1 else "Next Box ▶"
        tk.Button(nav, text=next_lbl, command=go_next,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 10, "bold"),
                  relief="flat", padx=12, pady=6).pack(side="right", padx=10, pady=8)

    # ── Ask automate vs monitor ───────────────────────────────────────────
    def _ask_automate_or_monitor(self, last_cfg):
        remaining = len(self.groups) - self.current_group_idx
        win = tk.Toplevel(self)
        win.title("Continue?")
        win.geometry("480x220")
        win.configure(bg="#1e1e2e")
        win.grab_set()

        tk.Label(win, text=f"{remaining} box(es) remaining.",
                 bg="#1e1e2e", fg="#cdd6f4", font=("Consolas", 12, "bold")).pack(pady=15)
        tk.Label(win, text="How would you like to continue?",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack()

        btn_frm = tk.Frame(win, bg="#1e1e2e")
        btn_frm.pack(pady=20)

        def automate():
            win.destroy()
            self._automate_remaining(last_cfg)

        def monitor():
            win.destroy()
            self._build_config_screen()

        tk.Button(btn_frm, text="🤖  Automate (use saved settings for all remaining)",
                  command=automate, bg="#cba6f7", fg="#1e1e2e",
                  font=("Consolas", 10), relief="flat", padx=10, pady=8).pack(pady=6)
        tk.Button(btn_frm, text="👁  Monitor (prompt me for each box)",
                  command=monitor, bg="#89b4fa", fg="#1e1e2e",
                  font=("Consolas", 10), relief="flat", padx=10, pady=8).pack(pady=6)

    def _automate_remaining(self, last_cfg):
        """Auto-fill remaining boxes by incrementing IDs and swapping names."""
        try:
            base_id = int(last_cfg["id"])
        except:
            messagebox.showerror("Error", "Cannot automate: last Box ID was not numeric.")
            self._build_config_screen()
            return

        id_counter = base_id + 1
        for i in range(self.current_group_idx, len(self.groups)):
            grp = self.groups[i]
            cfg = copy.deepcopy(last_cfg)
            cfg["id"]       = str(id_counter)
            cfg["name"]     = grp["box_name"]
            cfg["box_name"] = grp["box_name"]
            cfg["items"]    = grp["items"]
            # Preserve per-item rates from CSV if available
            cfg["item_rates"] = [
                it["rate"] if it["rate"] is not None else int(last_cfg["default_rate"])
                for it in grp["items"]
            ]
            self.box_configs.append(cfg)
            id_counter += 1

        self.current_group_idx = len(self.groups)
        self._build_output_screen()

    # ── Screen 2: Output ──────────────────────────────────────────────────
    def _build_output_screen(self):
        self._clear()
        tk.Label(self, text="Generated XML Output", font=("Consolas", 14, "bold"),
                 bg="#1e1e2e", fg="#cba6f7").pack(pady=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        # Build XML strings
        itemparam_rows     = []
        presentparam_rows  = []

        for cfg in self.box_configs:
            try:
                default_rate = int(cfg.get("default_rate", 50))
            except:
                default_rate = 50
            try:
                drop_cnt = int(cfg.get("drop_cnt", 1))
            except:
                drop_cnt = 1

            # Merge per-item rates
            items_with_rates = []
            for j, it in enumerate(cfg["items"]):
                rate = cfg["item_rates"][j] if j < len(cfg["item_rates"]) else default_rate
                items_with_rates.append({**it, "rate": rate})

            itemparam_rows.append(build_itemparam_row(cfg))
            presentparam_rows.append(
                build_presentparam_row(cfg["id"], items_with_rates,
                                       cfg["present_type"], drop_cnt, default_rate)
            )

        def make_tab(title, content):
            frame = tk.Frame(nb, bg="#1e1e2e")
            nb.add(frame, text=title)
            txt = scrolledtext.ScrolledText(frame, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4",
                                            insertbackground="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", "\n\n".join(content))
            txt.config(state="disabled")

            def copy_all():
                self.clipboard_clear()
                self.clipboard_append("\n\n".join(content))
                messagebox.showinfo("Copied", f"{title} copied to clipboard.")

            def save_file():
                default = "itemparam_rows.xml" if "Item" in title else "presentparam_rows.xml"
                path = filedialog.asksaveasfilename(defaultextension=".xml",
                                                    initialfile=default,
                                                    filetypes=[("XML","*.xml"),("All","*.*")])
                if path:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write("\n\n".join(content))
                    messagebox.showinfo("Saved", f"Saved to {path}")

            btn_row = tk.Frame(frame, bg="#1e1e2e")
            btn_row.pack(fill="x")
            tk.Button(btn_row, text="📋 Copy All", command=copy_all,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
            tk.Button(btn_row, text="💾 Save As...", command=save_file,
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)

        make_tab("itemparam.xml rows",         itemparam_rows)
        make_tab("PresentItemParam2.xml rows", presentparam_rows)

        # Also export a simple ID/name CSV for Script 2/3 compatibility
        csv_lines = ["ID,BoxName"]
        for cfg in self.box_configs:
            csv_lines.append(f"{cfg['id']},{cfg['box_name']}")
        csv_content = ["\n".join(csv_lines)]

        make_tab("Box ID List (for Script 2)", csv_content)

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    # ── Helpers ───────────────────────────────────────────────────────────
    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


if __name__ == "__main__":
    app = BoxGeneratorApp()
    app.mainloop()
