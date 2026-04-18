"""
SCRIPT 1 - Box XML Generator  (v4)
Reads a CSV (groups of 3 cols: ID, Level/Rate, ParentBoxName), lets you configure
each box's ItemParam and PresentItemParam2 settings, then exports XML rows.

Requirements: Python 3.x (tkinter is included in standard Python installs)
Run: python script1_box_generator.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, copy

# ─── Character data ───────────────────────────────────────────────────────────
# Unique character names (for the left combo)
CHR_NAMES = ["Bunny", "Buffalo", "Sheep", "Dragon", "Fox", "Lion",
             "Cat", "Raccoon", "Paula"]
# Job tiers
CHR_JOBS  = ["1st", "2nd", "3rd"]
# Full flag table  name -> value
CHR_FLAG_MAP = {
    "Bunny 1st":   1,        "Buffalo 1st": 2,      "Sheep 1st":   4,
    "Dragon 1st":  8,        "Fox 1st":     16,     "Lion 1st":    32,
    "Cat 1st":     64,       "Raccoon 1st": 124,    "Paula 1st":   256,
    "Bunny 2nd":   512,      "Buffalo 2nd": 1024,   "Sheep 2nd":   2048,
    "Dragon 2nd":  4096,     "Fox 2nd":     8192,   "Lion 2nd":    16384,
    "Cat 2nd":     32768,    "Raccoon 2nd": 65536,  "Paula 2nd":   131072,
    "Bunny 3rd":   262144,   "Buffalo 3rd": 524288, "Sheep 3rd":   1048576,
    "Dragon 3rd":  2097152,  "Fox 3rd":     4194304,"Lion 3rd":    8388608,
    "Cat 3rd":     16777216, "Raccoon 3rd": 33554432,"Paula 3rd":  67108864,
}
CHR_FLAG_REVERSE = {v: k for k, v in CHR_FLAG_MAP.items()}

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

# ─── Text template helpers ─────────────────────────────────────────────────────
def substitute_box_name(template_text, old_box_name, new_box_name):
    """
    Replace occurrences of old_box_name inside a text template with new_box_name.
    Handles partial matches so only the character/box name portion is swapped.
    """
    if not old_box_name or not template_text:
        return template_text
    # Escape for regex
    escaped = re.escape(old_box_name)
    result = re.sub(escaped, new_box_name, template_text, flags=re.IGNORECASE)
    return result

def apply_name_template(template, prev_box_name, new_box_name):
    """
    Given a saved template (e.g. 'Dragon Special Gear Box'),
    replace prev_box_name with new_box_name.
    Returns the substituted string.
    """
    if not template or not prev_box_name:
        return new_box_name
    return substitute_box_name(template, prev_box_name, new_box_name)

def deduplicate_name(proposed, existing_names):
    """
    If proposed is already in existing_names, append (2), (3), etc.
    """
    if proposed not in existing_names:
        return proposed
    i = 2
    while f"{proposed} ({i})" in existing_names:
        i += 1
    return f"{proposed} ({i})"

# ─── CSV Parser ───────────────────────────────────────────────────────────────
def parse_csv_text(text):
    reader = csv.reader(io.StringIO(text))
    rows   = list(reader)
    if not rows:
        return []
    headers = rows[0]
    groups  = []
    i = 0
    while i < len(headers):
        box_name = headers[i + 2].strip() if i + 2 < len(headers) else ""
        if box_name:
            items = []
            for row in rows[1:]:
                id_val = row[i].strip()   if i     < len(row) else ""
                lv_val = row[i+1].strip() if i + 1 < len(row) else ""
                if id_val and id_val.isdigit():
                    rate_clean = "".join(c for c in lv_val if c.isdigit())
                    items.append({
                        "id":    id_val,
                        "level": lv_val,
                        "rate":  int(rate_clean) if rate_clean else None,
                        "name":  row[i+2].strip() if i + 2 < len(row) else "",
                    })
            groups.append({"box_name": box_name, "items": items})
        i += 3
    return groups

# ─── XML Builders ─────────────────────────────────────────────────────────────
def build_options_str(check_states, recycle_val):
    base    = [2, 32]
    checked = [v for (_, v), on in zip(OPTIONS_CHECKS, check_states) if on]
    rec     = [recycle_val] if recycle_val > 0 else []
    return "/".join(str(x) for x in sorted(set(base + checked + rec)))

def build_itemparam_row(cfg):
    opts      = build_options_str(cfg["opt_checks"], cfg["opt_recycle"])
    chr_flags = sum(cfg["chr_type_flags"])
    fn        = cfg["file_name"]
    bn        = cfg["bundle_num"]
    return f"""<ROW>
<ID>{cfg['id']}</ID>
<Class>1</Class>
<Type>15</Type>
<SubType>0</SubType>
<ItemFType>0</ItemFType>
<n><![CDATA[{cfg['name']}]]></n>
<Comment><![CDATA[{cfg['comment']}]]></Comment>
<Use><![CDATA[{cfg['use']}]]></Use>
<Name_Eng><![CDATA[ ]]></Name_Eng>
<Comment_Eng><![CDATA[ ]]></Comment_Eng>
<FileName><![CDATA[{fn}]]></FileName>
<BundleNum>{bn}</BundleNum>
<InvFileName><![CDATA[{fn}]]></InvFileName>
<InvBundleNum>{bn}</InvBundleNum>
<CmtFileName><![CDATA[{cfg['cmt_file_name']}]]></CmtFileName>
<CmtBundleNum>{cfg['cmt_bundle_num']}</CmtBundleNum>
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
<Ncash>{cfg['ncash']}</Ncash>
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
    is_distrib      = (ptype == 2)
    actual_drop_cnt = len(items) if is_distrib else drop_cnt
    lines = ["<ROW>", f"<Id>{box_id}</Id>",
             f"<Type>{ptype}</Type>", f"<DropCnt>{actual_drop_cnt}</DropCnt>"]
    for i in range(20):
        if i < len(items):
            did   = items[i]["id"]
            drate = 100 if is_distrib else (items[i].get("rate") or default_rate)
        else:
            did, drate = 0, 0
        lines += [f"<DropId_{i}>{did}</DropId_{i}>",
                  f"<DropRate_{i}>{drate}</DropRate_{i}>",
                  f"<ItemCnt_{i}>0</ItemCnt_{i}>"]
    lines.append("</ROW>")
    return "\n".join(lines)

# ─── App ──────────────────────────────────────────────────────────────────────
class BoxGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 1 – Box XML Generator")
        self.geometry("1000x820")
        self.configure(bg="#1e1e2e")

        self.saved_settings    = None   # persisted field values
        self.continue_mode     = None   # "automate" | "monitor"
        self.groups            = []
        self.current_group_idx = 0
        self.box_configs       = []     # completed cfgs

        self._build_load_screen()

    # ══════════════════════════════════════════════════════════════════════
    # SCREEN 0 – Load CSV
    # ══════════════════════════════════════════════════════════════════════
    def _build_load_screen(self):
        self._clear()
        frm = tk.Frame(self, bg="#1e1e2e")
        frm.pack(expand=True)
        tk.Label(frm, text="BOX XML GENERATOR", font=("Consolas", 20, "bold"),
                 bg="#1e1e2e", fg="#cba6f7").pack(pady=(30, 5))
        tk.Label(frm,
                 text="Load a CSV with groups of 3 columns:  ID  |  Level/Rate  |  Parent Box Name",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack(pady=5)
        bf = tk.Frame(frm, bg="#1e1e2e")
        bf.pack(pady=15)
        tk.Button(bf, text="📂  Load CSV File", command=self._load_csv_file,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 11),
                  relief="flat", padx=16, pady=8).pack(side="left", padx=8)
        tk.Button(bf, text="📋  Paste CSV Text", command=self._paste_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 11),
                  relief="flat", padx=16, pady=8).pack(side="left", padx=8)

    def _load_csv_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
        if path:
            with open(path, encoding="utf-8-sig") as f:
                self._process_csv(f.read())

    def _paste_csv(self):
        win = tk.Toplevel(self); win.title("Paste CSV")
        win.geometry("600x400"); win.configure(bg="#1e1e2e")
        tk.Label(win, text="Paste CSV below:", bg="#1e1e2e",
                 fg="#cdd6f4", font=("Consolas", 10)).pack(anchor="w", padx=10, pady=5)
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9))
        txt.pack(fill="both", expand=True, padx=10, pady=5)
        def confirm():
            self._process_csv(txt.get("1.0","end")); win.destroy()
        tk.Button(win, text="Confirm", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _process_csv(self, text):
        groups = parse_csv_text(text)
        if not groups:
            messagebox.showerror("Error", "No valid box groups found in CSV.")
            return
        self.groups            = groups
        self.current_group_idx = 0
        self.box_configs       = []
        self.continue_mode     = None
        self.saved_settings    = None
        self._build_config_screen()

    # ══════════════════════════════════════════════════════════════════════
    # SCREEN 1 – Configure one box
    # ══════════════════════════════════════════════════════════════════════
    def _build_config_screen(self):
        self._clear()
        idx      = self.current_group_idx
        grp      = self.groups[idx]
        box_name = grp["box_name"]
        total    = len(self.groups)
        s        = self.saved_settings or {}

        # Previous box name used for template substitution
        prev_box_name = s.get("box_name", "")

        # ── Scroll canvas ────────────────────────────────────────────────
        outer = tk.Frame(self, bg="#1e1e2e")
        outer.pack(fill="both", expand=True)

        hdr = tk.Frame(outer, bg="#181825")
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Box {idx+1} / {total}:  {box_name}",
                 font=("Consolas", 14, "bold"), bg="#181825", fg="#cba6f7",
                 pady=8).pack(side="left", padx=15)
        if self.continue_mode:
            mode_txt = "🤖 AUTO" if self.continue_mode == "automate" else "👁 MONITOR"
            tk.Label(hdr, text=mode_txt, font=("Consolas", 10),
                     bg="#181825", fg="#fab387").pack(side="right", padx=15)

        canvas    = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        container = tk.Frame(canvas, bg="#1e1e2e")
        win_id    = canvas.create_window((0, 0), window=container, anchor="nw")
        container.bind("<Configure>",
                       lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        # ── UI helpers ────────────────────────────────────────────────────
        def section(title):
            f = tk.LabelFrame(container, text=title, bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=12, pady=5)
            return f

        def row_entry(parent, label, var, width=38):
            r = tk.Frame(parent, bg="#1e1e2e"); r.pack(fill="x", padx=8, pady=2)
            tk.Label(r, text=label, width=26, anchor="w", bg="#1e1e2e",
                     fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
            tk.Entry(r, textvariable=var, width=width,
                     bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                     font=("Consolas", 9), relief="flat").pack(side="left", padx=4)

        def row_num(parent, label, var, width=10):
            r = tk.Frame(parent, bg="#1e1e2e"); r.pack(fill="x", padx=8, pady=2)
            tk.Label(r, text=label, width=26, anchor="w", bg="#1e1e2e",
                     fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
            tk.Entry(r, textvariable=var, width=width,
                     bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                     font=("Consolas", 9), relief="flat").pack(side="left", padx=4)

        def note(parent, text):
            tk.Label(parent, text=text, bg="#1e1e2e", fg="#6c7086",
                     font=("Consolas", 8)).pack(anchor="w", padx=10, pady=(3,0))

        # ── Compute initial values with template substitution ─────────────

        # ID: last saved ID + 1
        try:
            next_id = str(int(s.get("id", "")) + 1)
        except (ValueError, TypeError):
            next_id = s.get("id", "")

        # Name: apply template substitution (prev_box_name → box_name)
        saved_name_template = s.get("name_template", "")
        if saved_name_template and prev_box_name:
            proposed_name = apply_name_template(saved_name_template, prev_box_name, box_name)
        else:
            proposed_name = box_name
        # Deduplicate against already-used names
        used_names = [c["name"] for c in self.box_configs]
        proposed_name = deduplicate_name(proposed_name, used_names)

        # Comment template
        saved_comment_template = s.get("comment_template", "")
        if saved_comment_template and prev_box_name:
            initial_comment = apply_name_template(saved_comment_template, prev_box_name, box_name)
        else:
            initial_comment = s.get("comment", "Special dice that contains amazing items.")

        # Use template
        saved_use_template = s.get("use_template", "")
        if saved_use_template and prev_box_name:
            initial_use = apply_name_template(saved_use_template, prev_box_name, box_name)
        else:
            initial_use = s.get("use", "Event Box.")

        # ── Variables ─────────────────────────────────────────────────────
        v_id       = tk.StringVar(value=next_id)
        v_name     = tk.StringVar(value=proposed_name)
        v_comment  = tk.StringVar(value=initial_comment)
        v_use      = tk.StringVar(value=initial_use)

        v_file_name     = tk.StringVar(value=s.get("file_name",      r"data\item\itm_pre_107.nri"))
        v_bundle_num    = tk.StringVar(value=s.get("bundle_num",     "0"))
        v_cmt_file_name = tk.StringVar(value=s.get("cmt_file_name",  r"data\item\itm_pre_illu_107.nri"))
        v_cmt_bundle    = tk.StringVar(value=s.get("cmt_bundle_num", "0"))

        opt_check_vars  = [tk.BooleanVar(value=s.get("opt_checks", [False]*8)[i])
                           for i in range(len(OPTIONS_CHECKS))]
        opt_recycle_var = tk.IntVar(value=s.get("opt_recycle", 0))

        chr_selected = list(s.get("chr_type_flags", []))

        present_type_var = tk.IntVar(value=s.get("present_type", 0))
        drop_cnt_var     = tk.StringVar(value=s.get("drop_cnt",     "1"))
        default_rate_var = tk.StringVar(value=s.get("default_rate", "50"))
        remember_present = tk.BooleanVar(value=s.get("remember_present", False))

        v_weight    = tk.StringVar(value=s.get("weight",    "1"))
        v_value     = tk.StringVar(value=s.get("value",     "0"))
        v_min_level = tk.StringVar(value=s.get("min_level", "1"))
        v_money     = tk.StringVar(value=s.get("money",     "0"))
        v_ticket    = tk.StringVar(value=s.get("ticket",    "0"))

        # Per-item rate vars
        item_rate_vars = []
        for it in grp["items"]:
            if present_type_var.get() == 2:
                initial = "100"
            elif it["rate"] is not None:
                initial = str(it["rate"])
            else:
                try:
                    initial = str(int(s.get("default_rate", "50")))
                except:
                    initial = "50"
            item_rate_vars.append(tk.StringVar(value=initial))

        # ═════════════════════════════════════════════════════════════════
        # UI SECTIONS
        # ═════════════════════════════════════════════════════════════════

        # ── Basic Info ───────────────────────────────────────────────────
        sec_basic = section("  ItemParam – Basic Info  ")
        row_entry(sec_basic, "Box ID (itemparam):", v_id, 20)
        note(sec_basic, "  Name auto-filled from CSV header (with template applied) — editable.")
        row_entry(sec_basic, "Name (CDATA):",    v_name)
        row_entry(sec_basic, "Comment (CDATA):", v_comment)
        row_entry(sec_basic, "Use (CDATA):",     v_use)

        # ── Filepaths ────────────────────────────────────────────────────
        sec_files = section("  Filepaths & Bundle Numbers  ")
        note(sec_files, "  FileName — also written to InvFileName automatically (identical).")
        row_entry(sec_files, "FileName (CDATA):", v_file_name)
        note(sec_files, "  BundleNum — animation frame index.  Also written to InvBundleNum.")
        row_num(sec_files, "BundleNum:", v_bundle_num, 8)
        note(sec_files, "  CmtFileName — comment/illustration portrait file.")
        row_entry(sec_files, "CmtFileName (CDATA):", v_cmt_file_name)
        note(sec_files, "  CmtBundleNum — animation frame index for comment portrait.")
        row_num(sec_files, "CmtBundleNum:", v_cmt_bundle, 8)

        # ── Options checkboxes ────────────────────────────────────────────
        sec_opts = section("  Options  ")
        chk_frm = tk.Frame(sec_opts, bg="#1e1e2e")
        chk_frm.pack(anchor="w", padx=8, pady=4)
        for i, (lbl, _) in enumerate(OPTIONS_CHECKS):
            tk.Checkbutton(chk_frm, text=lbl, variable=opt_check_vars[i],
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e", font=("Consolas", 9)
                           ).grid(row=i//4, column=i%4, sticky="w", padx=6, pady=2)

        # Recycle radio
        rec_frm = tk.Frame(sec_opts, bg="#1e1e2e")
        rec_frm.pack(anchor="w", padx=8, pady=4)
        tk.Label(rec_frm, text="Recycle:", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Consolas", 9)).pack(side="left", padx=(0,8))
        for lbl, val in [("None",0),("Recyclable",262144),("Non-Recyclable",8388608)]:
            tk.Radiobutton(rec_frm, text=lbl, variable=opt_recycle_var, value=val,
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e", font=("Consolas", 9)
                           ).pack(side="left", padx=6)

        # ── NCash / Ticket — shown only when Recyclable ───────────────────
        # (placed inside sec_opts, AFTER the recycle radios, NOT in its own section)
        nc_container = tk.Frame(sec_opts, bg="#1e1e2e")
        nc_container.pack(anchor="w", padx=8, pady=2)

        nc_inner = tk.Frame(nc_container, bg="#1e1e2e")
        tk.Label(nc_inner, text="Ticket value (NCash = tickets × 133, rounded):",
                 bg="#1e1e2e", fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
        tk.Entry(nc_inner, textvariable=v_ticket, width=10,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 font=("Consolas", 9), relief="flat").pack(side="left", padx=8)
        ncash_lbl = tk.Label(nc_inner, text="→ NCash: 0", bg="#1e1e2e",
                             fg="#a6e3a1", font=("Consolas", 9))
        ncash_lbl.pack(side="left")

        def update_ncash(*_):
            try:
                ncash_lbl.config(text=f"→ NCash: {round(float(v_ticket.get())*133)}")
            except:
                ncash_lbl.config(text="→ NCash: ?")
        v_ticket.trace_add("write", update_ncash)
        update_ncash()

        def toggle_ncash(*_):
            if opt_recycle_var.get() == 262144:  # Recyclable only
                nc_inner.pack(anchor="w")
            else:
                nc_inner.pack_forget()
        opt_recycle_var.trace_add("write", toggle_ncash)
        toggle_ncash()

        # ── ChrTypeFlags — Character / Job picker ─────────────────────────
        sec_chr = section("  ChrTypeFlags (up to 24 character flags)  ")

        # Top row: picker   "Character Type: [combo]   Job: [combo]   [+] [-]"
        picker_frm = tk.Frame(sec_chr, bg="#1e1e2e")
        picker_frm.pack(fill="x", padx=8, pady=4)

        tk.Label(picker_frm, text="Character Type:", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Consolas", 9)).pack(side="left")
        chr_name_combo = ttk.Combobox(picker_frm, values=CHR_NAMES,
                                      state="readonly", width=14, font=("Consolas", 9))
        chr_name_combo.pack(side="left", padx=(6, 12))

        tk.Label(picker_frm, text="Job:", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Consolas", 9)).pack(side="left")
        chr_job_combo = ttk.Combobox(picker_frm, values=CHR_JOBS,
                                     state="readonly", width=6, font=("Consolas", 9))
        chr_job_combo.pack(side="left", padx=(6, 12))

        def add_chr_flag():
            name = chr_name_combo.get()
            job  = chr_job_combo.get()
            if not name or not job:
                return
            key = f"{name} {job}"
            val = CHR_FLAG_MAP.get(key)
            if val and val not in chr_selected and len(chr_selected) < 24:
                chr_selected.append(val)
                refresh_chr_lb()

        def rem_chr_flag():
            sel = chr_lb.curselection()
            if sel:
                chr_selected.pop(sel[0])
                refresh_chr_lb()

        tk.Button(picker_frm, text="+", command=add_chr_flag,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", width=3).pack(side="left", padx=2)
        tk.Button(picker_frm, text="−", command=rem_chr_flag,
                  bg="#f38ba8", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", width=3).pack(side="left", padx=2)

        # Listbox of currently added flags
        lb_frm = tk.Frame(sec_chr, bg="#1e1e2e")
        lb_frm.pack(fill="x", padx=8, pady=(0, 6))
        tk.Label(lb_frm, text="Added:", bg="#1e1e2e", fg="#6c7086",
                 font=("Consolas", 8)).pack(anchor="w")
        chr_lb = tk.Listbox(lb_frm, height=4, width=36, bg="#313244", fg="#cdd6f4",
                            font=("Consolas", 9), selectbackground="#45475a",
                            activestyle="none")
        chr_lb.pack(anchor="w")

        def refresh_chr_lb():
            chr_lb.delete(0, "end")
            for val in chr_selected:
                chr_lb.insert("end", CHR_FLAG_REVERSE.get(val, str(val)))

        refresh_chr_lb()

        # ── Numeric Fields ────────────────────────────────────────────────
        sec_nums = section("  Numeric Fields  ")
        rf = tk.Frame(sec_nums, bg="#1e1e2e")
        rf.pack(fill="x", padx=8, pady=4)
        for ci, (lbl, var) in enumerate([("Weight:", v_weight), ("Value:", v_value),
                                          ("MinLevel:", v_min_level), ("Money:", v_money)]):
            tk.Label(rf, text=lbl, bg="#1e1e2e", fg="#cdd6f4",
                     font=("Consolas", 9), width=10, anchor="w").grid(row=0, column=ci*2, padx=4)
            tk.Entry(rf, textvariable=var, width=10, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas", 9),
                     relief="flat").grid(row=0, column=ci*2+1, padx=4)

        # ── PresentItemParam2 ─────────────────────────────────────────────
        sec_pres = section("  PresentItemParam2 Settings  ")

        id_disp = tk.Label(sec_pres, text=f"Box ID: {next_id}  |  {box_name}",
                           bg="#1e1e2e", fg="#6c7086", font=("Consolas", 9))
        id_disp.pack(anchor="w", padx=8, pady=(4,0))
        v_id.trace_add("write",
            lambda *_: id_disp.config(text=f"Box ID: {v_id.get() or '—'}  |  {box_name}"))

        type_frm = tk.Frame(sec_pres, bg="#1e1e2e")
        type_frm.pack(anchor="w", padx=8, pady=4)
        tk.Label(type_frm, text="Drop Type:", bg="#1e1e2e",
                 fg="#cdd6f4", font=("Consolas", 9)).pack(side="left")
        tk.Radiobutton(type_frm, text="Random",       variable=present_type_var, value=0,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=8)
        tk.Radiobutton(type_frm, text="Distributive", variable=present_type_var, value=2,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=8)
        tk.Checkbutton(type_frm, text="Remember this setting", variable=remember_present,
                       bg="#1e1e2e", fg="#fab387", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 9)).pack(side="left", padx=14)

        dc_frm = tk.Frame(sec_pres, bg="#1e1e2e")
        dc_frm.pack(anchor="w", padx=8, pady=2)
        dc_lbl = tk.Label(dc_frm, text="DropCnt:", bg="#1e1e2e", fg="#cdd6f4", font=("Consolas",9))
        dc_ent = tk.Entry(dc_frm, textvariable=drop_cnt_var, width=6,
                          bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                          font=("Consolas",9), relief="flat")
        dr_lbl = tk.Label(dc_frm, text="Default DropRate:", bg="#1e1e2e", fg="#cdd6f4", font=("Consolas",9))
        dr_ent = tk.Entry(dc_frm, textvariable=default_rate_var, width=6,
                          bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                          font=("Consolas",9), relief="flat")
        rate_note = tk.Label(sec_pres, text="", bg="#1e1e2e", fg="#6c7086", font=("Consolas",8))
        rate_note.pack(anchor="w", padx=8)

        # ── Box Contents table ────────────────────────────────────────────
        sec_items = section(f"  Box Contents ({len(grp['items'])} items)  ")
        items_frm = tk.Frame(sec_items, bg="#1e1e2e")
        items_frm.pack(fill="x", padx=8, pady=4)

        for ci, (txt, w) in enumerate([("#",4),("ID",10),("Name",44),("DropRate",10)]):
            tk.Label(items_frm, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas",9,"bold"), anchor="w"
                     ).grid(row=0, column=ci, padx=2, pady=2, sticky="w")

        rate_entry_widgets = []
        for i, item in enumerate(grp["items"]):
            bg = "#1e1e2e" if i % 2 == 0 else "#181825"
            tk.Label(items_frm, text=str(i),       width=4,  bg=bg, fg="#6c7086",
                     font=("Consolas",9)).grid(row=i+1, column=0, padx=2, pady=1)
            tk.Label(items_frm, text=item["id"],    width=10, bg=bg, fg="#cdd6f4",
                     font=("Consolas",9)).grid(row=i+1, column=1, padx=2)
            tk.Label(items_frm, text=item["name"][:52], width=44, bg=bg, fg="#a6adc8",
                     font=("Consolas",9), anchor="w"
                     ).grid(row=i+1, column=2, padx=2, sticky="w")
            ent = tk.Entry(items_frm, textvariable=item_rate_vars[i], width=8,
                           bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                           font=("Consolas",9), relief="flat")
            ent.grid(row=i+1, column=3, padx=2)
            rate_entry_widgets.append(ent)

        # Toggle drop fields + update rate entries when present_type changes
        def toggle_drop_fields(*_):
            is_distrib = (present_type_var.get() == 2)
            if is_distrib:
                dc_lbl.pack_forget(); dc_ent.pack_forget()
                dr_lbl.pack_forget(); dr_ent.pack_forget()
                rate_note.config(text="  Distributive: all DropRates are 100.")
                # Visually set all rate entries to 100 and make them read-only
                for i, var in enumerate(item_rate_vars):
                    var.set("100")
                for ent in rate_entry_widgets:
                    ent.config(state="disabled", disabledforeground="#6c7086")
            else:
                dc_lbl.pack(side="left", padx=4); dc_ent.pack(side="left", padx=4)
                dr_lbl.pack(side="left", padx=4); dr_ent.pack(side="left", padx=4)
                rate_note.config(text="  DropRate is editable per item below.")
                # Restore rates from CSV or default
                for i, (var, item) in enumerate(zip(item_rate_vars, grp["items"])):
                    if item["rate"] is not None:
                        var.set(str(item["rate"]))
                    else:
                        try:
                            var.set(str(int(default_rate_var.get())))
                        except:
                            var.set("50")
                for ent in rate_entry_widgets:
                    ent.config(state="normal")

        present_type_var.trace_add("write", toggle_drop_fields)
        toggle_drop_fields()   # apply on first render

        # ═════════════════════════════════════════════════════════════════
        # Gather / Save / Navigate
        # ═════════════════════════════════════════════════════════════════
        def gather_config():
            try:
                ncash = round(float(v_ticket.get() or "0") * 133)
            except:
                ncash = 0
            item_rates = []
            for iv in item_rate_vars:
                try:   item_rates.append(int(iv.get()))
                except: item_rates.append(100 if present_type_var.get()==2 else 50)
            return {
                "id":               v_id.get().strip(),
                "name":             v_name.get(),
                "name_template":    v_name.get(),       # save as template
                "comment":          v_comment.get(),
                "comment_template": v_comment.get(),    # save as template
                "use":              v_use.get(),
                "use_template":     v_use.get(),        # save as template
                "box_name":         box_name,           # current box name for next substitution
                "items":            grp["items"],
                "item_rates":       item_rates,
                "file_name":        v_file_name.get(),
                "bundle_num":       v_bundle_num.get() or "0",
                "cmt_file_name":    v_cmt_file_name.get(),
                "cmt_bundle_num":   v_cmt_bundle.get() or "0",
                "weight":           v_weight.get() or "1",
                "value":            v_value.get() or "0",
                "min_level":        v_min_level.get() or "1",
                "money":            v_money.get() or "0",
                "ncash":            ncash,
                "ticket":           v_ticket.get() or "0",
                "opt_checks":       [bv.get() for bv in opt_check_vars],
                "opt_recycle":      opt_recycle_var.get(),
                "chr_type_flags":   list(chr_selected),
                "present_type":     present_type_var.get(),
                "drop_cnt":         drop_cnt_var.get() or "1",
                "default_rate":     default_rate_var.get() or "50",
                "remember_present": remember_present.get(),
            }

        def save_settings(cfg):
            """Persist everything except per-box-unique fields (items list)."""
            self.saved_settings = {k: v for k, v in cfg.items() if k != "items"}

        # Nav bar
        nav = tk.Frame(container, bg="#181825")
        nav.pack(fill="x", pady=10)

        def go_next():
            cfg = gather_config()
            if not cfg["id"]:
                messagebox.showwarning("Missing ID", "Please enter a Box ID.")
                return
            self.box_configs.append(cfg)
            save_settings(cfg)
            self.current_group_idx += 1

            if self.current_group_idx >= len(self.groups):
                self._build_output_screen()
            elif self.continue_mode == "automate":
                self._automate_remaining(cfg)
            elif self.continue_mode == "monitor":
                self._build_config_screen()
            else:
                self._ask_automate_or_monitor(cfg)

        tk.Button(nav, text="◀  Back to CSV", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=10, pady=8)

        if idx > 0:
            def go_prev():
                self.current_group_idx -= 1
                if self.box_configs: self.box_configs.pop()
                self._build_config_screen()
            tk.Button(nav, text="◀  Previous Box", command=go_prev,
                      bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                      relief="flat", padx=12, pady=6).pack(side="left", padx=4, pady=8)

        if self.continue_mode:
            def change_mode():
                self.continue_mode = None
                self._build_config_screen()
            tk.Button(nav, text="⚙ Change Mode", command=change_mode,
                      bg="#45475a", fg="#cdd6f4", font=("Consolas",9),
                      relief="flat", padx=8, pady=6).pack(side="left", padx=4, pady=8)

        next_lbl = "Generate XML ✓" if idx == total-1 else "Next Box ▶"
        tk.Button(nav, text=next_lbl, command=go_next,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas",10,"bold"),
                  relief="flat", padx=12, pady=6).pack(side="right", padx=10, pady=8)

    # ══════════════════════════════════════════════════════════════════════
    # Automate vs Monitor dialog
    # ══════════════════════════════════════════════════════════════════════
    def _ask_automate_or_monitor(self, last_cfg):
        remaining = len(self.groups) - self.current_group_idx
        win = tk.Toplevel(self)
        win.title("Continue?"); win.geometry("530x260")
        win.configure(bg="#1e1e2e"); win.grab_set()

        tk.Label(win, text=f"{remaining} box(es) remaining.",
                 bg="#1e1e2e", fg="#cdd6f4", font=("Consolas",13,"bold")).pack(pady=14)
        tk.Label(win, text="How would you like to continue?",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas",10)).pack()

        remember_var = tk.BooleanVar(value=False)
        bf = tk.Frame(win, bg="#1e1e2e"); bf.pack(pady=10)

        def choose(mode):
            if remember_var.get():
                self.continue_mode = mode
            win.destroy()
            if mode == "automate":
                self._automate_remaining(last_cfg)
            else:
                self._build_config_screen()

        tk.Button(bf, text="🤖  Automate  —  use saved settings for all remaining boxes",
                  command=lambda: choose("automate"),
                  bg="#cba6f7", fg="#1e1e2e", font=("Consolas",10),
                  relief="flat", padx=10, pady=8).pack(pady=5)
        tk.Button(bf, text="👁  Monitor  —  review each box individually",
                  command=lambda: choose("monitor"),
                  bg="#89b4fa", fg="#1e1e2e", font=("Consolas",10),
                  relief="flat", padx=10, pady=8).pack(pady=5)
        tk.Checkbutton(win, text="Remember my choice for the rest of this session",
                       variable=remember_var, bg="#1e1e2e", fg="#fab387",
                       selectcolor="#313244", activebackground="#1e1e2e",
                       font=("Consolas",9)).pack(pady=4)

    # ══════════════════════════════════════════════════════════════════════
    # Automate remaining
    # ══════════════════════════════════════════════════════════════════════
    def _automate_remaining(self, last_cfg):
        try:
            base_id = int(last_cfg["id"])
        except:
            messagebox.showerror("Error",
                "Cannot automate: last Box ID was not a plain integer.\n"
                "Switching to Monitor mode.")
            self.continue_mode = "monitor"
            self._build_config_screen()
            return

        try:
            default_rate = int(last_cfg.get("default_rate", 50))
        except:
            default_rate = 50

        id_counter = base_id + 1
        prev_name  = last_cfg["box_name"]
        used_names = [c["name"] for c in self.box_configs]

        for i in range(self.current_group_idx, len(self.groups)):
            grp = self.groups[i]
            cfg = copy.deepcopy(last_cfg)

            # ID increments
            cfg["id"]       = str(id_counter)

            # Name template substitution + dedup
            proposed = apply_name_template(cfg.get("name_template",""), prev_name, grp["box_name"])
            proposed = deduplicate_name(proposed, used_names)
            cfg["name"]           = proposed
            cfg["name_template"]  = proposed

            # Comment + Use template substitution
            cfg["comment"] = apply_name_template(cfg.get("comment_template",""), prev_name, grp["box_name"])
            cfg["use"]     = apply_name_template(cfg.get("use_template",""),     prev_name, grp["box_name"])
            cfg["comment_template"] = cfg["comment"]
            cfg["use_template"]     = cfg["use"]

            cfg["box_name"] = grp["box_name"]
            cfg["items"]    = grp["items"]

            is_distrib = cfg["present_type"] == 2
            cfg["item_rates"] = [
                100 if is_distrib else (it["rate"] if it["rate"] is not None else default_rate)
                for it in grp["items"]
            ]

            used_names.append(proposed)
            prev_name = grp["box_name"]
            self.box_configs.append(cfg)
            id_counter += 1

        self.current_group_idx = len(self.groups)
        self._build_output_screen()

    # ══════════════════════════════════════════════════════════════════════
    # Output screen
    # ══════════════════════════════════════════════════════════════════════
    def _build_output_screen(self):
        self._clear()
        tk.Label(self, text="Generated XML Output", font=("Consolas",14,"bold"),
                 bg="#1e1e2e", fg="#cba6f7").pack(pady=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        itemparam_rows    = []
        presentparam_rows = []

        for cfg in self.box_configs:
            try:   default_rate = int(cfg.get("default_rate", 50))
            except: default_rate = 50
            try:   drop_cnt = int(cfg.get("drop_cnt", 1))
            except: drop_cnt = 1

            is_distrib = (cfg["present_type"] == 2)
            items_with_rates = []
            for j, it in enumerate(cfg["items"]):
                rate = 100 if is_distrib else (cfg["item_rates"][j]
                       if j < len(cfg["item_rates"]) else default_rate)
                items_with_rates.append({**it, "rate": rate})

            itemparam_rows.append(build_itemparam_row(cfg))
            presentparam_rows.append(
                build_presentparam_row(cfg["id"], items_with_rates,
                                       cfg["present_type"], drop_cnt, default_rate))

        # ── closure-safe tab builder: content is bound at call time via default arg ──
        def make_tab(title, lines, fname, _content=None):
            # _content default arg captures the value NOW, not by reference
            _content = "\n\n".join(lines)
            frm = tk.Frame(nb, bg="#1e1e2e")
            nb.add(frm, text=title)

            # Button row FIRST so it's always visible above the text
            br = tk.Frame(frm, bg="#1e1e2e")
            br.pack(fill="x", side="bottom")

            tk.Label(br, text=f"  {title}", bg="#1e1e2e", fg="#89b4fa",
                     font=("Consolas", 9, "bold")).pack(side="left", padx=6, pady=4)

            # Both buttons receive their own copy of _content via default arg
            tk.Button(br, text="📋 Copy All",
                      command=lambda c=_content: (
                          self.clipboard_clear(),
                          self.clipboard_append(c),
                          messagebox.showinfo("Copied", f"{title} copied to clipboard.")
                      ),
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="right", padx=4, pady=4)

            tk.Button(br, text="💾 Save As…",
                      command=lambda c=_content, f=fname: _save(c, f),
                      bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="right", padx=4, pady=4)

            txt = scrolledtext.ScrolledText(frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4",
                                            insertbackground="#cdd6f4")
            txt.pack(fill="both", expand=True, padx=4, pady=4)
            txt.insert("1.0", _content)
            txt.config(state="disabled")

        def _save(content, fname):
            path = filedialog.asksaveasfilename(
                defaultextension=".xml", initialfile=fname,
                filetypes=[("XML","*.xml"),("CSV","*.csv"),("All","*.*")])
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                messagebox.showinfo("Saved", f"Saved to {path}")

        make_tab("itemparam.xml rows",         itemparam_rows,    "itemparam_rows.xml")
        make_tab("PresentItemParam2.xml rows", presentparam_rows, "presentparam_rows.xml")
        csv_lines = ["ID,BoxName"] + [f"{c['id']},{c['box_name']}" for c in self.box_configs]
        make_tab("Box ID List (for Script 2)", ["\n".join(csv_lines)], "box_id_list.csv")

        # Make the PresentItemParam2 tab visually obvious — select it by default
        nb.select(1)

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas",10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


if __name__ == "__main__":
    app = BoxGeneratorApp()
    app.mainloop()
