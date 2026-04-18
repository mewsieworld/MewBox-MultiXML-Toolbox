"""
SCRIPT 3 – NCash / Ticket Cost Updater  (v3)
─────────────────────────────────────────────
Modes:
  Uniform — one Ticket Cost applied to every item in the CSV.
  Manual  — enter cost per item individually; warns if any left blank.

ItemParam loader: pick any one of the 4 files; sibling files are auto-loaded.
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

def build_item_lib(xml_text):
    """Returns {id_str: name_str} from all ROW blocks in the combined XML text."""
    lib = {}
    for row in ROW_RE.findall(xml_text):
        rid  = _get_tag(row, "ID")
        name = _get_tag(row, "n")
        if rid.isdigit() and name:
            lib[rid] = name
    return lib

def bulk_update_ncash(xml_text, updates):
    """
    updates: {id_str: ncash_int}
    Returns (modified_xml_text, {id_str: bool_found})
    Single pass over all ROWs — no per-item full-text regex.
    """
    found = {id_: False for id_ in updates}
    def replace_row(m):
        block = m.group(0)
        rid   = _get_tag(block, "ID")
        if rid not in updates:
            return block
        found[rid] = True
        block = re.sub(r'<Ncash>\d+</Ncash>',
                       f'<Ncash>{updates[rid]}</Ncash>', block)
        return block
    result = ROW_RE.sub(replace_row, xml_text)
    return result, found

# ═══════════════════════════════════════════════════════════════════════════════
# CSV PARSER
# ═══════════════════════════════════════════════════════════════════════════════
def parse_csv_text(text):
    reader  = csv.DictReader(io.StringIO(text.strip()))
    rows    = list(reader)
    if not rows: return []
    headers = list(rows[0].keys())
    items, seen = [], set()

    def add(id_str, cost):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str)
            items.append({"id": id_str, "ticket_cost": cost})

    item_cols = [h for h in headers if re.match(r'Item\d+_ID', h, re.I)]
    if item_cols:
        for row in rows:
            for col in item_cols:
                add((row.get(col) or "").strip(), None)
        return items

    if len(headers) >= 2:
        id_col, cost_col = headers[0], headers[1]
        for row in rows:
            raw = (row.get(cost_col) or "").strip()
            try:    cost = float(raw)
            except: cost = None
            add((row.get(id_col) or "").strip(), cost)
        return items

    for row in rows:
        add((row.get(headers[0]) or "").strip(), None)
    return items

# ═══════════════════════════════════════════════════════════════════════════════
# APP
# ═══════════════════════════════════════════════════════════════════════════════
TARGET_FILES = {"itemparam2.xml", "itemparamcm2.xml", "itemparamex2.xml", "itemparamex.xml"}

class NCashUpdaterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 3 – NCash / Ticket Cost Updater")
        self.geometry("980x700")
        self.configure(bg="#1e1e2e")
        self.csv_items = []
        self.xml_text  = ""
        self.item_lib  = {}   # {id: name} from loaded ItemParam files
        self.mode_var  = tk.StringVar(value="uniform")
        self._build_load_screen()

    # ─────────────────────────────────────────────────────────────────────────
    # LOAD SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="NCASH / TICKET UPDATER",
                 font=("Consolas", 18, "bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(24, 4))
        tk.Label(self, text="Formula: NCash = round(tickets × 133)",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack(pady=(0, 8))

        csv_status = tk.StringVar(value="No file loaded")
        xml_status = tk.StringVar(value="No file loaded")

        def section(title):
            f = tk.LabelFrame(self, text=f"  {title}  ", bg="#1e1e2e", fg="#89b4fa",
                              font=("Consolas", 10, "bold"), bd=1, relief="groove")
            f.pack(fill="x", padx=30, pady=5)
            return f

        # ── CSV ──────────────────────────────────────────────────────────────
        csv_frm = section("Step 1 — Box Contents CSV (from Script 2, or ID list)")
        tk.Label(csv_frm, textvariable=csv_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)
        def load_csv():
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
            items = parse_csv_text(text)
            if not items:
                messagebox.showerror("Error", "No item IDs found in CSV."); return
            self.csv_items = items
            # Refresh names if lib already loaded
            if self.item_lib:
                for it in self.csv_items:
                    it["name"] = self.item_lib.get(it["id"], "")
            csv_status.set(f"✓  {os.path.basename(path)}  —  {len(items)} items")
        tk.Button(csv_frm, text="📂 Load", command=load_csv,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── ItemParam ────────────────────────────────────────────────────────
        xml_frm = section("Step 2 — ItemParam XML (pick any one of the 4 files)")
        tk.Label(xml_frm, textvariable=xml_status, bg="#1e1e2e",
                 fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)
        def load_xml():
            path = filedialog.askopenfilename(
                title="Select any one of the 4 ItemParam XML files",
                filetypes=[("XML","*.xml"),("All","*.*")])
            if not path: return
            folder   = os.path.dirname(path)
            combined, found = [], []
            for fname in os.listdir(folder):
                if fname.lower() in TARGET_FILES:
                    try:
                        with open(os.path.join(folder, fname),
                                  encoding="utf-8-sig", errors="replace") as f:
                            combined.append(f.read())
                        found.append(fname)
                    except Exception:
                        pass
            if not combined:
                messagebox.showerror("Error", "None of the 4 ItemParam XML files found."); return
            self.xml_text = "\n".join(combined)
            self.item_lib = build_item_lib(self.xml_text)
            # Back-fill names on already-loaded csv items
            for it in self.csv_items:
                it["name"] = self.item_lib.get(it["id"], "")
            xml_status.set(
                f"✓  {len(found)}/4 files  |  {len(self.item_lib)} items indexed  "
                f"({', '.join(found)})")
        tk.Button(xml_frm, text="📂 Load", command=load_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        # ── Mode ─────────────────────────────────────────────────────────────
        mode_frm = section("Step 3 — Mode")
        mf = tk.Frame(mode_frm, bg="#1e1e2e"); mf.pack(anchor="w", padx=10, pady=6)
        tk.Radiobutton(mf, text="Uniform  —  one ticket cost applied to every item",
                       variable=self.mode_var, value="uniform",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(anchor="w", pady=2)
        tk.Radiobutton(mf, text="Manual   —  set ticket cost per item individually",
                       variable=self.mode_var, value="manual",
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", font=("Consolas", 10)).pack(anchor="w", pady=2)

        def proceed():
            if not self.csv_items:
                messagebox.showwarning("Missing", "Load a CSV first."); return
            if not self.xml_text:
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
        tk.Label(self, text="Uniform Ticket Cost",
                 font=("Consolas", 14, "bold"), bg="#1e1e2e", fg="#f38ba8").pack(pady=(20, 4))
        tk.Label(self,
                 text=f"This value will be applied to all {len(self.csv_items)} items in the CSV.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10)).pack(pady=(0, 12))

        frm = tk.Frame(self, bg="#1e1e2e"); frm.pack()
        tk.Label(frm, text="Ticket Cost:", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Consolas", 12)).pack(side="left", padx=8)
        tv = tk.StringVar()
        ent = tk.Entry(frm, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                       insertbackground="#cdd6f4", font=("Consolas", 12), relief="flat")
        ent.pack(side="left", padx=8)
        ent.focus()

        ncash_var = tk.StringVar(value="NCash: —")
        tk.Label(self, textvariable=ncash_var, bg="#1e1e2e", fg="#a6e3a1",
                 font=("Consolas", 12, "bold")).pack(pady=8)

        def on_change(*_):
            try:    ncash_var.set(f"NCash: {round(float(tv.get()) * 133)}")
            except: ncash_var.set("NCash: —")
        tv.trace_add("write", on_change)

        def apply_uniform():
            try:
                cost = float(tv.get())
            except:
                messagebox.showwarning("Invalid", "Enter a valid ticket cost."); return
            for it in self.csv_items:
                it["ticket_cost"] = cost
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
        tk.Label(self, text="Leave blank to skip an item.  All blank items will trigger a warning.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 9)).pack(pady=(0, 4))

        # ── Scroll canvas ─────────────────────────────────────────────────
        outer = tk.Frame(self, bg="#1e1e2e"); outer.pack(fill="both", expand=True, padx=20, pady=4)
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

        # Header
        hdr = tk.Frame(cont, bg="#181825"); hdr.pack(fill="x", pady=2)
        for txt, w in [("Item ID", 12), ("Item Name", 36), ("Ticket Cost", 14), ("NCash (calc)", 14)]:
            tk.Label(hdr, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas", 9, "bold"), anchor="w").pack(side="left", padx=6, pady=4)

        ticket_vars = []
        for i, item in enumerate(self.csv_items):
            bg   = "#1e1e2e" if i % 2 == 0 else "#181825"
            row  = tk.Frame(cont, bg=bg); row.pack(fill="x")

            # ID (greyed)
            tk.Label(row, text=item["id"], width=12, bg=bg, fg="#585b70",
                     font=("Consolas", 9), anchor="w").pack(side="left", padx=6, pady=2)
            # Name from library
            name = item.get("name") or self.item_lib.get(item["id"], "—")
            tk.Label(row, text=name[:38], width=36, bg=bg, fg="#a6adc8",
                     font=("Consolas", 9), anchor="w").pack(side="left", padx=6, pady=2)
            # Ticket entry
            tv = tk.StringVar()
            prev_cost = item.get("ticket_cost")
            if prev_cost is not None:
                tv.set(str(prev_cost))
            ticket_vars.append(tv)
            tk.Entry(row, textvariable=tv, width=12, bg="#313244", fg="#cdd6f4",
                     insertbackground="#cdd6f4", font=("Consolas", 9),
                     relief="flat").pack(side="left", padx=6, pady=2)
            # NCash live calc
            ncash_lbl = tk.Label(row, text="—", width=14, bg=bg, fg="#a6e3a1",
                                 font=("Consolas", 9), anchor="w")
            ncash_lbl.pack(side="left", padx=6)
            def make_trace(var, lbl):
                def cb(*_):
                    try:    lbl.config(text=str(round(float(var.get()) * 133)))
                    except: lbl.config(text="—")
                var.trace_add("write", cb)
                cb()  # run once to populate from pre-filled value
            make_trace(tv, ncash_lbl)

        def confirm():
            blanks = []
            for i, item in enumerate(self.csv_items):
                raw = ticket_vars[i].get().strip()
                try:
                    item["ticket_cost"] = float(raw)
                except:
                    item["ticket_cost"] = None
                    blanks.append(item["id"])
            if blanks:
                ans = messagebox.askyesno(
                    "Missed a spot",
                    f"{len(blanks)} item(s) have no ticket cost and will be SKIPPED:\n\n"
                    + ", ".join(blanks[:20]) + ("…" if len(blanks) > 20 else "")
                    + "\n\nContinue anyway?")
                if not ans: return
            self._process_and_show()

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x", pady=6)
        tk.Button(bot, text="◀  Back", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(side="left", padx=14)
        tk.Button(bot, text="✓  Apply & Update XML", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", padx=16, pady=8).pack(side="right", padx=14)

    # ─────────────────────────────────────────────────────────────────────────
    # PROCESS
    # ─────────────────────────────────────────────────────────────────────────
    def _process_and_show(self):
        # Build updates dict for items that have a cost — skipped items excluded
        skipped  = {it["id"]: it.get("name","") for it in self.csv_items
                    if it["ticket_cost"] is None}
        updates  = {it["id"]: round(it["ticket_cost"] * 133)
                    for it in self.csv_items if it["ticket_cost"] is not None}
        ncash_map = {it["id"]: round(it["ticket_cost"] * 133)
                     for it in self.csv_items if it["ticket_cost"] is not None}
        name_map  = {it["id"]: it.get("name","") for it in self.csv_items}

        # Single pass over all XML
        modified_xml, found_map = bulk_update_ncash(self.xml_text, updates)

        results = []
        for item in self.csv_items:
            iid  = item["id"]
            name = name_map.get(iid, "")
            if iid in skipped:
                results.append((iid, name, None, False))
            else:
                results.append((iid, name, ncash_map[iid], found_map.get(iid, False)))
        self._build_output_screen(modified_xml, results)

    # ─────────────────────────────────────────────────────────────────────────
    # OUTPUT SCREEN
    # ─────────────────────────────────────────────────────────────────────────
    def _build_output_screen(self, modified_xml, results):
        self._clear()
        found_count   = sum(1 for _,_,n,f in results if f)
        skipped_count = sum(1 for _,_,n,_ in results if n is None)
        missing_count = sum(1 for _,_,n,f in results if n is not None and not f)

        summary = (f"✓ Updated: {found_count}    "
                   f"⚠ Not in XML: {missing_count}    "
                   f"— Skipped: {skipped_count}")
        tk.Label(self, text=summary, font=("Consolas", 10, "bold"),
                 bg="#1e1e2e", fg="#a6e3a1").pack(pady=8)
        if missing_count:
            missing = [id_ for id_,_,n,f in results if n is not None and not f]
            tk.Label(self, text="IDs not found: " + ", ".join(missing),
                     bg="#1e1e2e", fg="#f38ba8", font=("Consolas", 9)).pack()

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

        log_lines = ["ID             Name                                NCash        Status",
                     "─" * 74]
        for id_, name, ncash, found in results:
            name_str = (name or "—")[:32]
            if ncash is None:
                log_lines.append(f"{id_:<15}{name_str:<34}{'—':<13}SKIPPED")
            elif found:
                log_lines.append(f"{id_:<15}{name_str:<34}{ncash:<13}✓ Updated")
            else:
                log_lines.append(f"{id_:<15}{name_str:<34}{ncash:<13}⚠ Not found in XML")

        make_tab("Modified ItemParam.xml", modified_xml,      "ItemParam_updated.xml")
        make_tab("Update Log",             "\n".join(log_lines), "ncash_update_log.txt")
        nb.select(0)

        bot = tk.Frame(self, bg="#1e1e2e"); bot.pack(fill="x", pady=6)
        _exports = [
            ("ItemParam_updated.xml",  modified_xml),
            ("ncash_update_log.txt",   "\n".join(log_lines)),
        ]
        def export_all():
            folder = filedialog.askdirectory(title="Choose export folder")
            if not folder: return
            saved = []
            for fname, content in _exports:
                with open(os.path.join(folder, fname), "w", encoding="utf-8") as f:
                    f.write(content)
                saved.append(fname)
            messagebox.showinfo("Export Complete",
                f"Saved to:\n{folder}\n\n" + "\n".join(saved))
        tk.Button(bot, text="💾  Export All Files", command=export_all,
                  bg="#cba6f7", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", padx=20, pady=8).pack(side="left", padx=14)
        tk.Button(bot, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
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
    NCashUpdaterApp().mainloop()
