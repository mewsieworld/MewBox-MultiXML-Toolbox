"""
SCRIPT 3 – NCash / Ticket Cost Updater
────────────────────────────────────────
Workflow:
  1. Load the box-contents CSV produced by Script 2
     (columns: BoxID, BoxName, Item1_ID, Item2_ID, …)
     OR any CSV that has at least an ID column and optionally a TicketCost column.
  2. Load your ItemParam.xml (the main item database).
  3. For each item ID found in the CSV:
       - If the CSV has a TicketCost column, use it directly.
       - Otherwise, prompt the user to enter the ticket cost for each item
         (showing ID + greyed-out calculated NCash preview).
  4. Update the <Ncash> value in the ItemParam.xml for each matching ID.
     Formula: NCash = round(tickets * 133)
  5. Export the modified ItemParam.xml rows.

Requirements: Python 3.x  (standard library only)
Run: python script3_ncash_updater.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv, io, re, math, os

# ─── Helpers ─────────────────────────────────────────────────────────────────
def parse_csv_text(text):
    """
    Returns a list of dicts with keys: id, ticket_cost (None if not present).
    Accepts:
      - CSV from Script 2: BoxID, BoxName, Item1_ID, Item2_ID, …
        → flattens all ItemN_ID values into individual IDs
      - Simple 2-col CSV: ID, TicketCost
      - Single-column CSV: just IDs
    """
    reader = csv.DictReader(io.StringIO(text.strip()))
    rows   = list(reader)
    if not rows:
        return []

    headers = list(rows[0].keys())
    items   = []
    seen    = set()

    def add(id_str, cost):
        id_str = id_str.strip()
        if id_str and id_str.isdigit() and id_str not in seen:
            seen.add(id_str)
            items.append({"id": id_str, "ticket_cost": cost})

    # Detect Script 2 style (BoxID, BoxName, Item#_ID…)
    item_cols = [h for h in headers if re.match(r'Item\d+_ID', h, re.I)]
    if item_cols:
        for row in rows:
            for col in item_cols:
                val = row.get(col) or ""
                add(val.strip(), None)
        return items

    # Detect simple 2-col: ID, TicketCost
    if len(headers) >= 2:
        id_col   = headers[0]
        cost_col = headers[1]
        for row in rows:
            raw_cost = (row.get(cost_col) or "").strip()
            cost = None
            try:
                cost = float(raw_cost)
            except:
                pass
            add((row.get(id_col) or "").strip(), cost)
        return items

    id_col = headers[0]
    for row in rows:
        add((row.get(id_col) or "").strip(), None)
    return items

def find_and_update_ncash(xml_text, id_str, new_ncash):
    """
    Find the <ROW> block containing <ID>{id_str}</ID> and update its <Ncash> value.
    Returns (modified_xml, found_bool).
    """
    pattern = re.compile(
        r'(<ROW>(?:(?!</ROW>).)*?<ID>' + re.escape(id_str) + r'</ID>(?:(?!</ROW>).)*?</ROW>)',
        re.DOTALL
    )
    found = [False]

    def replacer(m):
        found[0] = True
        block = m.group(1)
        block = re.sub(r'<Ncash>\d+</Ncash>', f'<Ncash>{new_ncash}</Ncash>', block)
        return block

    result = pattern.sub(replacer, xml_text)
    return result, found[0]

# ─── App ─────────────────────────────────────────────────────────────────────
class NCashUpdaterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Script 3 – NCash / Ticket Cost Updater")
        self.geometry("900x680")
        self.configure(bg="#1e1e2e")
        self.csv_items  = []   # list of {id, ticket_cost}
        self.xml_text   = ""
        self._build_load_screen()

    # ── Screen 0: Load files ──────────────────────────────────────────────
    def _build_load_screen(self):
        self._clear()
        tk.Label(self, text="NCASH / TICKET UPDATER", font=("Consolas", 18, "bold"),
                 bg="#1e1e2e", fg="#f38ba8").pack(pady=(30, 5))
        tk.Label(self,
                 text="Updates <Ncash> values in ItemParam.xml based on ticket costs.\n"
                      "Formula: NCash = round(tickets × 133)",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 10), justify="center").pack(pady=8)

        csv_status = tk.StringVar(value="No file loaded")
        xml_status = tk.StringVar(value="No file loaded")

        def section(title, status_var, load_cmd):
            frm = tk.LabelFrame(self, text=f"  {title}  ", bg="#1e1e2e", fg="#89b4fa",
                                font=("Consolas", 10, "bold"), bd=1, relief="groove")
            frm.pack(fill="x", padx=30, pady=6)
            tk.Label(frm, textvariable=status_var, bg="#1e1e2e",
                     fg="#6c7086", font=("Consolas", 9)).pack(side="left", padx=10)
            tk.Button(frm, text="📂 Load", command=load_cmd,
                      bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                      relief="flat", padx=10, pady=4).pack(side="right", padx=8, pady=6)

        def load_csv():
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
            if not path: return
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
            items = parse_csv_text(text)
            if not items:
                messagebox.showerror("Error", "No item IDs found in CSV.")
                return
            self.csv_items = items
            has_cost = any(it["ticket_cost"] is not None for it in items)
            csv_status.set(f"✓  {os.path.basename(path)}  —  {len(items)} items"
                           + ("  (with costs)" if has_cost else "  (no costs — will prompt)"))

        TARGET_FILES = {"itemparam2.xml", "itemparamcm2.xml", "itemparamex2.xml", "itemparamex.xml"}

        def load_xml():
            path = filedialog.askopenfilename(
                title="Select any one of the 4 ItemParam XML files",
                filetypes=[("XML","*.xml"),("All","*.*")])
            if not path: return
            folder   = os.path.dirname(path)
            combined = []
            found    = []
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
                messagebox.showerror("Error", "None of the 4 ItemParam XML files found in that folder.")
                return
            self.xml_text = "\n".join(combined)
            xml_status.set(f"✓  {len(found)}/4 files loaded: {', '.join(found)}")

        section("Box Contents CSV (from Script 2, or ID list)", csv_status, load_csv)
        section("ItemParam.xml  (the main item database)",       xml_status, load_xml)

        def proceed():
            if not self.csv_items:
                messagebox.showwarning("Missing", "Load a CSV first.")
                return
            if not self.xml_text:
                messagebox.showwarning("Missing", "Load ItemParam.xml first.")
                return
            has_cost = any(it["ticket_cost"] is not None for it in self.csv_items)
            if has_cost:
                self._process_with_costs()
            else:
                self._build_prompt_screen()

        tk.Button(self, text="▶  Continue →", command=proceed,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 12, "bold"),
                  relief="flat", padx=20, pady=8).pack(pady=20)

    # ── Screen 1a: Prompt for costs ───────────────────────────────────────
    def _build_prompt_screen(self):
        self._clear()
        tk.Label(self, text="Enter Ticket Costs", font=("Consolas", 14, "bold"),
                 bg="#1e1e2e", fg="#f38ba8").pack(pady=10)
        tk.Label(self,
                 text="NCash is calculated automatically.  Leave blank to skip an item.",
                 bg="#1e1e2e", fg="#a6adc8", font=("Consolas", 9)).pack(pady=2)

        # Scroll canvas
        outer = tk.Frame(self, bg="#1e1e2e")
        outer.pack(fill="both", expand=True, padx=20, pady=6)
        canvas = tk.Canvas(outer, bg="#1e1e2e", highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        container = tk.Frame(canvas, bg="#1e1e2e")
        win_id = canvas.create_window((0, 0), window=container, anchor="nw")
        container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        # Header row
        hdr = tk.Frame(container, bg="#181825")
        hdr.pack(fill="x", pady=2)
        for col, w, txt in [(0,12,"Item ID"),(1,16,"Ticket Cost"),(2,16,"NCash (calc)")]:
            tk.Label(hdr, text=txt, width=w, bg="#181825", fg="#89b4fa",
                     font=("Consolas", 9, "bold"), anchor="w").grid(row=0, column=col, padx=6, pady=4)

        ticket_vars = []
        ncash_labels = []

        for i, item in enumerate(self.csv_items):
            bg = "#1e1e2e" if i % 2 == 0 else "#181825"
            row_frm = tk.Frame(container, bg=bg)
            row_frm.pack(fill="x")

            tk.Label(row_frm, text=item["id"], width=12, bg=bg, fg="#cdd6f4",
                     font=("Consolas", 9), anchor="w").grid(row=0, column=0, padx=6, pady=2)

            tv = tk.StringVar()
            ticket_vars.append(tv)
            ent = tk.Entry(row_frm, textvariable=tv, width=16,
                           bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                           font=("Consolas", 9), relief="flat")
            ent.grid(row=0, column=1, padx=6, pady=2)

            ncash_lbl = tk.Label(row_frm, text="—", width=16, bg=bg, fg="#a6e3a1",
                                 font=("Consolas", 9), anchor="w")
            ncash_lbl.grid(row=0, column=2, padx=6)
            ncash_labels.append(ncash_lbl)

            def make_trace(var, lbl):
                def cb(*_):
                    try:
                        t = float(var.get())
                        lbl.config(text=str(round(t * 133)))
                    except:
                        lbl.config(text="—")
                var.trace_add("write", cb)
            make_trace(tv, ncash_lbl)

        def confirm():
            for i, item in enumerate(self.csv_items):
                raw = ticket_vars[i].get().strip()
                try:
                    item["ticket_cost"] = float(raw)
                except:
                    item["ticket_cost"] = None   # skip
            self._process_with_costs()

        tk.Button(self, text="✓  Apply & Update XML", command=confirm,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 11, "bold"),
                  relief="flat", padx=16, pady=8).pack(pady=10)

    # ── Process ───────────────────────────────────────────────────────────
    def _process_with_costs(self):
        modified_xml = self.xml_text
        results = []   # (id, ncash, found)

        for item in self.csv_items:
            if item["ticket_cost"] is None:
                results.append((item["id"], None, False))
                continue
            ncash = round(item["ticket_cost"] * 133)
            modified_xml, found = find_and_update_ncash(modified_xml, item["id"], ncash)
            results.append((item["id"], ncash, found))

        self._build_output_screen(modified_xml, results)

    # ── Screen 2: Output ──────────────────────────────────────────────────
    def _build_output_screen(self, modified_xml, results):
        self._clear()
        found_count   = sum(1 for _, _, f in results if f)
        skipped_count = sum(1 for _, n, _ in results if n is None)
        missing_count = sum(1 for _, n, f in results if n is not None and not f)

        summary = (f"✓ Updated: {found_count}    "
                   f"⚠ Not found in XML: {missing_count}    "
                   f"— Skipped (no cost): {skipped_count}")
        tk.Label(self, text=summary, font=("Consolas", 10, "bold"),
                 bg="#1e1e2e", fg="#a6e3a1").pack(pady=10)

        if missing_count:
            missing = [id_ for id_, n, f in results if n is not None and not f]
            tk.Label(self, text="IDs not found: " + ", ".join(missing),
                     bg="#1e1e2e", fg="#f38ba8", font=("Consolas", 9)).pack()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        # Tab 1: modified XML
        xml_frm = tk.Frame(nb, bg="#1e1e2e")
        nb.add(xml_frm, text="Modified ItemParam.xml")
        xml_txt = scrolledtext.ScrolledText(xml_frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4")
        xml_txt.pack(fill="both", expand=True, padx=4, pady=4)
        xml_txt.insert("1.0", modified_xml)
        xml_txt.config(state="disabled")

        def copy_xml():
            self.clipboard_clear(); self.clipboard_append(modified_xml)
            messagebox.showinfo("Copied", "XML copied to clipboard.")

        def save_xml():
            path = filedialog.asksaveasfilename(initialfile="ItemParam_updated.xml",
                                                defaultextension=".xml",
                                                filetypes=[("XML","*.xml"),("All","*.*")])
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(modified_xml)
                messagebox.showinfo("Saved", f"Saved to {path}")

        brow1 = tk.Frame(xml_frm, bg="#1e1e2e")
        brow1.pack(fill="x")
        tk.Button(brow1, text="📋 Copy", command=copy_xml,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)
        tk.Button(brow1, text="💾 Save As…", command=save_xml,
                  bg="#a6e3a1", fg="#1e1e2e", font=("Consolas", 9),
                  relief="flat", padx=10, pady=4).pack(side="left", padx=6, pady=4)

        # Tab 2: summary log
        log_frm = tk.Frame(nb, bg="#1e1e2e")
        nb.add(log_frm, text="Update Log")
        log_txt = scrolledtext.ScrolledText(log_frm, font=("Consolas", 9),
                                            bg="#181825", fg="#cdd6f4")
        log_txt.pack(fill="both", expand=True, padx=4, pady=4)
        log_lines = ["ID           NCash       Status"]
        log_lines.append("-" * 40)
        for id_, ncash, found in results:
            if ncash is None:
                log_lines.append(f"{id_:<14} {'—':<12} SKIPPED (no cost entered)")
            elif found:
                log_lines.append(f"{id_:<14} {ncash:<12} ✓ Updated")
            else:
                log_lines.append(f"{id_:<14} {ncash:<12} ⚠ ID not found in XML")
        log_txt.insert("1.0", "\n".join(log_lines))
        log_txt.config(state="disabled")

        tk.Button(self, text="◀  Start Over", command=self._build_load_screen,
                  bg="#313244", fg="#cdd6f4", font=("Consolas", 10),
                  relief="flat", padx=12, pady=6).pack(pady=8)

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


if __name__ == "__main__":
    app = NCashUpdaterApp()
    app.mainloop()
