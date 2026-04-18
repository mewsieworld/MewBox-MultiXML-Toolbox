import csv
import io
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def generate_xml(rows):
    lines = []
    for item_id, name in rows:
        lines.append(f"  <ROW>\n    <ItemID>{item_id}</ItemID>\n    <Comment><![CDATA[{name}]]></Comment>\n  </ROW>")
    return "<RecycleExceptItem>\n" + "\n".join(lines) + "\n</RecycleExceptItem>"

def save_xml(xml):
    output_file = "RecycleExceptItem.xml"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(xml)
    return os.path.abspath(output_file)

def parse_csv_text(text):
    rows = []
    reader = csv.DictReader(io.StringIO(text.strip()))
    if not reader.fieldnames:
        return None, "No columns found."
    id_col = next((h for h in reader.fieldnames if h.strip().lower() == 'id'), None)
    name_col = next((h for h in reader.fieldnames if h.strip().lower() in ('name', 'comment')), None)
    if not id_col:
        return None, "CSV must have an 'ID' column."
    if not name_col:
        return None, "CSV must have a 'Name' or 'Comment' column."
    for row in reader:
        item_id = row[id_col].strip()
        name = row[name_col].strip()
        if item_id and name:
            rows.append((item_id, name))
    if not rows:
        return None, "No valid rows found."
    return rows, None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RecycleExceptItem XML Generator")
        self.resizable(True, True)
        self.configure(bg="#1a1a1a")
        self.minsize(700, 560)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background="#1a1a1a", borderwidth=0)
        style.configure("TNotebook.Tab", background="#2a2a2a", foreground="#aaa",
                        padding=[14, 6], font=("Courier New", 10))
        style.map("TNotebook.Tab", background=[("selected", "#333")], foreground=[("selected", "#c8a96e")])
        style.configure("TFrame", background="#1a1a1a")

        title = tk.Label(self, text="RecycleExceptItem Generator",
                         bg="#111", fg="#c8a96e",
                         font=("Courier New", 14, "bold"),
                         pady=12, padx=20, anchor="w")
        title.pack(fill="x")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=16, pady=(10, 0))

        tab_file = ttk.Frame(notebook)
        notebook.add(tab_file, text="  Load CSV File  ")
        self._build_file_tab(tab_file)

        tab_manual = ttk.Frame(notebook)
        notebook.add(tab_manual, text="  Manual Entry  ")
        self._build_manual_tab(tab_manual)

        out_frame = tk.Frame(self, bg="#1a1a1a")
        out_frame.pack(fill="both", expand=True, padx=16, pady=(8, 0))

        tk.Label(out_frame, text="XML OUTPUT", bg="#1a1a1a", fg="#555",
                 font=("Courier New", 9), anchor="w").pack(fill="x")

        text_scroll_frame = tk.Frame(out_frame, bg="#080808")
        text_scroll_frame.pack(fill="both", expand=True)

        self.output_text = tk.Text(text_scroll_frame, height=10, bg="#080808", fg="#a0c8e0",
                                   font=("Courier New", 11), relief="flat",
                                   insertbackground="#c8a96e", wrap="none")
        scrollbar = tk.Scrollbar(text_scroll_frame, command=self.output_text.yview, bg="#1a1a1a")
        self.output_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.output_text.pack(fill="both", expand=True)

        btn_frame = tk.Frame(self, bg="#1a1a1a")
        btn_frame.pack(fill="x", padx=16, pady=10)

        self.status_label = tk.Label(btn_frame, text="", bg="#1a1a1a", fg="#888",
                                     font=("Courier New", 10))
        self.status_label.pack(side="left")

        tk.Button(btn_frame, text="Save  RecycleExceptItem.xml",
                  command=self.save, bg="#2a2a10", fg="#c8a96e",
                  font=("Courier New", 10, "bold"), relief="flat",
                  padx=16, pady=6, cursor="hand2",
                  activebackground="#3a3a18", activeforeground="#c8a96e").pack(side="right")

        tk.Button(btn_frame, text="Copy XML",
                  command=self.copy, bg="#1e1e1e", fg="#888",
                  font=("Courier New", 10), relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  activebackground="#2a2a2a", activeforeground="#aaa").pack(side="right", padx=(0, 8))

    def _build_file_tab(self, parent):
        self.drop_frame = tk.Frame(parent, bg="#111",
                                   highlightbackground="#333", highlightthickness=2)
        self.drop_frame.pack(fill="both", expand=True, padx=20, pady=16)

        self.drop_label = tk.Label(self.drop_frame,
                                   text="Click to browse for a CSV file",
                                   bg="#111", fg="#555",
                                   font=("Courier New", 12),
                                   cursor="hand2")
        self.drop_label.pack(expand=True, pady=30)

        sub = tk.Label(self.drop_frame,
                       text="Expects columns: ID, Name (or Comment)",
                       bg="#111", fg="#3a3a3a", font=("Courier New", 9))
        sub.pack(pady=(0, 20))

        for widget in (self.drop_frame, self.drop_label, sub):
            widget.bind("<Button-1>", lambda e: self.browse_file())

    def _build_manual_tab(self, parent):
        tk.Label(parent, text="Paste CSV content or type rows below:",
                 bg="#1a1a1a", fg="#666", font=("Courier New", 9)).pack(anchor="w", padx=16, pady=(10, 2))

        self.manual_text = tk.Text(parent, height=8, bg="#0a0a0a", fg="#c8e0a0",
                                   font=("Courier New", 12), relief="flat",
                                   insertbackground="#c8a96e")
        self.manual_text.insert("1.0", "ID,Name\n2264,Miracle Blue Potion EV\n")
        self.manual_text.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        tk.Button(parent, text="Generate XML ->",
                  command=self.generate_from_manual,
                  bg="#1a2a0a", fg="#8bc34a",
                  font=("Courier New", 10, "bold"), relief="flat",
                  padx=14, pady=5, cursor="hand2",
                  activebackground="#253510").pack(anchor="e", padx=16, pady=(0, 10))

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, encoding='utf-8') as f:
                text = f.read()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        rows, err = parse_csv_text(text)
        if err:
            messagebox.showerror("Parse Error", err)
            return
        self.drop_label.config(text=f"  {os.path.basename(path)}  ({len(rows)} rows)", fg="#8bc34a")
        self.render_xml(rows)

    def generate_from_manual(self):
        text = self.manual_text.get("1.0", "end")
        rows, err = parse_csv_text(text)
        if err:
            messagebox.showerror("Parse Error", err)
            return
        self.render_xml(rows)

    def render_xml(self, rows):
        xml = generate_xml(rows)
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", xml)
        self.status_label.config(text=f"{len(rows)} rows ready", fg="#8bc34a")

    def copy(self):
        xml = self.output_text.get("1.0", "end").strip()
        if not xml:
            return
        self.clipboard_clear()
        self.clipboard_append(xml)
        self.status_label.config(text="Copied to clipboard", fg="#c8a96e")

    def save(self):
        xml = self.output_text.get("1.0", "end").strip()
        if not xml:
            messagebox.showwarning("Nothing to save", "Generate XML first.")
            return
        path = save_xml(xml)
        self.status_label.config(text=f"Saved -> {path}", fg="#8bc34a")
        messagebox.showinfo("Saved", f"File saved to:\n{path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
