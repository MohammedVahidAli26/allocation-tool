import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import math
from tkinter.ttk import Progressbar, Style
import os

class AllocationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("✨ Excel Allocation Tool")
        self.df = None
        self.file_path = ""

        # Scrollable layout
        canvas = tk.Canvas(root, borderwidth=0, background="#f0f3f5")
        self.frame = tk.Frame(canvas, background="#f0f3f5")
        vsb = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((10, 10), window=self.frame, anchor="nw", tags="self.frame")

        self.frame.bind("<Configure>", lambda event, canvas=canvas: canvas.configure(scrollregion=canvas.bbox("all")))
        root.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        # Theme
        style = Style()
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))

        self.build_ui()

    def build_ui(self):
        tk.Label(self.frame, text="🌈 Excel Allocation Assistant", font=("Segoe UI", 18, "bold"), fg="#2980b9", bg="#f0f3f5").pack(pady=10)

        # Upload Section
        self.section("1️⃣ Upload File", "#1abc9c")
        ttk.Button(self.frame, text="📁 Upload Excel File", command=self.upload_file).pack(pady=5)
        self.file_label = ttk.Label(self.frame, text="", foreground="#34495e", background="#f0f3f5")
        self.file_label.pack()

        # Column Section
        self.section("2️⃣ Select Columns", "#f39c12")
        ttk.Label(self.frame, text="Group By Column:", background="#f0f3f5").pack(anchor="w", padx=10)
        self.column_dropdown = ttk.Combobox(self.frame, state="readonly")
        self.column_dropdown.pack(fill="x", padx=10, pady=5)

        ttk.Label(self.frame, text="Assign To Column:", background="#f0f3f5").pack(anchor="w", padx=10)
        self.assign_column_dropdown = ttk.Combobox(self.frame, state="readonly")
        self.assign_column_dropdown.pack(fill="x", padx=10, pady=5)

        # HC Input
        self.section("3️⃣ Headcount", "#9b59b6")
        hc_row = tk.Frame(self.frame, bg="#f0f3f5")
        hc_row.pack(pady=5)
        ttk.Label(hc_row, text="Enter Headcount:").pack(side=tk.LEFT)
        self.hc_entry = tk.Entry(hc_row, width=10)
        self.hc_entry.pack(side=tk.LEFT, padx=10)
        ttk.Button(hc_row, text="🔍 Calculate", command=self.set_hc).pack(side=tk.LEFT)
        self.group_label = ttk.Label(self.frame, text="", background="#f0f3f5", foreground="#2e86de", justify="left")
        self.group_label.pack(pady=5)

        # Associate Input
        self.section("4️⃣ Associates", "#e74c3c")
        ttk.Label(self.frame, text="Enter names (comma-separated):", background="#f0f3f5").pack(anchor="w", padx=10)
        self.name_entry = tk.Entry(self.frame, width=70)
        self.name_entry.pack(padx=10, pady=5)
        ttk.Button(self.frame, text="🚀 Start Allocation", command=self.allocate).pack(pady=10)

        # Side panel for breakdown
        self.breakdown_label = ttk.Label(self.frame, text="", background="#f0f3f5", justify="left", foreground="#2c3e50")
        self.breakdown_label.pack(pady=10)

        # Progress bar
        self.section("5️⃣ Finish & Save", "#3498db")
        ttk.Label(self.frame, text="Progress:", font=("Segoe UI", 10, "bold"), background="#f0f3f5").pack(pady=(5, 0))
        self.progress = Progressbar(self.frame, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(pady=5)
        self.progress["value"] = 0

        self.download_btn = ttk.Button(self.frame, text="💾 Download File", command=self.download_file, state=tk.DISABLED)
        self.download_btn.pack(pady=10)

    def section(self, title, color):
        tk.Label(self.frame, text=title, font=("Segoe UI", 12, "bold"), fg="white", bg=color).pack(fill="x", pady=(15, 5))

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if not self.file_path:
            return
        try:
            self.df = pd.read_excel(self.file_path, sheet_name="Production Allocation")
            self.file_label.config(text=f"✅ Loaded: {len(self.df)} rows")
            cols = list(self.df.columns)
            self.column_dropdown['values'] = cols
            self.assign_column_dropdown['values'] = cols
            self.column_dropdown.set("product_id")
            default_assign_col = "Production Associate Name" if "Production Associate Name" in cols else cols[0]
            self.assign_column_dropdown.set(default_assign_col)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")

    def set_hc(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Upload a file first.")
            return
        try:
            hc = int(self.hc_entry.get())
            if hc <= 0:
                raise ValueError

            selected_col = self.column_dropdown.get()
            if selected_col not in self.df.columns:
                messagebox.showerror("Invalid Column", "Select a valid allocation column.")
                return

            grouped_counts = self.df[selected_col].value_counts()
            total_groups = len(grouped_counts)
            total_rows = len(self.df)
            groups_per_head = math.ceil(total_groups / hc)

            est_row_dist = [0] * hc
            group_sizes = grouped_counts.tolist()
            for i, size in enumerate(group_sizes):
                est_row_dist[i % hc] += size
            rows_per_head = max(est_row_dist)

            self.group_label.config(
                text=f"🧮 Unique Groups: {total_groups}\n"
                     f"📄 Total Rows: {total_rows}\n"
                     f"🎯 Target Per Head:\n"
                     f"   • ~{groups_per_head} groups\n"
                     f"   • ~{rows_per_head} rows"
            )
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number for HC.")

    def allocate(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Upload a file first.")
            return

        associates = [a.strip() for a in self.name_entry.get().split(",") if a.strip()]
        if not associates:
            messagebox.showwarning("Warning", "Enter associate names.")
            return

        group_by_col = self.column_dropdown.get()
        assign_col = self.assign_column_dropdown.get()

        if group_by_col not in self.df.columns or assign_col not in self.df.columns:
            messagebox.showerror("Error", "Selected columns are invalid.")
            return

        grouped = self.df.groupby(group_by_col)
        group_sizes = [(gid, len(gdf)) for gid, gdf in grouped]
        group_sizes.sort(key=lambda x: -x[1])  # Sort by descending size

        allocation = {name: {"ids": [], "rows": 0} for name in associates}

        for gid, count in group_sizes:
            least_loaded = min(allocation.items(), key=lambda x: x[1]["rows"])[0]
            allocation[least_loaded]["ids"].append(gid)
            allocation[least_loaded]["rows"] += count

        # Update side panel
        breakdown = "\n📊 Allocation Breakdown:\n"
        for name in associates:
            breakdown += f"• {name:<15} → {len(allocation[name]['ids'])} groups, {allocation[name]['rows']} rows\n"
        self.breakdown_label.config(text=breakdown)

        # Progress bar
        self.progress["maximum"] = len(self.df)
        self.progress["value"] = 0
        self.frame.update_idletasks()

        # Apply assignment
        allocation_map = {}
        for name in associates:
            for gid in allocation[name]["ids"]:
                allocation_map[gid] = name

        assigned_names = []
        for i, row in self.df.iterrows():
            assigned_names.append(allocation_map.get(row[group_by_col], ""))
            if i % 10 == 0:
                self.progress["value"] += 10
                self.frame.update_idletasks()

        self.df[assign_col] = assigned_names
        self.progress["value"] = self.progress["maximum"]
        self.download_btn.config(state=tk.NORMAL)
        messagebox.showinfo("✅ Success", "Allocation complete!")

    def download_file(self):
        original_name = os.path.basename(self.file_path)
        base_name = os.path.splitext(original_name)[0]
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{base_name}_allocated.xlsx"
        )
        if save_path:
            try:
                self.df.to_excel(save_path, index=False)
                messagebox.showinfo("Saved", f"✅ File saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("750x650")
    app = AllocationApp(root)
    root.mainloop()
