import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataAnalyzerApp:

    def __init__(self, root):

        self.root = root
        self.root.title("Corporate Data Analyzer")
        self.root.geometry("1200x720")

        self.file_path = None
        self.df = None
        self.report_df = None
        self.canvas = None

        self.build_gui()

    # ---------------- GUI ----------------

    def build_gui(self):

        style = ttk.Style()
        style.theme_use("clam")

        # HEADER
        header = ttk.Label(
            self.root,
            text="Corporate Data Analysis Tool",
            font=("Segoe UI", 16, "bold")
        )
        header.grid(row=0, column=0, columnspan=2, pady=10)

        # FILE FRAME
        file_frame = ttk.LabelFrame(self.root, text="File Selection", padding=10)
        file_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10)

        self.file_label = ttk.Label(file_frame, text="No file selected", width=80)
        self.file_label.grid(row=0, column=0, padx=5)

        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=1)
        ttk.Button(file_frame, text="Read File", command=self.read_file).grid(row=0, column=2)

        # DATA INFO
        info_frame = ttk.LabelFrame(self.root, text="Dataset Information", padding=10)
        info_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        self.rows_label = ttk.Label(info_frame, text="Rows:")
        self.rows_label.grid(row=0, column=0, padx=5)

        self.cols_label = ttk.Label(info_frame, text="Columns:")
        self.cols_label.grid(row=0, column=1, padx=5)

        self.columns_text = tk.Text(info_frame, height=2, width=100)
        self.columns_text.grid(row=1, column=0, columnspan=4, pady=5)

        # ANALYSIS OPTIONS
        options = ttk.LabelFrame(self.root, text="Analysis Options", padding=10)
        options.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10)

        ttk.Label(options, text="Group By").grid(row=0, column=0)
        self.group_col = ttk.Combobox(options, state="readonly", width=20)
        self.group_col.grid(row=0, column=1)

        ttk.Label(options, text="Value Column").grid(row=0, column=2)
        self.value_col = ttk.Combobox(options, state="readonly", width=20)
        self.value_col.grid(row=0, column=3)

        ttk.Label(options, text="Aggregate").grid(row=0, column=4)
        self.agg_method = ttk.Combobox(
            options,
            values=["sum", "avg", "max", "min"],
            state="readonly",
            width=10
        )
        self.agg_method.grid(row=0, column=5)

        ttk.Button(options, text="Preview Report",
                   command=self.preview_report).grid(row=0, column=6, padx=10)

        # TABLE FRAME
        table_frame = ttk.LabelFrame(self.root, text="Report Table", padding=10)
        table_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=5)

        self.tree = ttk.Treeview(table_frame)

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=scroll_y.set,
                            xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # CHART FRAME
        chart_frame = ttk.LabelFrame(self.root, text="Chart Preview", padding=10)
        chart_frame.grid(row=4, column=1, sticky="nsew", padx=10, pady=5)

        self.chart_container = ttk.Frame(chart_frame)
        self.chart_container.pack(fill="both", expand=True)

        # CHART CONTROLS
        chart_controls = ttk.LabelFrame(self.root, text="Chart Controls", padding=10)
        chart_controls.grid(row=5, column=0, columnspan=2, sticky="ew", padx=10)

        ttk.Label(chart_controls, text="Chart Type").grid(row=0, column=0)

        self.chart_type = ttk.Combobox(
            chart_controls,
            values=["Bar Chart", "Pie Chart", "Line Chart", "Column Chart"],
            state="readonly"
        )
        self.chart_type.grid(row=0, column=1)

        ttk.Button(chart_controls, text="Preview Chart",
                   command=self.preview_chart).grid(row=0, column=2, padx=10)

        ttk.Button(chart_controls, text="Export Chart PNG",
                   command=self.export_chart).grid(row=0, column=3)

        # EXPORT REPORT
        export_frame = ttk.LabelFrame(self.root, text="Export Report", padding=10)
        export_frame.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        ttk.Label(export_frame, text="Format").grid(row=0, column=0)

        self.export_format = ttk.Combobox(
            export_frame,
            values=["Excel (.xlsx)", "CSV (.csv)"],
            state="readonly"
        )
        self.export_format.grid(row=0, column=1)

        ttk.Button(export_frame, text="Export Report",
                   command=self.export_report).grid(row=0, column=2, padx=10)

        # GRID EXPANSION
        self.root.rowconfigure(4, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

    # ---------------- FILE ----------------

    def browse_file(self):

        path = filedialog.askopenfilename(
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")]
        )

        if path:
            self.file_path = path
            self.file_label.config(text=path)

    def read_file(self):

        if not self.file_path:
            messagebox.showerror("Error", "Select a file first")
            return

        if self.file_path.endswith(".csv"):
            self.df = pd.read_csv(self.file_path)
        else:
            self.df = pd.read_excel(self.file_path)

        rows, cols = self.df.shape

        self.rows_label.config(text=f"Rows: {rows}")
        self.cols_label.config(text=f"Columns: {cols}")

        self.columns_text.delete("1.0", tk.END)
        self.columns_text.insert(tk.END, ", ".join(self.df.columns))

        text_cols = self.df.select_dtypes(include=["object"]).columns.tolist()
        num_cols = self.df.select_dtypes(include=["number"]).columns.tolist()

        self.group_col["values"] = text_cols
        self.value_col["values"] = num_cols

    # ---------------- REPORT ----------------

    def preview_report(self):

        if self.df is None:
            messagebox.showerror("Error", "Load dataset first")
            return

        group = self.group_col.get()
        value = self.value_col.get()
        agg = self.agg_method.get()

        if not group or not value or not agg:
            messagebox.showerror("Error", "Select group, value column and aggregation")
            return

        agg_map = {"sum": "sum", "avg": "mean", "max": "max", "min": "min"}

        self.report_df = (
            self.df.groupby(group)[value]
            .agg(agg_map[agg])
            .sort_values(ascending=False)
            .reset_index()
        )

        self.reset_outputs()

        self.tree["columns"] = list(self.report_df.columns)
        self.tree["show"] = "headings"

        for col in self.report_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        for _, row in self.report_df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

    # ---------------- CHART ----------------

    def preview_chart(self):

        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first")
            return

        chart = self.chart_type.get()

        if not chart:
            messagebox.showerror("Error", "Select chart type")
            return

        x = self.report_df.iloc[:, 0]
        y = self.report_df.iloc[:, 1]

        fig = Figure(figsize=(5,4), dpi=100)
        ax = fig.add_subplot(111)

        if chart in ["Bar Chart", "Column Chart"]:
            ax.bar(x, y)

        elif chart == "Line Chart":
            ax.plot(x, y, marker="o")

        elif chart == "Pie Chart":
            ax.pie(y, labels=x, autopct="%1.1f%%")

        ax.set_title(chart)
        fig.tight_layout()

        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        self.canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    # ---------------- EXPORT ----------------

    def export_report(self):

        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first")
            return

        fmt = self.export_format.get()
        folder = os.path.dirname(self.file_path)

        if fmt.startswith("Excel"):
            path = os.path.join(folder, "report_output.xlsx")
            self.report_df.to_excel(path, index=False)
        else:
            path = os.path.join(folder, "report_output.csv")
            self.report_df.to_csv(path, index=False)

        messagebox.showinfo("Success", f"Report exported\n{path}")

    def export_chart(self):

        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first")
            return

        folder = os.path.dirname(self.file_path)
        path = os.path.join(folder, "chart_output.png")

        x = self.report_df.iloc[:, 0]
        y = self.report_df.iloc[:, 1]

        fig = Figure(figsize=(6,4))
        ax = fig.add_subplot(111)
        ax.bar(x, y)

        fig.savefig(path)

        messagebox.showinfo("Success", f"Chart saved\n{path}")

    # ---------------- RESET ----------------

    def reset_outputs(self):

        for row in self.tree.get_children():
            self.tree.delete(row)

        if self.canvas:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None


if __name__ == "__main__":

    root = tk.Tk()

    root.lift()
    root.attributes("-topmost", True)
    root.after(100, lambda: root.attributes("-topmost", False))
    root.focus_force()

    app = DataAnalyzerApp(root)

    root.mainloop()
