"""
Excel Analyzer + Q&A (Tkinter)
-------------------------------
Browse an Excel/CSV file, then choose:
  - "Analyze" mode: shows summary statistics of the data (like Excel's Analyze Data feature)
  - "Ask Questions" mode: ask natural-language questions about the data,
    answered using HuggingFace's local `table-question-answering` pipeline
    (TAPAS model). No API key / token required — runs fully locally.

Requirements:
    pip install pandas openpyxl transformers torch

Note: the first question asked downloads the TAPAS model (~440 MB) from
HuggingFace and caches it under %USERPROFILE%\\.cache\\huggingface — this can
take a few minutes depending on your connection. Every run after that is
instant since the model is loaded from the local cache.
"""

import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

# ----------------------------------------------------------------- Theme ---
BG = "#1e1e2e"          # app background
PANEL_BG = "#282838"    # panel/card background
ACCENT = "#7c3aed"      # purple accent (buttons)
ACCENT_HOVER = "#9333ea"
ANALYZE_COLOR = "#06b6d4"   # cyan
QA_COLOR = "#f472b6"        # pink
TEXT_LIGHT = "#f1f5f9"
TEXT_MUTED = "#94a3b8"
OUTPUT_BG = "#11111b"
OUTPUT_FG = "#a6e3a1"       # terminal-green text
ENTRY_BG = "#313244"
SUCCESS = "#22c55e"
WARNING = "#fbbf24"
ERROR = "#f87171"


class ExcelQAApp:
    MAX_TABLE_ROWS_FOR_QA = 500  # TAPAS works on small tables; truncate for speed/accuracy

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Analyzer & Q&A")
        self.root.geometry("880x680")
        self.root.configure(bg=BG)
        self.root.minsize(700, 550)

        self.file_path = tk.StringVar()
        self.mode = tk.StringVar(value="analyze")
        self.df = None
        self.current_path = None
        self.qa_pipeline = None  # loaded lazily on first question

        self._setup_styles()
        self._build_ui()

    # ------------------------------------------------------------ Styles ---
    def _setup_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background=BG)
        style.configure("Panel.TFrame", background=PANEL_BG)

        style.configure(
            "Header.TLabel", background=BG, foreground=TEXT_LIGHT,
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "Sub.TLabel", background=BG, foreground=TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Path.TEntry", fieldbackground=ENTRY_BG, foreground=TEXT_LIGHT,
            insertcolor=TEXT_LIGHT, bordercolor=ACCENT, lightcolor=ENTRY_BG,
            darkcolor=ENTRY_BG,
        )
        style.configure(
            "Question.TEntry", fieldbackground=ENTRY_BG, foreground=TEXT_LIGHT,
            insertcolor=TEXT_LIGHT, padding=8,
        )

        style.configure(
            "Accent.TButton", background=ACCENT, foreground="white",
            font=("Segoe UI", 10, "bold"), padding=(14, 8), borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", ACCENT_HOVER)])

        style.configure(
            "Ask.TButton", background=QA_COLOR, foreground="#1e1e2e",
            font=("Segoe UI", 10, "bold"), padding=(14, 8), borderwidth=0,
        )
        style.map("Ask.TButton", background=[("active", "#f9a8d4")])

        style.configure(
            "Run.TButton", background=ANALYZE_COLOR, foreground="#1e1e2e",
            font=("Segoe UI", 10, "bold"), padding=(14, 8), borderwidth=0,
        )
        style.map("Run.TButton", background=[("active", "#22d3ee")])

        style.configure(
            "Analyze.TRadiobutton", background=BG, foreground=ANALYZE_COLOR,
            font=("Segoe UI", 11, "bold"), padding=6,
        )
        style.map("Analyze.TRadiobutton", background=[("active", BG)])

        style.configure(
            "QA.TRadiobutton", background=BG, foreground=QA_COLOR,
            font=("Segoe UI", 11, "bold"), padding=6,
        )
        style.map("QA.TRadiobutton", background=[("active", BG)])

        style.configure(
            "Status.TLabel", background=PANEL_BG, foreground=SUCCESS,
            font=("Segoe UI", 9, "italic"), padding=6,
        )

        style.configure(
            "Colorful.Horizontal.TProgressbar", troughcolor=PANEL_BG,
            background=QA_COLOR, bordercolor=PANEL_BG, lightcolor=QA_COLOR,
            darkcolor=QA_COLOR, thickness=8,
        )

    # ---------------------------------------------------------------- UI ---
    def _build_ui(self):
        header = ttk.Frame(self.root, style="App.TFrame", padding=(20, 16, 20, 8))
        header.pack(fill="x")
        ttk.Label(header, text="📊 Excel Analyzer & Q&A", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Browse a spreadsheet, then Analyze it or Ask it questions — 100% local, no API key.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # --- Browse row ---
        browse_card = ttk.Frame(self.root, style="Panel.TFrame", padding=14)
        browse_card.pack(fill="x", padx=20, pady=(10, 6))

        ttk.Button(
            browse_card, text="📁 Browse...", style="Accent.TButton", command=self.browse_file
        ).pack(side="left")
        ttk.Entry(
            browse_card, textvariable=self.file_path, state="readonly", style="Path.TEntry"
        ).pack(side="left", padx=10, fill="x", expand=True, ipady=4)

        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(
            browse_card, textvariable=self.sheet_var, state="readonly", width=18,
        )
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # --- Mode row ---
        mode_frame = ttk.Frame(self.root, style="App.TFrame", padding=(20, 4))
        mode_frame.pack(fill="x")

        ttk.Radiobutton(
            mode_frame, text="🔍  Analyze Data", variable=self.mode, value="analyze",
            style="Analyze.TRadiobutton", command=self.on_mode_change,
        ).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(
            mode_frame, text="💬  Ask Questions (NotebookLM style)", variable=self.mode,
            value="qa", style="QA.TRadiobutton", command=self.on_mode_change,
        ).pack(side="left")

        # --- content container ---
        container = ttk.Frame(self.root, style="App.TFrame")
        container.pack(fill="both", expand=True, padx=20, pady=(6, 0))

        # --- Analyze panel ---
        self.analyze_frame = ttk.Frame(container, style="Panel.TFrame", padding=12)
        self.analyze_output = scrolledtext.ScrolledText(
            self.analyze_frame, wrap="word", bg=OUTPUT_BG, fg=OUTPUT_FG,
            insertbackground=OUTPUT_FG, font=("Consolas", 10), relief="flat",
            borderwidth=0,
        )
        self.analyze_output.pack(fill="both", expand=True)

        analyze_btn_row = ttk.Frame(self.analyze_frame, style="Panel.TFrame")
        analyze_btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(
            analyze_btn_row, text="⚡ Run Analysis", style="Run.TButton",
            command=self.run_analysis,
        ).pack(side="left")
        ttk.Button(
            analyze_btn_row, text="💾 Export Report", style="Accent.TButton",
            command=self.export_analysis,
        ).pack(side="left", padx=(8, 0))

        # --- QA panel ---
        self.qa_frame = ttk.Frame(container, style="Panel.TFrame", padding=12)

        self.columns_var = tk.StringVar(value="Browse a file to see its column names here.")
        ttk.Label(
            self.qa_frame, textvariable=self.columns_var, style="Sub.TLabel",
            background=PANEL_BG, wraplength=800, justify="left",
        ).pack(fill="x", pady=(0, 6), anchor="w")

        q_row = ttk.Frame(self.qa_frame, style="Panel.TFrame")
        q_row.pack(fill="x")
        self.question_var = tk.StringVar()
        question_entry = ttk.Entry(
            q_row, textvariable=self.question_var, style="Question.TEntry",
            font=("Segoe UI", 11),
        )
        question_entry.pack(side="left", fill="x", expand=True, ipady=4)
        question_entry.bind("<Return>", lambda e: self.ask_question())
        ttk.Button(
            q_row, text="✨ Ask", style="Ask.TButton", command=self.ask_question
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            q_row, text="💾 Export Chat", style="Accent.TButton", command=self.export_chat
        ).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(
            self.qa_frame, mode="indeterminate", style="Colorful.Horizontal.TProgressbar"
        )

        self.qa_output = scrolledtext.ScrolledText(
            self.qa_frame, wrap="word", bg=OUTPUT_BG, fg=OUTPUT_FG,
            insertbackground=OUTPUT_FG, font=("Consolas", 10), relief="flat",
            borderwidth=0,
        )
        self.qa_output.pack(fill="both", expand=True, pady=(8, 0))
        self.qa_output.tag_config("question", foreground="#89b4fa", font=("Consolas", 10, "bold"))
        self.qa_output.tag_config("answer", foreground=SUCCESS, font=("Consolas", 10, "bold"))
        self.qa_output.tag_config("error", foreground=ERROR)
        self.qa_output.tag_config("muted", foreground=TEXT_MUTED)

        # --- status bar ---
        status_bar = ttk.Frame(self.root, style="Panel.TFrame")
        status_bar.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel", anchor="w").pack(
            fill="x", padx=10
        )

        self.on_mode_change()

    def on_mode_change(self):
        if self.mode.get() == "analyze":
            self.qa_frame.pack_forget()
            self.analyze_frame.pack(fill="both", expand=True)
        else:
            self.analyze_frame.pack_forget()
            self.qa_frame.pack(fill="both", expand=True)

    # ------------------------------------------------------------- Load ---
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel/CSV file",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if not path:
            return

        self.current_path = path
        self.sheet_combo.pack_forget()

        try:
            if path.lower().endswith(".csv"):
                self.df = pd.read_csv(path)
            else:
                excel_file = pd.ExcelFile(path)
                sheet_names = excel_file.sheet_names
                if len(sheet_names) > 1:
                    self.sheet_combo["values"] = sheet_names
                    self.sheet_var.set(sheet_names[0])
                    self.sheet_combo.pack(side="left", padx=(0, 0))
                self.df = excel_file.parse(sheet_names[0])
        except Exception as exc:
            messagebox.showerror("Error loading file", str(exc))
            return

        self.file_path.set(path)
        self._on_data_loaded()

    def on_sheet_change(self, event=None):
        try:
            self.df = pd.read_excel(self.current_path, sheet_name=self.sheet_var.get())
        except Exception as exc:
            messagebox.showerror("Error loading sheet", str(exc))
            return
        self._on_data_loaded()

    def _on_data_loaded(self):
        path = self.file_path.get()
        sheet_note = f" [sheet: {self.sheet_var.get()}]" if self.sheet_var.get() else ""
        self.status_var.set(
            f"✅ Loaded {os.path.basename(path)}{sheet_note} — {self.df.shape[0]} rows x {self.df.shape[1]} cols"
        )
        self.analyze_output.delete("1.0", tk.END)
        self.qa_output.delete("1.0", tk.END)

        col_list = ", ".join(f"{self._excel_letter(i)}={c}" for i, c in enumerate(self.df.columns))
        self.columns_var.set(f"Columns: {col_list}")

    # --------------------------------------------------------- Analyze ---
    def run_analysis(self):
        if self.df is None:
            messagebox.showwarning("No file", "Please browse and select a file first.")
            return

        df = self.df
        lines = []
        lines.append(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns\n")

        lines.append("Columns and dtypes:")
        for col, dtype in df.dtypes.items():
            lines.append(f"  - {col}: {dtype}")
        lines.append("")

        lines.append("Missing values per column:")
        na_counts = df.isna().sum()
        for col, cnt in na_counts.items():
            if cnt > 0:
                lines.append(f"  - {col}: {cnt} missing")
        if na_counts.sum() == 0:
            lines.append("  (none)")
        lines.append("")

        numeric_df = df.select_dtypes(include="number")
        if not numeric_df.empty:
            lines.append("Numeric summary statistics:")
            lines.append(numeric_df.describe().to_string())
            lines.append("")

        cat_df = df.select_dtypes(exclude="number")
        if not cat_df.empty:
            lines.append("Categorical columns — top values:")
            for col in cat_df.columns:
                top = df[col].value_counts().head(5)
                lines.append(f"  {col}:")
                for val, cnt in top.items():
                    lines.append(f"    {val}: {cnt}")
            lines.append("")

        if not numeric_df.empty and numeric_df.shape[1] > 1:
            lines.append("Correlation matrix (numeric columns):")
            lines.append(numeric_df.corr().to_string())

        self.analyze_output.delete("1.0", tk.END)
        self.analyze_output.insert(tk.END, "\n".join(lines))
        self.status_var.set("✅ Analysis complete.")

    def export_analysis(self):
        content = self.analyze_output.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Nothing to export", "Run an analysis first.")
            return
        self._export_text(content, default_name="analysis_report.txt")

    def export_chat(self):
        content = self.qa_output.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Nothing to export", "Ask a question first.")
            return
        self._export_text(content, default_name="qa_chat.txt")

    def _export_text(self, content, default_name):
        path = filedialog.asksaveasfilename(
            title="Save as",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception as exc:
            messagebox.showerror("Error saving file", str(exc))
            return
        self.status_var.set(f"💾 Saved to {os.path.basename(path)}")

    # ---------------------------------------------------------------- QA ---
    def ask_question(self):
        if self.df is None:
            messagebox.showwarning("No file", "Please browse and select a file first.")
            return
        question = self.question_var.get().strip()
        if not question:
            return

        self._append_qa(f"Q: {question}\n", "question")
        self.question_var.set("")
        self.progress.pack(fill="x", pady=(8, 0))
        self.progress.start(12)
        threading.Thread(target=self._answer_question_thread, args=(question,), daemon=True).start()

    def _append_qa(self, text, tag=None):
        self.qa_output.insert(tk.END, text, tag)
        self.qa_output.see(tk.END)

    def _load_qa_pipeline(self):
        if self.qa_pipeline is not None:
            return self.qa_pipeline
        self.root.after(
            0, self.status_var.set,
            "⏳ First run: downloading HuggingFace TAPAS model (~440MB) — this can take a few minutes...",
        )
        from transformers import pipeline
        self.qa_pipeline = pipeline(
            task="table-question-answering",
            model="google/tapas-base-finetuned-wtq",
        )
        self.root.after(0, self.status_var.set, "✅ Model loaded and cached for future runs.")
        return self.qa_pipeline

    def _answer_question_thread(self, question):
        try:
            resolved_question = self._resolve_excel_letters(question, self.df.columns)

            direct = self._try_direct_aggregation(resolved_question)
            if direct is not None:
                self.root.after(0, self._append_qa, f"A: {direct}\n", "answer")
                self.root.after(0, self._append_qa, "   (computed directly with pandas)\n\n", "muted")
                return

            qa = self._load_qa_pipeline()
            table = self._prepare_table_for_qa()
            result = qa(table=table, query=resolved_question)

            answer = result.get("answer", "No answer found.")
            cells = result.get("cells", [])
            aggregator = result.get("aggregator", "")

            detail = f"A: {answer}"
            if aggregator:
                detail += f"  (aggregation: {aggregator})"
            self.root.after(0, self._append_qa, detail + "\n", "answer")
            if cells:
                self.root.after(0, self._append_qa, f"   cells used: {cells}\n", "muted")
            self.root.after(0, self._append_qa, "\n")
        except Exception as exc:
            msg = (
                f"Error answering question: {exc}\n"
                "   Tip: this model answers table lookups/aggregations, e.g.\n"
                "   \"how many rows have status = Done?\", \"what is the sum of Amount?\",\n"
                "   \"which row has the highest Sales?\" — not open-ended summaries.\n\n"
            )
            self.root.after(0, self._append_qa, msg, "error")
        finally:
            self.root.after(0, self._finish_qa)

    _AGG_PATTERNS = [
        (re.compile(r"\b(sum|total)\b", re.IGNORECASE), "sum"),
        (re.compile(r"\b(average|avg|mean)\b", re.IGNORECASE), "mean"),
        (re.compile(r"\b(max|maximum|highest|largest)\b", re.IGNORECASE), "max"),
        (re.compile(r"\b(min|minimum|lowest|smallest)\b", re.IGNORECASE), "min"),
        (re.compile(r"\b(count|number of rows|how many rows)\b", re.IGNORECASE), "count"),
    ]

    def _find_column_in_question(self, question):
        """Fuzzy-match a column name mentioned in the question (case/space
        insensitive substring match), longest name first to avoid partial
        matches stealing a shorter column's hit."""
        q_lower = question.lower()
        for col in sorted(self.df.columns, key=lambda c: -len(str(c))):
            if str(col).lower() in q_lower:
                return col
        return None

    def _try_direct_aggregation(self, question):
        """Handle common aggregation questions ("sum of X", "average of X",
        "how many rows are there") reliably with pandas instead of relying
        on TAPAS, which is unreliable on messy real-world headers/values."""
        q_lower = question.lower()

        if re.search(r"\b(how many rows|number of rows|row count|rows (are|is) there)\b", q_lower):
            return f"{len(self.df)} rows"

        op = None
        for pattern, name in self._AGG_PATTERNS:
            if pattern.search(question):
                op = name
                break
        if op is None:
            return None

        col = self._find_column_in_question(question)
        if col is None:
            return None

        if op == "count":
            return f"{self.df[col].count()} non-empty values in '{col}'"

        series = pd.to_numeric(
            self.df[col].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce",
        )
        if series.dropna().empty:
            return None  # not a numeric-looking column; let TAPAS try instead

        value = getattr(series, op)()
        return f"{value:g} (column '{col}')"

    def _prepare_table_for_qa(self):
        """TAPAS requires a table of plain strings with unique, string column
        names and no NaN/float values, or it raises cryptic errors such as
        "'float' object is not iterable"."""
        table = self.df.head(self.MAX_TABLE_ROWS_FOR_QA).copy()
        table.columns = [str(c) for c in table.columns]
        # de-duplicate any repeated column names
        seen = {}
        new_cols = []
        for c in table.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        table.columns = new_cols
        table = table.fillna("").astype(str)
        return table.reset_index(drop=True)

    @staticmethod
    def _excel_letter(index):
        """0 -> A, 1 -> B, ... 25 -> Z, 26 -> AA, ..."""
        letters = ""
        index += 1
        while index > 0:
            index, rem = divmod(index - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    def _resolve_excel_letters(self, question, columns):
        """TAPAS matches column names by their literal header text, not by
        spreadsheet position — so "column G" would otherwise be searched for
        as literal text "G". Translate phrases like "column G" or
        "G column" into the real header name."""
        letter_to_col = {self._excel_letter(i): col for i, col in enumerate(columns)}

        def replace(match):
            letter = match.group(1) or match.group(2)
            col = letter_to_col.get(letter.upper())
            return col if col else match.group(0)

        pattern = re.compile(
            r"\bcolumn\s+([A-Za-z]{1,2})\b|\b([A-Za-z]{1,2})\s+column\b",
            re.IGNORECASE,
        )
        return pattern.sub(replace, question)

    def _finish_qa(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.status_var.set("Ready.")


def main():
    root = tk.Tk()
    ExcelQAApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
