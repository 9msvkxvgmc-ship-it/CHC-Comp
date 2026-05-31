"""
Graphical interface for charge-supervision-matrix.
"""

import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import runner, suggest
from .config import Config


_CATEGORY_OPTIONS = [
    "Exclude (remove from output)",
    "Add as APP (matrix column)",
    "Reclassify as MD (matrix row)",
]

_CATEGORY_MAP = {
    "Exclude (remove from output)": "exclude",
    "Add as APP (matrix column)": "app",
    "Reclassify as MD (matrix row)": "md",
}


class _ScrollableFrame(ttk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        _sb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self.inner = ttk.Frame(self._canvas)
        self.inner.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self._canvas.configure(yscrollcommand=_sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        _sb.pack(side="right", fill="y")
        self._canvas.bind_all("<MouseWheel>", self._scroll)
        self._canvas.bind_all("<Button-4>", self._scroll)
        self._canvas.bind_all("<Button-5>", self._scroll)

    def _scroll(self, event):
        if event.num == 4:
            self._canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self._canvas.yview_scroll(1, "units")
        else:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Charge Supervision Matrix")
        self.geometry("1000x780")
        self.minsize(800, 600)

        self._analysis: dict | None = None
        self._report_type = tk.StringVar(value="inpatient")
        self._app_vars: dict[str, tk.BooleanVar] = {}
        self._md_vars: dict[str, tk.BooleanVar] = {}
        self._review_vars: dict[str, tk.StringVar] = {}

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame", background="#f2f2f2")
        s.configure("TLabelframe", background="#f2f2f2")
        s.configure("TLabelframe.Label", font=("Arial", 11, "bold"), background="#f2f2f2")
        s.configure("TLabel", background="#f2f2f2", font=("Arial", 10))
        s.configure("TCheckbutton", background="#f2f2f2")
        s.configure("Head.TLabel", font=("Arial", 15, "bold"), background="#f2f2f2")
        s.configure("Sub.TLabel", font=("Arial", 10), foreground="#666666", background="#f2f2f2")
        s.configure("ColHdr.TLabel", font=("Arial", 9, "bold"), background="#f2f2f2")
        s.configure("Warn.TLabel", font=("Arial", 9), foreground="#c0392b", background="#f2f2f2")
        s.configure("Big.TButton", font=("Arial", 11, "bold"), padding="8 4")
        self.configure(bg="#f2f2f2")

    def _build_ui(self):
        ttk.Label(self, text="Charge Supervision Matrix", style="Head.TLabel").pack(
            anchor="w", padx=14, pady=(12, 6)
        )

        # Step 1 — file selection + report type
        f1 = ttk.LabelFrame(self, text="Step 1 — Select Input File", padding=10)
        f1.pack(fill="x", padx=12, pady=(0, 4))

        self._inp_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self._inp_var, width=58).pack(
            side="left", fill="x", expand=True, padx=(0, 6)
        )
        ttk.Button(f1, text="Browse…", command=self._browse_input).pack(side="left")

        # Report type selector
        rt_frame = ttk.Frame(f1)
        rt_frame.pack(side="left", padx=(14, 0))
        ttk.Label(rt_frame, text="Type:").pack(side="left", padx=(0, 4))
        ttk.Radiobutton(
            rt_frame, text="Inpatient", variable=self._report_type, value="inpatient"
        ).pack(side="left")
        ttk.Radiobutton(
            rt_frame, text="Outpatient", variable=self._report_type, value="outpatient"
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            f1, text="  Analyze Signers  ", style="Big.TButton", command=self._start_analyze
        ).pack(side="left", padx=(12, 0))

        self._info_lbl = ttk.Label(self, text="", style="Sub.TLabel")
        self._info_lbl.pack(anchor="w", padx=16)

        # Step 2 — signer review
        f2 = ttk.LabelFrame(self, text="Step 2 — Review & Categorize Signers", padding=8)
        f2.pack(fill="both", expand=True, padx=12, pady=4)

        self._scroll = _ScrollableFrame(f2)
        self._scroll.pack(fill="both", expand=True)
        self._body = self._scroll.inner

        ttk.Label(
            self._body,
            text="Select an input file and click Analyze Signers to begin.",
            style="Sub.TLabel",
        ).pack(pady=24)

        # Step 3 — output + generate
        f3 = ttk.LabelFrame(self, text="Step 3 — Output & Generate", padding=10)
        f3.pack(fill="x", padx=12, pady=(0, 4))

        ttk.Label(f3, text="Save to:").pack(side="left")
        self._out_var = tk.StringVar()
        ttk.Entry(f3, textvariable=self._out_var, width=60).pack(
            side="left", fill="x", expand=True, padx=6
        )
        ttk.Button(f3, text="Browse…", command=self._browse_output).pack(side="left")
        ttk.Button(
            f3, text="  Generate Report  ", style="Big.TButton", command=self._generate
        ).pack(side="left", padx=(10, 0))

        # Status bar
        self._status_var = tk.StringVar(value="Ready.")
        ttk.Label(
            self, textvariable=self._status_var, relief="sunken", anchor="w", padding="6 3"
        ).pack(fill="x", side="bottom")

    # ── File pickers ──────────────────────────────────────────────

    def _browse_input(self):
        p = filedialog.askopenfilename(
            title="Select All Signed Charges export",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self._inp_var.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if p:
            self._out_var.set(p)

    # ── Analysis ──────────────────────────────────────────────────

    def _start_analyze(self):
        path = self._inp_var.get().strip()
        if not path:
            messagebox.showerror("No file selected", "Please browse to an input file first.")
            return
        self._status_var.set("Analyzing…")
        self.update_idletasks()
        rt = self._report_type.get()
        threading.Thread(target=self._run_analyze, args=(path, rt), daemon=True).start()

    def _run_analyze(self, path, report_type):
        try:
            result = suggest.analyze(path, report_type=report_type)
            self.after(0, lambda: self._show_review(result))
        except Exception as exc:
            self.after(0, lambda: (
                self._status_var.set("Error during analysis."),
                messagebox.showerror("Analysis failed", str(exc)),
            ))

    def _show_review(self, result):
        self._analysis = result
        inp = Path(self._inp_var.get())
        if not self._out_var.get():
            # Default output to ~/ChargeSupervision Reports/ — a home-directory
            # subfolder that is always writable by unsigned .app bundles.
            # (Downloads, Documents, and Desktop are TCC-protected on macOS.)
            out_dir = Path.home() / "ChargeSupervision Reports"
            out_dir.mkdir(exist_ok=True)
            self._out_var.set(str(out_dir / (inp.stem + "_report.xlsx")))

        dr = result.get("date_range", "")
        tc = result.get("total_charges", 0)
        self._info_lbl.config(text=f"  {dr}    |    {tc:,} charges parsed")

        for w in self._body.winfo_children():
            w.destroy()
        self._app_vars.clear()
        self._md_vars.clear()
        self._review_vars.clear()

        cc = result["charge_counts"]
        sole = result["sole_signer_for"]

        # Confirmed APPs
        self._section(
            "Confirmed APPs  —  will appear as matrix columns  "
            "(uncheck to exclude a specific APP)"
        )
        if result["app_signers"]:
            self._col_headers("Name", "Charges", "Include?")
            for name in result["app_signers"]:
                var = tk.BooleanVar(value=True)
                self._app_vars[name] = var
                row = ttk.Frame(self._body)
                row.pack(fill="x", padx=14, pady=1)
                ttk.Label(row, text=name, width=44, anchor="w").pack(side="left")
                ttk.Label(
                    row, text=f"{cc.get(name, 0):,}", width=10, anchor="w", style="Sub.TLabel"
                ).pack(side="left")
                ttk.Checkbutton(row, variable=var).pack(side="left")
        else:
            ttk.Label(self._body, text="  (none detected)", style="Sub.TLabel").pack(
                anchor="w", padx=14
            )

        # MDs signing charges
        self._section(
            "MDs signing charges  —  check to reclassify as supervising MD  "
            "(prevents them from appearing as APP columns)"
        )
        if result["md_signers"]:
            self._col_headers("Name", "Charges", "Reclassify as MD?")
            for name in result["md_signers"]:
                var = tk.BooleanVar(value=True)
                self._md_vars[name] = var
                row = ttk.Frame(self._body)
                row.pack(fill="x", padx=14, pady=1)
                ttk.Label(row, text=name, width=44, anchor="w").pack(side="left")
                ttk.Label(
                    row, text=f"{cc.get(name, 0):,}", width=10, anchor="w", style="Sub.TLabel"
                ).pack(side="left")
                ttk.Checkbutton(row, variable=var).pack(side="left")
        else:
            ttk.Label(self._body, text="  (none detected)", style="Sub.TLabel").pack(
                anchor="w", padx=14
            )

        # Unrecognized signers
        needs = result["other_signers"] + result["unknown_signers"]
        self._section(
            f"Unrecognized signers  —  {len(needs)} signer(s) need categorization"
        )
        if needs:
            self._col_headers("Name", "Charges", "Categorize as", "Notes")
            for name in needs:
                is_sole = name in sole
                default = (
                    "Add as APP (matrix column)"
                    if is_sole
                    else "Exclude (remove from output)"
                )
                var = tk.StringVar(value=default)
                self._review_vars[name] = var

                row = ttk.Frame(self._body)
                row.pack(fill="x", padx=14, pady=3)
                ttk.Label(row, text=name, width=40, anchor="w").pack(side="left")
                ttk.Label(
                    row, text=f"{cc.get(name, 0):,}", width=10, anchor="w", style="Sub.TLabel"
                ).pack(side="left")
                ttk.Combobox(
                    row, textvariable=var, values=_CATEGORY_OPTIONS, state="readonly", width=28
                ).pack(side="left")

                if is_sole:
                    mds = ", ".join(sole[name])
                    ttk.Label(
                        row, text=f"  ⚠  Sole signer for: {mds}", style="Warn.TLabel"
                    ).pack(side="left")
                    self._attach_sole_warning(name, var, mds)
        else:
            ttk.Label(
                self._body, text="  (all signers recognized — no review needed)", style="Sub.TLabel"
            ).pack(anchor="w", padx=14)

        n_review = len(needs)
        if n_review:
            self._status_var.set(
                f"Analysis complete — {n_review} signer(s) need review. "
                "Adjust categories above, then click Generate Report."
            )
        else:
            self._status_var.set("Analysis complete — all signers recognized. Ready to generate.")

    def _section(self, title: str):
        ttk.Separator(self._body).pack(fill="x", padx=6, pady=(12, 2))
        ttk.Label(self._body, text=title, font=("Arial", 10, "bold")).pack(
            anchor="w", padx=14, pady=(0, 4)
        )

    def _col_headers(self, *labels):
        widths = [44, 10, 28, 0]
        row = ttk.Frame(self._body)
        row.pack(fill="x", padx=14, pady=(0, 2))
        for i, lbl in enumerate(labels):
            w = widths[i] if i < len(widths) else 0
            kw = {"width": w, "anchor": "w"} if w else {"anchor": "w"}
            ttk.Label(row, text=lbl, style="ColHdr.TLabel", **kw).pack(side="left")

    def _attach_sole_warning(self, name: str, var: tk.StringVar, mds: str):
        def _check(*_):
            if var.get().startswith("Exclude"):
                messagebox.showwarning(
                    "Sole Signer Warning",
                    f'"{name}" is the sole signer for:\n  {mds}\n\n'
                    "Excluding this person will remove those supervising MDs from the "
                    "Supervision Matrix entirely.\n\n"
                    'Consider using "Add as APP (matrix column)" instead.',
                )
        var.trace_add("write", _check)

    # ── Generation ────────────────────────────────────────────────

    def _generate(self):
        if not self._analysis:
            messagebox.showerror("Not ready", "Please analyze an input file first.")
            return

        # Always confirm the save location through the native macOS save panel.
        # This is required for unsigned apps to get write permission to protected
        # folders (Downloads, Documents, Desktop, etc.).
        suggested = self._out_var.get().strip()
        initial_dir  = str(Path(suggested).parent) if suggested else str(Path.home())
        initial_file = Path(suggested).name        if suggested else "report.xlsx"
        out = filedialog.asksaveasfilename(
            title="Save report as",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not out:
            return  # user cancelled
        self._out_var.set(out)

        excluded, add_app, reclassify = [], [], []

        for name, var in self._app_vars.items():
            if not var.get():
                excluded.append(name)

        for name, var in self._md_vars.items():
            if var.get():
                reclassify.append(name)

        for name, var in self._review_vars.items():
            cat = _CATEGORY_MAP.get(var.get(), "exclude")
            if cat == "exclude":
                excluded.append(name)
            elif cat == "app":
                add_app.append(name)
            elif cat == "md":
                reclassify.append(name)

        config = Config(
            excluded_signers=excluded,
            reclassify_as_supervising_md=reclassify,
            add_to_app_list=add_app,
        )

        self._status_var.set("Generating report…")
        self.update_idletasks()
        inp = self._inp_var.get().strip()
        rt = self._report_type.get()

        def _run():
            import traceback
            try:
                written = runner.run(input_path=inp, output_path=out, config=config, report_type=rt)
                self.after(0, lambda: self._done(written))
            except Exception as exc:
                tb = traceback.format_exc()
                try:
                    with open("/tmp/csm_error.log", "w") as f:
                        f.write(tb)
                except Exception:
                    pass
                msg = str(exc) or "(unknown — see /tmp/csm_error.log)"
                self.after(0, lambda m=msg: self._status_var.set(f"Error: {m}"))

        threading.Thread(target=_run, daemon=True).start()

    def _done(self, path: str):
        self._status_var.set(f"Report saved: {path}")
        if messagebox.askyesno("Done!", f"Report written to:\n{path}\n\nOpen it now?"):
            if sys.platform == "darwin":
                subprocess.run(["open", path])
            elif sys.platform == "win32":
                subprocess.run(["start", path], shell=True)
            else:
                subprocess.run(["xdg-open", path])


def main():
    app = App()
    app.mainloop()
