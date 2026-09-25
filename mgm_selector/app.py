"""MGM-Varvel RF/RK product selection UI (DS-021-0046 report)."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from datetime import datetime
from pathlib import Path

from .catalogue import ROOT, Catalogue, display_value
from .excel_report import (
    apply_excel_to_report,
    default_excel_name,
    default_pdf_name,
    generate_excel,
    load_excel,
    selection_fingerprint,
    validate_rows,
    workbook_has_edits,
)
from .pdf_report import generate_pdf
from .selection import SelectionEngine

def _exports_dir() -> Path:
    desktop = Path.home() / "Desktop" / "MGM-Varvel-Excel"
    try:
        desktop.mkdir(parents=True, exist_ok=True)
        probe = desktop / ".write_ok"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return desktop
    except OSError:
        fallback = ROOT / "exports"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


EXPORTS = _exports_dir()

APP_TITLE = "MGM-Varvel Product Selection"


class SelectorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x760")
        self.minsize(860, 680)
        self.configure(bg="#f4f6f8")

        self.engine = SelectionEngine(Catalogue())
        self.current_report = None
        self.confirmed_report = None
        self.last_excel_path = None
        self._excel_fingerprint = ""
        self.match_models: list[str] = []

        self.series = tk.StringVar(value="RF")
        self.power = tk.StringVar()
        self.ratio = tk.StringVar()
        self.model = tk.StringVar()
        self.mounting = tk.StringVar(value="M1")
        self.shaft = tk.StringVar(value="AC")
        self.motor = tk.StringVar(value="SMX")
        self.term = tk.StringVar(value="KP1")
        self.cable = tk.StringVar(value="A")

        self._build()
        self._fill_powers()
        self.series.trace_add("write", lambda *_: self._on_series_change())
        self.power.trace_add("write", lambda *_: self._on_power_change())
        self.ratio.trace_add("write", lambda *_: self._refresh())
        self.model.trace_add("write", lambda *_: self._refresh())
        for var in (self.mounting, self.shaft, self.motor, self.term, self.cable):
            var.trace_add("write", lambda *_: self._refresh())

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 6}
        header = tk.Frame(self, bg="#003366")
        header.pack(fill="x")
        tk.Button(
            header,
            text="Convert into Excel",
            command=self._convert_excel,
            bg="#0055a5",
            fg="white",
            activebackground="#003366",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=6,
            relief="flat",
            cursor="hand2",
        ).pack(side="right", padx=16, pady=12)
        tk.Label(
            header,
            text="MGM-VARVEL PRODUCT SELECTION",
            fg="white",
            bg="#003366",
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w", padx=16, pady=12)
        tk.Label(
            header,
            text="RF / RK catalogue selection  →  DS-021-0046 fields  →  Convert into Excel / PDF",
            fg="#c5d6ea",
            bg="#003366",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=16, pady=(0, 10))

        form = tk.LabelFrame(self, text="Selection inputs", bg="#f4f6f8", font=("Segoe UI", 10, "bold"))
        form.pack(fill="x", **pad)

        def add_row(r, label, widget):
            tk.Label(form, text=label, bg="#f4f6f8", font=("Segoe UI", 9)).grid(
                row=r, column=0, sticky="w", padx=8, pady=4
            )
            widget.grid(row=r, column=1, sticky="ew", padx=8, pady=4)

        form.columnconfigure(1, weight=1)
        self.cb_series = ttk.Combobox(
            form, textvariable=self.series, values=["RF", "RK"], state="readonly", width=40
        )
        self.cb_power = ttk.Combobox(form, textvariable=self.power, state="readonly")
        self.cb_ratio = ttk.Combobox(form, textvariable=self.ratio, state="readonly")
        self.cb_model = ttk.Combobox(form, textvariable=self.model, state="readonly")
        self.cb_mount = ttk.Combobox(
            form,
            textvariable=self.mounting,
            values=["M1", "M2", "M3", "M4", "M5", "M6"],
            state="readonly",
        )
        self.cb_shaft = ttk.Combobox(
            form,
            textvariable=self.shaft,
            values=["AC", "AS", "AD", "ACC"],
            state="readonly",
        )
        self.cb_motor = ttk.Combobox(
            form, textvariable=self.motor, values=["SMX", "BAX"], state="readonly"
        )
        self.cb_term = ttk.Combobox(
            form, textvariable=self.term, values=["KP1", "KP2", "KP3", "KP4"], state="readonly"
        )
        self.cb_cable = ttk.Combobox(
            form, textvariable=self.cable, values=["A", "B", "C", "D"], state="readonly"
        )

        add_row(0, "Series (RF / RK)", self.cb_series)
        add_row(1, "Rated motor power P1 (kW)", self.cb_power)
        add_row(2, "Overall gear ratio (i)", self.cb_ratio)
        add_row(3, "Configuration (if more than one)", self.cb_model)
        add_row(4, "Mounting position", self.cb_mount)
        add_row(5, "Output shaft (AC / AS / AD / ACC)", self.cb_shaft)
        add_row(6, "Motor (SMX standard / BAX brake)", self.cb_motor)
        add_row(7, "Terminal box position", self.cb_term)
        add_row(8, "Cable entry position", self.cb_cable)

        hint = tk.Label(
            form,
            text="AC = hollow with keyway (catalogue default). Ratios listed are only those tabulated for the selected power.",
            bg="#f4f6f8",
            fg="#555",
            font=("Segoe UI", 8),
            wraplength=700,
            justify="left",
        )
        hint.grid(row=9, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8))

        mid = tk.LabelFrame(
            self, text="Selected product  (DS-021-0046 fields only)", bg="#f4f6f8",
            font=("Segoe UI", 10, "bold")
        )
        mid.pack(fill="both", expand=True, **pad)

        cols = ("parameter", "value")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=16)
        self.tree.heading("parameter", text="DS-021-0046 field")
        self.tree.heading("value", text="Value")
        self.tree.column("parameter", width=280, anchor="w")
        self.tree.column("value", width=560, anchor="w")
        ys = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        ys.pack(side="right", fill="y", pady=8, padx=(0, 8))

        self.status = tk.StringVar(value="Select series, power and ratio.")
        tk.Label(self, textvariable=self.status, bg="#f4f6f8", fg="#333", font=("Segoe UI", 9)).pack(
            fill="x", padx=12
        )
        self.excel_status = tk.StringVar(
            value="Select the product, then click Convert into Excel to open the DS-021-0046 fields in Excel."
        )
        tk.Label(self, textvariable=self.excel_status, bg="#f4f6f8", fg="#0055a5", font=("Segoe UI", 9)).pack(
            fill="x", padx=12, pady=(0, 4)
        )

        btns = tk.Frame(self, bg="#f4f6f8")
        btns.pack(fill="x", pady=8, padx=10)
        ttk.Button(btns, text="Generate PDF report", command=self._save_pdf).pack(side="left", padx=4)
        ttk.Button(btns, text="Clear", command=self._clear).pack(side="left", padx=4)

    def _fill_powers(self) -> None:
        powers = self.engine.powers(self.series.get())
        labels = [self._fmt(p) for p in powers]
        self.cb_power["values"] = labels
        if labels:
            self.power.set(labels[0])
        else:
            self.power.set("")
        self._on_power_change()

    def _on_series_change(self) -> None:
        self._fill_powers()

    def _on_power_change(self) -> None:
        try:
            p = float(self.power.get())
        except ValueError:
            self.cb_ratio["values"] = []
            self.ratio.set("")
            return
        ratios = self.engine.ratios(self.series.get(), p)
        labels = [self._fmt(r) for r in ratios]
        self.cb_ratio["values"] = labels
        self.ratio.set(labels[0] if labels else "")
        self._refresh()

    @staticmethod
    def _fmt(v: float) -> str:
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return f"{v:.2f}".rstrip("0").rstrip(".")

    def _reset_excel_state(self) -> None:
        self.confirmed_report = None
        self.last_excel_path = None
        self._excel_fingerprint = ""
        self.excel_status.set(
            "Product changed. Click Convert into Excel to open this selection in Excel."
        )

    def _refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.current_report = None
        try:
            power = float(self.power.get())
            ratio = float(self.ratio.get())
        except ValueError:
            self.status.set("Select rated motor power and overall gear ratio (i).")
            return

        series = self.series.get()
        result = self.engine.lookup(series, power, ratio)
        if not result["ok"] and result.get("matches"):
            models = sorted({m["model"] for m in result["matches"]})
            self.cb_model["values"] = models
            if self.model.get() not in models:
                self.model.set(models[0] if models else "")
            result = self.engine.lookup(
                series,
                power,
                ratio,
                model=self.model.get() or None,
                mounting=self.mounting.get(),
                shaft_code=self.shaft.get(),
                motor_kind=self.motor.get(),
                terminal_box=self.term.get(),
                cable_entry=self.cable.get(),
            )
        elif result["ok"]:
            models = sorted({m["model"] for m in result["matches"]})
            self.cb_model["values"] = models
            if models and self.model.get() not in models:
                self.model.set(models[0])
            result = self.engine.lookup(
                series,
                power,
                ratio,
                model=self.model.get() or None,
                mounting=self.mounting.get(),
                shaft_code=self.shaft.get(),
                motor_kind=self.motor.get(),
                terminal_box=self.term.get(),
                cable_entry=self.cable.get(),
            )

        if not result["ok"]:
            self.status.set(result["error"])
            self._reset_excel_state()
            return

        report = result["report"]
        self.current_report = report
        fp = selection_fingerprint(report)
        if self._excel_fingerprint and fp != self._excel_fingerprint:
            self._reset_excel_state()
        display = self.confirmed_report or report
        for field in display["fields"]:
            self.tree.insert("", "end", values=(field["label"], display_value(field)))
        page = report["internal"].get("catalogue_page")
        self.status.set(
            f"Exact match: {report['internal']['model']}   |   catalogue page {page}   |   "
            f"{len(report['fields'])} DS-021-0046 fields"
        )

    def _preview_confirmed(self) -> None:
        if not self.confirmed_report:
            return
        self.tree.delete(*self.tree.get_children())
        for field in self.confirmed_report["fields"]:
            mark = "  [Edited]" if field.get("status") == "Edited" else ""
            self.tree.insert("", "end", values=(field["label"] + mark, display_value(field)))

    def _convert_excel(self) -> None:
        if not self.current_report:
            messagebox.showwarning(APP_TITLE, "Select a product first (power and ratio).")
            return
        folder = _exports_dir()
        path = folder / default_excel_name(self.current_report)
        if path.exists():
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = folder / default_excel_name(self.current_report).replace(
                ".xlsx", f"_{stamp}.xlsx"
            )
        try:
            generate_excel(self.current_report, path, include_edit_sheets=True)
        except PermissionError:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = folder / default_excel_name(self.current_report).replace(
                ".xlsx", f"_{stamp}.xlsx"
            )
            try:
                generate_excel(self.current_report, path, include_edit_sheets=True)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Could not create Excel.\n\n{exc}")
                return
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not create Excel.\n\n{exc}")
            return
        if not path.exists():
            messagebox.showerror(APP_TITLE, "Excel was not generated.")
            return
        self.last_excel_path = str(path)
        self._excel_fingerprint = selection_fingerprint(self.current_report)
        self.confirmed_report = self.current_report
        self.excel_status.set(
            f"Excel generated with the selected DS-021-0046 fields ({len(self.current_report['fields'])} rows):\n{path}"
        )
        opened = self._open_path(path)
        if not opened:
            messagebox.showinfo(APP_TITLE, f"Excel file saved here:\n{path}")

    def _open_path(self, path) -> bool:
        try:
            os.startfile(str(path))
            return True
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Excel file was saved but could not be opened.\n\n{path}\n\n{exc}")
            return False

    def _should_warn_overwrite(self) -> bool:
        if self.confirmed_report and (self.confirmed_report.get("excel_meta") or {}).get("edited_count"):
            return True
        if self.last_excel_path:
            try:
                return workbook_has_edits(self.last_excel_path)
            except Exception:
                return True
        return False

    def _import_excel(self) -> None:
        if not self.current_report:
            messagebox.showwarning(APP_TITLE, "Select a catalogue product first.")
            return
        initial = self.last_excel_path or ""
        path = filedialog.askopenfilename(
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=initial,
            title="Import edited DS-021-0046 Excel",
        )
        if not path:
            return
        try:
            loaded = load_excel(path)
            required = [f["key"] for f in self.current_report["fields"]]
            check = validate_rows(loaded["editable"], loaded["original"], required)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not read the Excel file.\n\n{exc}")
            return
        if not check["ok"]:
            messagebox.showerror(
                APP_TITLE,
                "The Excel file did not pass validation:\n\n" + "\n".join(check["errors"]),
            )
            return
        try:
            confirmed = apply_excel_to_report(self.current_report, loaded)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.confirmed_report = confirmed
        self.last_excel_path = path
        self._excel_fingerprint = selection_fingerprint(self.current_report)
        n_edit = len(check["edited"])
        extra = f"{n_edit} value(s) marked Edited." if n_edit else "No manual edits (all Auto)."
        warn = ""
        if check["warnings"]:
            warn = "\n" + "\n".join(check["warnings"])
        self.excel_status.set(f"Excel imported and confirmed: {path}\n{extra}{warn}")
        self._preview_confirmed()
        messagebox.showinfo(APP_TITLE, f"Excel imported.\n{extra}\n\nPDF generation will use these confirmed values.")

    def _review_changes(self) -> None:
        report = self.confirmed_report
        if not report:
            messagebox.showinfo(APP_TITLE, "Import an Excel workbook first to review changes.")
            return
        edits = (report.get("excel_meta") or {}).get("edits") or []
        if not edits:
            messagebox.showinfo(APP_TITLE, "No values were edited from the original catalogue data.")
            return
        win = tk.Toplevel(self)
        win.title("Review Excel changes")
        win.geometry("720x360")
        tk.Label(
            win,
            text="These values differ from the original automatically fetched catalogue data.",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=10, pady=8)
        cols = ("field", "original", "edited")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=12)
        tree.heading("field", text="Field")
        tree.heading("original", text="Original value")
        tree.heading("edited", text="Edited value")
        tree.column("field", width=220)
        tree.column("original", width=220)
        tree.column("edited", width=220)
        for item in edits:
            tree.insert("", "end", values=(item["label"], item["original"], item["edited"]))
        tree.pack(fill="both", expand=True, padx=10, pady=8)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)

    def _save_pdf(self) -> None:
        if not self.current_report:
            messagebox.showwarning(APP_TITLE, "No exact catalogue match to export.")
            return
        report = self.current_report
        if self.last_excel_path:
            try:
                loaded = load_excel(self.last_excel_path)
                required = [f["key"] for f in self.current_report["fields"]]
                check = validate_rows(loaded["editable"], loaded["original"], required)
                if check["ok"]:
                    report = apply_excel_to_report(self.current_report, loaded)
                    self.confirmed_report = report
            except Exception:
                report = self.current_report
        default = default_pdf_name(report)
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=default,
            title="Save DS-021-0046 report",
        )
        if not path:
            return
        generate_pdf(report, path)
        messagebox.showinfo(APP_TITLE, f"Report saved:\n{path}")

    def _edited_warning_dialog(self) -> str:
        result = {"choice": "cancel"}
        win = tk.Toplevel(self)
        win.title(APP_TITLE)
        win.transient(self)
        win.grab_set()
        tk.Label(
            win,
            text=(
                "Some values have been manually edited from the original catalogue data. "
                "Please verify the changes before generating the final PDF."
            ),
            wraplength=460,
            justify="left",
            font=("Segoe UI", 10),
        ).pack(padx=16, pady=16)

        def set_choice(value: str) -> None:
            result["choice"] = value
            win.destroy()

        row = tk.Frame(win)
        row.pack(pady=10)
        ttk.Button(row, text="Review Changes", command=lambda: set_choice("review")).pack(side="left", padx=6)
        ttk.Button(row, text="Generate PDF", command=lambda: set_choice("generate")).pack(side="left", padx=6)
        ttk.Button(row, text="Cancel", command=lambda: set_choice("cancel")).pack(side="left", padx=6)
        win.wait_window()
        return result["choice"]

    def _clear(self) -> None:
        self._reset_excel_state()
        self._fill_powers()


def main() -> None:
    app = SelectorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
