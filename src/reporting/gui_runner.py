from __future__ import annotations

import calendar
import datetime as dt
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from src.reporting.template_filler import generate_escalacrec_report

# Paleta clara
BG_COLOR = "#f7f7f9"
CARD_COLOR = "#ffffff"
ACCENT_COLOR = "#4b5563"
TEXT_COLOR = "#1f2933"
MUTED_COLOR = "#6b7280"
INPUT_BG = "#ffffff"

SOR_LOGO_PATH = Path(__file__).resolve().parents[2] / "input" / "Soriana-Logo.png"
PRGX_LOGO_PATH = Path(__file__).resolve().parents[2] / "input" / "Prgx-Logo.png"


def run_generation(vendor_entry: tk.Entry, ini_var: tk.StringVar, fin_var: tk.StringVar, status_var: tk.StringVar):
    vendor = vendor_entry.get().strip()
    fecha_ini = ini_var.get().strip()
    fecha_fin = fin_var.get().strip()

    if not vendor or not fecha_ini or not fecha_fin:
        messagebox.showwarning("Campos requeridos", "Completa proveedor, fecha inicio y fecha fin.")
        return

    status_var.set("Generando...")
    try:
        output_path = generate_escalacrec_report(vendor, fecha_ini, fecha_fin)
        status_var.set("Listo")
        messagebox.showinfo("Reporte generado", f"Archivo creado en:\n{output_path}")
    except Exception as exc:  # pragma: no cover
        status_var.set("Error")
        messagebox.showerror("Error", f"Ocurrió un error:\n{exc}")


def _set_style(root: tk.Misc):
    style = ttk.Style(master=root)
    style.theme_use("clam")
    style.configure("TFrame", background=BG_COLOR)
    style.configure("Card.TFrame", background=CARD_COLOR, relief="groove", borderwidth=1)
    style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10))
    style.configure("Card.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10))
    style.configure("Muted.TLabel", background=CARD_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 9))
    style.configure("TEntry", fieldbackground=INPUT_BG, background=INPUT_BG, foreground=TEXT_COLOR, insertcolor=TEXT_COLOR)
    style.configure("Accent.TButton", background=ACCENT_COLOR, foreground="#ffffff", font=("Segoe UI", 10))
    style.map(
        "Accent.TButton",
        background=[("active", "#374151")],
        foreground=[("active", "#ffffff")],
    )


def _load_logo(path: Path, max_width: int, max_height: int, master: tk.Misc | None = None) -> tk.PhotoImage | None:
    if not path.exists():
        return None
    try:
        img = tk.PhotoImage(master=master, file=str(path))
        # Escalado simple por subsample
        scale = max(img.width() / max_width, img.height() / max_height, 1)
        factor = int(scale) if scale > 1 else 1
        if factor > 1:
            img = img.subsample(factor, factor)
        return img
    except Exception:
        return None


def _build_date_selector(parent: tk.Misc, label: str, default: dt.date) -> tk.StringVar:
    frame = ttk.Frame(parent, style="Card.TFrame")
    ttk.Label(frame, text=label, style="Card.TLabel").grid(row=0, column=0, sticky="w")

    min_year = 2020
    years = [str(y) for y in range(min_year, default.year + 6)]
    months = [("01", "Ene"), ("02", "Feb"), ("03", "Mar"), ("04", "Abr"), ("05", "May"), ("06", "Jun"),
              ("07", "Jul"), ("08", "Ago"), ("09", "Sep"), ("10", "Oct"), ("11", "Nov"), ("12", "Dic")]

    year_var = tk.StringVar(value=str(default.year))
    month_var = tk.StringVar(value=f"{default.month:02d}")
    day_var = tk.StringVar(value=f"{default.day:02d}")
    date_str = tk.StringVar(value=default.isoformat())

    def update_days(*_):
        try:
            y = int(year_var.get())
            m = int(month_var.get())
            days_in_month = calendar.monthrange(y, m)[1]
            day_values = [f"{d:02d}" for d in range(1, days_in_month + 1)]
            day_combo["values"] = day_values
            if int(day_var.get()) > days_in_month:
                day_var.set(f"{days_in_month:02d}")
        except Exception:
            pass
        update_date()

    def update_date(*_):
        try:
            date_val = dt.date(int(year_var.get()), int(month_var.get()), int(day_var.get()))
            date_str.set(date_val.isoformat())
        except Exception:
            pass

    year_combo = ttk.Combobox(frame, textvariable=year_var, values=years, width=6, state="readonly")
    month_combo = ttk.Combobox(frame, textvariable=month_var, values=[m[0] for m in months], width=4, state="readonly")
    day_combo = ttk.Combobox(frame, textvariable=day_var, width=4, state="readonly")

    year_combo.grid(row=1, column=0, sticky="w", padx=(0, 8))
    month_combo.grid(row=1, column=1, sticky="w", padx=(0, 8))
    day_combo.grid(row=1, column=2, sticky="w")

    update_days()
    year_combo.bind("<<ComboboxSelected>>", update_days)
    month_combo.bind("<<ComboboxSelected>>", update_days)
    day_combo.bind("<<ComboboxSelected>>", update_date)

    return frame, date_str


def build_gui():
    root = tk.Tk()
    _set_style(root)
    root.title("Automatización Escala de Crecimiento")
    root.geometry("720x380")
    root.resizable(False, False)
    root.configure(bg=BG_COLOR)

    header = ttk.Frame(root, style="TFrame", padding=(20, 12))
    header.pack(fill="x")

    title_frame = ttk.Frame(header, style="TFrame")
    title_frame.pack(side="left", fill="x", expand=True)
    ttk.Label(title_frame, text="Automatización Escala de Crecimiento", font=("Segoe UI Semibold", 14)).pack(anchor="w")
    ttk.Label(title_frame, text="Genera el reporte ingresando proveedor y rango de fechas", style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

    logo_frame = ttk.Frame(header, style="TFrame")
    logo_frame.pack(side="right", padx=(8, 0))

    sor_logo = _load_logo(SOR_LOGO_PATH, max_width=110, max_height=32, master=root)
    prgx_logo = _load_logo(PRGX_LOGO_PATH, max_width=150, max_height=62, master=root)
    if sor_logo:
        lbl_s = ttk.Label(logo_frame, image=sor_logo, style="TLabel")
        lbl_s.image = sor_logo
        lbl_s.pack(side="left", padx=5)
    if prgx_logo:
        lbl_p = ttk.Label(logo_frame, image=prgx_logo, style="TLabel")
        lbl_p.image = prgx_logo
        lbl_p.pack(side="left", padx=5)

    main = ttk.Frame(root, style="Card.TFrame", padding=20)
    main.pack(fill="both", expand=True)

    ttk.Label(main, text="Proveedor:", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    vendor_entry = ttk.Entry(main, width=35)
    vendor_entry.grid(row=0, column=1, pady=5, sticky="w")

    today = dt.date.today()
    default_start = dt.date(2020, 1, 1)
    ini_frame, ini_var = _build_date_selector(main, "Fecha inicio:", default_start)
    ini_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
    fin_frame, fin_var = _build_date_selector(main, "Fecha fin:", today)
    fin_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

    status_var = tk.StringVar(value="Listo")
    status_label = ttk.Label(main, textvariable=status_var, style="Card.TLabel")
    status_label.grid(row=3, column=0, columnspan=2, pady=(12, 5), sticky="w")

    run_button = ttk.Button(
        main,
        text="Generar reporte",
        command=lambda: run_generation(vendor_entry, ini_var, fin_var, status_var),
        style="Accent.TButton",
    )
    run_button.grid(row=4, column=0, columnspan=2, pady=12)

    vendor_entry.focus()
    root.mainloop()


if __name__ == "__main__":
    build_gui()
