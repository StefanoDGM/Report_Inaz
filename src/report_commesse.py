import copy
import datetime as dt
import os
import re
import sys
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover
    tk = None
    filedialog = None
    messagebox = None
    ttk = None


SOURCE_SHEET_NAME = "Stampa Commesse Dipendente"
DETAIL_SHEET_NAME = "Stampa Commesse Dipendente"
SUMMARY_SHEET_NAME = "Riepilogo Viaggi"

TOTAL_PROJ_MARKER = "Totale"
TRAVEL_KEYWORDS = ("COMMESSA", "CHIUSURA")
SUMMARY_SOURCE_COLUMNS = [
    (7, "Reparto"),
    (8, "Cod. Progetto"),
    (9, "Progetto"),
    (10, "Cod. Argomento"),
    (11, "Argomento"),
    (14, "Codice dipendente"),
    (15, "Nominativo"),
    (16, "Data"),
]


@dataclass
class GroupSummary:
    order: int
    base_row: list[Any]
    total_hours: float
    travel_gross_hours: float
    travel_net_hours: float

    @property
    def gross_ratio(self) -> float:
        if self.total_hours <= 0:
            return 0.0
        return self.travel_gross_hours / self.total_hours

    @property
    def net_ratio(self) -> float:
        if self.total_hours <= 0:
            return 0.0
        return self.travel_net_hours / self.total_hours


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def input_dir() -> Path:
    return app_root() / "input"


def output_dir() -> Path:
    return app_root() / "output"


def ensure_workspace() -> None:
    input_dir().mkdir(parents=True, exist_ok=True)
    output_dir().mkdir(parents=True, exist_ok=True)


def latest_input_file() -> Optional[Path]:
    files = sorted(
        [p for p in input_dir().glob("*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: (p.stat().st_mtime, p.name.lower()),
    )
    return files[-1] if files else None


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def parse_duration(value: Any) -> Optional[float]:
    if value is None:
        return None

    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 3600.0

    if isinstance(value, dt.datetime):
        return value.hour + value.minute / 60.0 + value.second / 3600.0

    if isinstance(value, dt.time):
        return value.hour + value.minute / 60.0 + value.second / 3600.0

    if is_number(value):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    match = re.fullmatch(r"(?:(\d+):)?(\d{1,3}):(\d{2})", text)
    if match:
        hours_part = int(match.group(1) or match.group(2))
        minutes_part = int(match.group(3))
        return hours_part + minutes_part / 60.0

    match = re.fullmatch(r"(\d+(?:[.,]\d+)?)", text)
    if match:
        return float(match.group(1).replace(",", "."))

    return None


def to_centesimal_hours(value: Any) -> Any:
    parsed = parse_duration(value)
    if parsed is None:
        return value
    return round(parsed, 5)


def is_date_value(value: Any) -> bool:
    return isinstance(value, (dt.datetime, dt.date))


def is_total_row(row: list[Any]) -> bool:
    for idx in (8, 9, 10, 11):
        if idx < len(row):
            cell = row[idx]
            if isinstance(cell, str) and cell.strip().lower() == TOTAL_PROJ_MARKER.lower():
                return True
    return False


def travel_row(row: list[Any]) -> bool:
    for idx in (8, 9, 10, 11):
        if idx < len(row):
            cell = row[idx]
            if isinstance(cell, str) and any(keyword in cell.upper() for keyword in TRAVEL_KEYWORDS):
                return True
    return False


def normalize_text(value: Any) -> str:
    return str(value or "").strip().casefold()


def is_generic_commessa_name(project_name: str) -> bool:
    return normalize_text(project_name) == "commessa"


def copy_cell_style(src, dst) -> None:
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy.copy(src.protection)


def clone_sheet_layout(src_ws, dst_ws) -> None:
    dst_ws.sheet_format = copy.copy(src_ws.sheet_format)
    dst_ws.sheet_properties = copy.copy(src_ws.sheet_properties)
    dst_ws.page_margins = copy.copy(src_ws.page_margins)
    dst_ws.page_setup = copy.copy(src_ws.page_setup)
    dst_ws.print_options = copy.copy(src_ws.print_options)
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.auto_filter.ref = src_ws.auto_filter.ref

    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter] = copy.copy(dim)

    for row_idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx] = copy.copy(dim)

    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))


def build_detail_sheet(src_ws, dst_ws) -> None:
    clone_sheet_layout(src_ws, dst_ws)

    for r in range(1, src_ws.max_row + 1):
        for c in range(1, src_ws.max_column + 1):
            src_cell = src_ws.cell(r, c)
            dst_cell = dst_ws.cell(r, c)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)

            if c == 18:
                converted = to_centesimal_hours(src_cell.value)
                dst_cell.value = converted
                if isinstance(converted, (int, float)):
                    dst_cell.number_format = "0.00000"

    for c in range(1, src_ws.max_column + 1):
        dst_ws.column_dimensions[get_column_letter(c)].width = src_ws.column_dimensions[get_column_letter(c)].width

    dst_ws.auto_filter.ref = src_ws.auto_filter.ref or f"A1:{get_column_letter(src_ws.max_column)}{src_ws.max_row}"


def build_summary_rows(src_ws) -> list[GroupSummary]:
    grouped: OrderedDict[tuple[Any, Any, Any], dict[str, Any]] = OrderedDict()

    for row_idx in range(2, src_ws.max_row + 1):
        values = [src_ws.cell(row_idx, c).value for c in range(1, src_ws.max_column + 1)]
        if is_total_row(values):
            continue
        if len(values) < 18:
            continue

        employee_code = values[14]
        employee_name = values[15]
        day_value = values[16]

        if not employee_name or not is_date_value(day_value):
            continue

        key = (employee_code, employee_name, day_value)
        info = grouped.get(key)
        if info is None:
            grouped[key] = {
                "order": len(grouped),
                "rows": [],
                "total_hours": 0.0,
                "travel_gross_hours": 0.0,
            }
            info = grouped[key]

        info["rows"].append(values)
        hours = parse_duration(values[17]) or 0.0
        info["total_hours"] += hours
        if travel_row(values):
            info["travel_gross_hours"] += hours

    summaries: list[GroupSummary] = []
    for info in grouped.values():
        base_row = list(info["rows"][0])
        gross_travel = round(info["travel_gross_hours"], 5)
        is_manutentori = str(base_row[7] or "").strip().upper() == "MANUTENTORI"
        if is_manutentori:
            travel_net_hours = round(gross_travel - 1.0, 5)
        else:
            travel_net_hours = 0.0

        summaries.append(
            GroupSummary(
                order=info["order"],
                base_row=base_row,
                total_hours=round(info["total_hours"], 5),
                travel_gross_hours=gross_travel if is_manutentori else 0.0,
                travel_net_hours=travel_net_hours,
            )
        )

    return summaries


def build_summary_sheet(src_ws, dst_ws) -> None:
    clone_sheet_layout(src_ws, dst_ws)

    headers = [label for _, label in SUMMARY_SOURCE_COLUMNS]
    headers.extend([
        "Ore lavorate totali",
        "Ore viaggio lorde",
        "Ore viaggio nette",
        "% viaggio lorde",
        "% viaggio nette",
    ])

    summaries = build_summary_rows(src_ws)

    title = dst_ws["A1"]
    title.value = "Riepilogo Viaggi"
    title.font = Font(bold=True, size=14)

    for c, header in enumerate(headers, start=1):
        cell = dst_ws.cell(4, c)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        thin = Side(style="thin", color="999999")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, summary in enumerate(summaries, start=5):
        values = [summary.base_row[idx] for idx, _ in SUMMARY_SOURCE_COLUMNS]
        values.extend(
            [
                summary.total_hours,
                summary.travel_gross_hours,
                summary.travel_net_hours,
                round(summary.gross_ratio, 4),
                round(summary.net_ratio, 4),
            ]
        )

        for c, value in enumerate(values, start=1):
            cell = dst_ws.cell(row_idx, c)
            cell.value = value
            if c in (9, 10, 11):
                cell.number_format = "0.00000"
            elif c in (12, 13):
                cell.number_format = "0.0000%"
            elif c == 8:
                cell.number_format = "dd/mm/yyyy"

    dst_ws.freeze_panes = "A5"
    last_row = len(summaries) + 4
    last_col = get_column_letter(len(headers))
    dst_ws.auto_filter.ref = f"A4:{last_col}{last_row}"

    for c in range(1, len(headers) + 1):
        if c == 1:
            dst_ws.column_dimensions[get_column_letter(c)].width = 22
        elif c == 2:
            dst_ws.column_dimensions[get_column_letter(c)].width = 16
        elif c == 3:
            dst_ws.column_dimensions[get_column_letter(c)].width = 26
        elif c in (4, 5):
            dst_ws.column_dimensions[get_column_letter(c)].width = 26
        elif c == 6:
            dst_ws.column_dimensions[get_column_letter(c)].width = 16
        elif c == 7:
            dst_ws.column_dimensions[get_column_letter(c)].width = 24
        elif c == 8:
            dst_ws.column_dimensions[get_column_letter(c)].width = 14
        elif c in (9, 10):
            dst_ws.column_dimensions[get_column_letter(c)].width = 16
        elif c == 11:
            dst_ws.column_dimensions[get_column_letter(c)].width = 16
        elif c in (12, 13):
            dst_ws.column_dimensions[get_column_letter(c)].width = 14


def process_file(source_path: Path) -> Path:
    wb = openpyxl.load_workbook(source_path)
    if SOURCE_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Foglio '{SOURCE_SHEET_NAME}' non trovato nel file di input.")

    src_ws = wb[SOURCE_SHEET_NAME]

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    out_wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)

    detail_ws = out_wb.create_sheet(DETAIL_SHEET_NAME)
    summary_ws = out_wb.create_sheet(SUMMARY_SHEET_NAME)

    build_detail_sheet(src_ws, detail_ws)
    build_summary_sheet(src_ws, summary_ws)

    output_name = f"{source_path.stem}_elaborato.xlsx"
    output_path = output_dir() / output_name
    out_wb.save(output_path)
    return output_path


def choose_source_file() -> Optional[Path]:
    latest = latest_input_file()
    if latest:
        return latest
    if filedialog is None:
        return None

    selected = filedialog.askopenfilename(
        title="Seleziona il file di stampa commesse",
        initialdir=str(input_dir()),
        filetypes=[("Excel", "*.xlsx")],
    )
    if not selected:
        return None
    return Path(selected)


def run_gui() -> None:
    ensure_workspace()

    root = tk.Tk()
    root.title("Report Commesse")
    root.geometry("560x260")
    root.minsize(560, 260)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    main = ttk.Frame(root, padding=18)
    main.pack(fill="both", expand=True)

    title = ttk.Label(main, text="Elaborazione file stampa commesse", font=("Segoe UI", 14, "bold"))
    title.pack(anchor="w")

    info_var = tk.StringVar()
    latest = latest_input_file()
    if latest:
        info_var.set(f"File trovato: {latest.name}")
    else:
        info_var.set("Nessun file .xlsx trovato nella cartella input.")

    ttk.Label(main, textvariable=info_var, wraplength=520).pack(anchor="w", pady=(10, 12))

    output_var = tk.StringVar(value=f"Output: {output_dir()}")
    ttk.Label(main, textvariable=output_var, wraplength=520).pack(anchor="w", pady=(0, 12))

    def on_process() -> None:
        try:
            source = choose_source_file()
            if source is None:
                return

            result = process_file(source)
            messagebox.showinfo(
                "Completato",
                f"File elaborato con successo.\n\nInput: {source.name}\nOutput: {result.name}",
            )
            output_var.set(f"Creato: {result}")
        except Exception as exc:
            messagebox.showerror("Errore", str(exc))

    button_row = ttk.Frame(main)
    button_row.pack(fill="x", pady=(8, 0))
    ttk.Button(button_row, text="Elabora file", command=on_process).pack(side="left")

    def open_input_folder() -> None:
        os.startfile(str(input_dir()))

    def open_output_folder() -> None:
        os.startfile(str(output_dir()))

    ttk.Button(button_row, text="Apri input", command=open_input_folder).pack(side="left", padx=8)
    ttk.Button(button_row, text="Apri output", command=open_output_folder).pack(side="left")

    hint = ttk.Label(
        main,
        text="Il programma legge il file .xlsx piu recente da input, crea il file elaborato in output e lascia il filtro manuale a Excel.",
        wraplength=520,
    )
    hint.pack(anchor="w", pady=(18, 0))

    root.mainloop()


def run_cli(argv: list[str]) -> int:
    ensure_workspace()

    source = None

    args = list(argv)
    if "--input" in args:
        idx = args.index("--input")
        if idx + 1 >= len(args):
            print("Manca il percorso dopo --input")
            return 2
        source = Path(args[idx + 1])

    if source is None:
        source = latest_input_file()

    if source is None or not source.exists():
        print("Nessun file trovato in input.")
        return 1

    result = process_file(source)
    print(f"Creato: {result}")
    return 0


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        return run_cli(sys.argv[2:])

    if tk is None:
        return run_cli(sys.argv[1:])

    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
