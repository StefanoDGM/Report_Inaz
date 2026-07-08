import copy
import datetime as dt
import os
import subprocess
import re
import sys
import tempfile
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
ERRORS_SHEET_NAME = "Probabili errori"
DETAIL_OUTPUT_NAME = "ore_analitica.xls"
SUMMARY_OUTPUT_SUFFIX = "_riepilogo"

TOTAL_PROJ_MARKER = "Totale"
TRAVEL_KEYWORDS = ("COMMESSA", "CHIUSURA")
SUMMARY_SOURCE_COLUMNS = [
    (7, "Reparto"),
    (14, "Codice dipendente"),
    (15, "Nominativo"),
    (16, "Data"),
]


@dataclass
class GroupSummary:
    order: int
    first_row_idx: int
    base_row: list[Any]
    total_hours: float
    office_hours: float
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

    @property
    def office_ratio(self) -> float:
        if self.total_hours <= 0:
            return 0.0
        return self.office_hours / self.total_hours


@dataclass
class ProcessingResult:
    detail_path: Path
    summary_path: Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def input_dir() -> Path:
    return app_root() / "input"


def output_dir() -> Path:
    return app_root() / "output"


def shell_quote(value: Path | str) -> str:
    text = str(value)
    return "'" + text.replace("'", "''") + "'"


def convert_xlsx_to_xls(src_xlsx: Path, dst_xls: Path) -> None:
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$src = {shell_quote(src_xlsx)}
$dst = {shell_quote(dst_xls)}
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($src)
$xlExcel5 = 39
$workbook.SaveAs($dst, $xlExcel5)
$workbook.Close($false)
$excel.Quit()
"""

    completed = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "Impossibile convertire il file in formato .xls tramite Excel. "
            f"Dettagli: {completed.stderr.strip() or completed.stdout.strip() or 'errore sconosciuto'}"
        )


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


def is_blank_row(row: list[Any]) -> bool:
    return all(str(cell or "").strip() == "" for cell in row)


def travel_row(row: list[Any]) -> bool:
    for idx in (8, 9, 10, 11):
        if idx < len(row):
            cell = row[idx]
            if isinstance(cell, str) and any(keyword in cell.upper() for keyword in TRAVEL_KEYWORDS):
                return True
    return False


def is_chiusura_row(row: list[Any]) -> bool:
    for idx in (8, 9, 10, 11):
        if idx < len(row):
            cell = row[idx]
            if isinstance(cell, str) and "CHIUSURA" in cell.upper():
                return True
    return False


def is_generic_commessa_row(row: list[Any]) -> bool:
    if len(row) <= 9:
        return False
    return normalize_text(row[9]) == "commessa"


def is_presence_row(row: list[Any]) -> bool:
    project = normalize_text(row[9]) if len(row) > 9 else ""
    cod_project = normalize_text(row[8]) if len(row) > 8 else ""
    cod_argomento = normalize_text(row[10]) if len(row) > 10 else ""
    argomento = normalize_text(row[11]) if len(row) > 11 else ""
    return not any((project, cod_project, cod_argomento, argomento))


def row_project_name(row: list[Any]) -> str:
    if len(row) <= 9:
        return ""
    return normalize_text(row[9])


def row_argument_name(row: list[Any]) -> str:
    if len(row) <= 11:
        return ""
    return normalize_text(row[11])


def is_office_row(row: list[Any]) -> bool:
    if len(row) <= 9:
        return False
    project_name = str(row[9] or "").strip().casefold()
    return "sede ufficio" in project_name


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


def row_quantity(value: Any) -> float:
    parsed = parse_duration(value)
    return round(parsed or 0.0, 5)


def distribute_amount(rows: list[dict[str, Any]], amount: float) -> float:
    if amount <= 0 or not rows:
        return amount

    total = sum(max(row["quantity"], 0.0) for row in rows)
    if total <= 0:
        return amount

    deducted = min(amount, total)
    for row in rows:
        share = deducted * (max(row["quantity"], 0.0) / total)
        row["quantity"] = round(row["quantity"] - share, 5)

    return round(amount - deducted, 5)


def adjust_lunch_break_for_group(rows: list[dict[str, Any]]) -> None:
    manutentori_rows = [row for row in rows if row["is_manutentori"] and row["quantity"] > 0]
    if not manutentori_rows:
        return

    remaining = 1.0
    priority_buckets = [
        [row for row in manutentori_rows if row["bucket"] == "generic_commessa"],
        [row for row in manutentori_rows if row["bucket"] == "chiusura"],
        [row for row in manutentori_rows if row["bucket"] == "other_commessa"],
    ]

    for bucket_rows in priority_buckets:
        if remaining <= 0:
            break
        remaining = distribute_amount(bucket_rows, remaining)

    if remaining > 0:
        remaining = distribute_amount(manutentori_rows, remaining)

    for row in manutentori_rows:
        row["quantity"] = round(row["quantity"], 5)


def collect_group_summaries(src_ws) -> OrderedDict[tuple[Any, Any, Any], dict[str, Any]]:
    grouped: OrderedDict[tuple[Any, Any, Any], dict[str, Any]] = OrderedDict()

    for row_idx in range(2, src_ws.max_row + 1):
        values = [src_ws.cell(row_idx, c).value for c in range(1, src_ws.max_column + 1)]
        if is_blank_row(values) or is_total_row(values):
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
                "first_row_idx": row_idx,
                "rows": [],
                "total_hours": 0.0,
                "office_hours": 0.0,
                "travel_gross_hours": 0.0,
            }
            info = grouped[key]

        quantity = parse_duration(values[17]) or 0.0
        row_info = {
            "row_idx": row_idx,
            "values": values,
            "quantity": quantity,
            "is_manutentori": str(values[7] or "").strip().upper() == "MANUTENTORI",
            "bucket": "other_commessa",
        }

        if is_generic_commessa_row(values):
            row_info["bucket"] = "generic_commessa"
        elif is_chiusura_row(values):
            row_info["bucket"] = "chiusura"

        info["rows"].append(row_info)
        info["total_hours"] += quantity
        if row_info["is_manutentori"] and is_office_row(values):
            info["office_hours"] += quantity
        if row_info["is_manutentori"] and travel_row(values):
            info["travel_gross_hours"] += quantity

    return grouped


def build_detail_sheet(src_ws, dst_ws) -> dict[int, int]:
    clone_sheet_layout(src_ws, dst_ws)

    grouped = collect_group_summaries(src_ws)
    source_to_output_row: dict[int, int] = {}

    for c in range(1, src_ws.max_column + 1):
        src_cell = src_ws.cell(1, c)
        dst_cell = dst_ws.cell(1, c)
        dst_cell.value = src_cell.value
        copy_cell_style(src_cell, dst_cell)

    output_row = 2
    for info in grouped.values():
        detail_rows = [dict(row) for row in info["rows"]]
        adjust_lunch_break_for_group(detail_rows)

        for row in detail_rows:
            source_row_idx = row["row_idx"]
            source_to_output_row[source_row_idx] = output_row
            for c in range(1, src_ws.max_column + 1):
                src_cell = src_ws.cell(source_row_idx, c)
                dst_cell = dst_ws.cell(output_row, c)
                dst_cell.value = src_cell.value
                copy_cell_style(src_cell, dst_cell)

                if c == 18:
                    dst_cell.value = round(row["quantity"], 5)
                    dst_cell.number_format = "0.00000"

            output_row += 1

    for c in range(1, src_ws.max_column + 1):
        dst_ws.column_dimensions[get_column_letter(c)].width = src_ws.column_dimensions[get_column_letter(c)].width

    last_col = get_column_letter(src_ws.max_column)
    dst_ws.auto_filter.ref = f"A1:{last_col}{max(output_row - 1, 1)}"
    return source_to_output_row


def build_summary_rows(src_ws) -> list[GroupSummary]:
    grouped = collect_group_summaries(src_ws)

    summaries: list[GroupSummary] = []
    for info in grouped.values():
        base_row = list(info["rows"][0]["values"])
        gross_travel = round(info["travel_gross_hours"], 5)
        is_manutentori = str(base_row[7] or "").strip().upper() == "MANUTENTORI"
        if is_manutentori:
            travel_net_hours = round(gross_travel - 1.0, 5)
        else:
            travel_net_hours = 0.0

        summaries.append(
            GroupSummary(
                order=info["order"],
                first_row_idx=info["first_row_idx"],
                base_row=base_row,
                total_hours=round(info["total_hours"], 5),
                office_hours=round(info["office_hours"], 5) if is_manutentori else 0.0,
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
        "Ore sede ufficio",
        "Ore viaggio lorde",
        "Ore viaggio nette",
        "% sede ufficio",
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
                summary.office_hours,
                summary.travel_gross_hours,
                summary.travel_net_hours,
                round(summary.office_ratio, 4),
                round(summary.gross_ratio, 4),
                round(summary.net_ratio, 4),
            ]
        )

        for c, value in enumerate(values, start=1):
            cell = dst_ws.cell(row_idx, c)
            cell.value = value
            if c in (5, 6, 7, 8):
                cell.number_format = "0.00000"
            elif c in (9, 10, 11):
                cell.number_format = "0.0000%"
            elif c == 4:
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
            dst_ws.column_dimensions[get_column_letter(c)].width = 24
        elif c == 4:
            dst_ws.column_dimensions[get_column_letter(c)].width = 14
        elif c in (5, 6, 7, 8):
            dst_ws.column_dimensions[get_column_letter(c)].width = 16
        elif c in (9, 10, 11):
            dst_ws.column_dimensions[get_column_letter(c)].width = 14


def build_errors_sheet(src_ws, dst_ws, detail_row_map: dict[int, int]) -> None:
    clone_sheet_layout(src_ws, dst_ws)

    grouped = collect_group_summaries(src_ws)
    error_rows: list[dict[str, Any]] = []

    for info in grouped.values():
        first_row_idx = info["first_row_idx"]
        first_row_values = info["rows"][0]["values"]
        day_label = first_row_values[16]
        employee_name = first_row_values[15]
        total_hours = info["total_hours"]
        travel_gross_hours = info["travel_gross_hours"]

        if total_hours > 0 and abs(travel_gross_hours - total_hours) <= 0.00001 and travel_gross_hours > 0:
            error_rows.append(
                {
                    "row_idx": detail_row_map.get(first_row_idx, first_row_idx),
                    "data": day_label,
                    "nominativo": employee_name,
                    "errore": "ore viaggio 100%",
                }
            )

        for row in info["rows"]:
            values = row["values"]
            project_name = row_project_name(values)
            argument_name = row_argument_name(values)
            if project_name and project_name != "commessa" and not argument_name:
                error_rows.append(
                    {
                        "row_idx": detail_row_map.get(row["row_idx"], row["row_idx"]),
                        "data": values[16],
                        "nominativo": values[15],
                        "errore": f'manca argomento per "{values[9]}"',
                    }
                )

    error_rows.sort(key=lambda item: (item["row_idx"], item["errore"]))

    title = dst_ws["A1"]
    title.value = "Probabili errori"
    title.font = Font(bold=True, size=14)

    headers = ["Numero riga", "Data", "Nominativo", "Errore"]
    for c, header in enumerate(headers, start=1):
        cell = dst_ws.cell(4, c)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F8D7DA")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        thin = Side(style="thin", color="999999")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, item in enumerate(error_rows, start=5):
        dst_ws.cell(row_idx, 1).value = item["row_idx"]
        dst_ws.cell(row_idx, 2).value = item["data"]
        dst_ws.cell(row_idx, 3).value = item["nominativo"]
        dst_ws.cell(row_idx, 4).value = item["errore"]
        dst_ws.cell(row_idx, 2).number_format = "dd/mm/yyyy"

    dst_ws.freeze_panes = "A5"
    dst_ws.auto_filter.ref = f"A4:D{max(len(error_rows) + 4, 4)}"
    dst_ws.column_dimensions["A"].width = 14
    dst_ws.column_dimensions["B"].width = 14
    dst_ws.column_dimensions["C"].width = 24
    dst_ws.column_dimensions["D"].width = 48


def process_file(source_path: Path) -> ProcessingResult:
    wb = openpyxl.load_workbook(source_path)
    if SOURCE_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Foglio '{SOURCE_SHEET_NAME}' non trovato nel file di input.")

    src_ws = wb[SOURCE_SHEET_NAME]

    detail_wb = Workbook()
    detail_default = detail_wb.active
    detail_wb.remove(detail_default)
    detail_wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    detail_ws = detail_wb.create_sheet(DETAIL_SHEET_NAME)
    detail_row_map = build_detail_sheet(src_ws, detail_ws)

    summary_wb = Workbook()
    summary_default = summary_wb.active
    summary_wb.remove(summary_default)
    summary_wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    summary_ws = summary_wb.create_sheet(SUMMARY_SHEET_NAME)
    errors_ws = summary_wb.create_sheet(ERRORS_SHEET_NAME)
    build_summary_sheet(src_ws, summary_ws)
    build_errors_sheet(src_ws, errors_ws, detail_row_map)

    detail_path = output_dir() / DETAIL_OUTPUT_NAME
    summary_path = output_dir() / f"{source_path.stem}{SUMMARY_OUTPUT_SUFFIX}.xlsx"
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_xlsx = Path(tmpdir) / f"{source_path.stem}_ore_analitica_temp.xlsx"
        detail_wb.save(temp_xlsx)
        convert_xlsx_to_xls(temp_xlsx, detail_path)
    summary_wb.save(summary_path)
    return ProcessingResult(detail_path=detail_path, summary_path=summary_path)


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
                (
                    "File elaborato con successo.\n\n"
                    f"Input: {source.name}\n"
                    f"Dettaglio: {result.detail_path.name}\n"
                    f"Riepilogo: {result.summary_path.name}"
                ),
            )
            output_var.set(f"Creati: {result.detail_path.name} e {result.summary_path.name}")
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
        text="Il programma legge il file .xlsx piu recente da input, crea due file distinti in output e lascia il filtro manuale a Excel.",
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
    print(f"Creati: {result.detail_path} | {result.summary_path}")
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
