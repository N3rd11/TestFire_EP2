"""Resumen de los movimientos validados para el correo institucional."""
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

REPORT = "Analisis_Valores_EP3.xlsx"
MOVEMENT_ATTACHMENT = "Detalle_Movimientos_Altoro_Mutual.xlsx"
UNREPORTED_CURRENCY = "No informada"
AMOUNT_FORMAT = "#,##0.00;[Red]-#,##0.00"


def currency_label(movement):
    return (movement.get("currency") or "").strip() or UNREPORTED_CURRENCY


def operation_label(movement):
    return " ".join(movement["description"].split())


def summarize_operations(movements):
    operations = defaultdict(lambda: [0, Decimal("0")])
    for movement in movements:
        group = operations[operation_label(movement), currency_label(movement)]
        group[0] += 1
        group[1] += Decimal(movement["amount"])
    return [(action, currency, count, total)
            for (action, currency), (count, total) in sorted(operations.items())]


def summarize(movements):
    groups = defaultdict(lambda: [0, Decimal("0"), Decimal("0")])
    for movement in movements:
        group = groups[movement["account"], currency_label(movement)]
        amount = Decimal(movement["amount"])
        group[0] += 1
        if amount >= 0:
            group[1] += amount
        else:
            group[2] -= amount
    return [(account, currency, count, credits, debits, credits - debits)
            for (account, currency), (count, credits, debits) in sorted(groups.items())]


def write_analysis(folder, origin, start, end, movements, groups, classifications):
    book = Workbook()
    by_type = book.active
    by_type.title = "Por tipo"
    by_type.append(("Operación original", "Moneda", "Cantidad", "Importe acumulado"))
    for row in summarize_operations(movements):
        by_type.append(row)
        action, currency, _, total = by_type[by_type.max_row]
        action.data_type = currency.data_type = "s"
        total.number_format = AMOUNT_FORMAT

    detail = book.create_sheet("Detalle de operaciones")
    detail.append(("Fecha y hora", "Cuenta", "ID de transacción", "Operación original", "Importe", "Moneda"))
    for movement in movements:
        stamp = movement.get("date_raw") or movement["date"]
        detail.append((datetime.fromisoformat(stamp), movement["account"], movement.get("web_id"),
                       operation_label(movement), Decimal(movement["amount"]), currency_label(movement)))
        date, account, transaction, action, amount, currency = detail[detail.max_row]
        for cell in (account, transaction, action, currency):
            if isinstance(cell.value, str):
                cell.data_type = "s"
        date.number_format = "yyyy-mm-dd hh:mm:ss" if len(stamp) > 10 else "yyyy-mm-dd"
        amount.number_format = AMOUNT_FORMAT

    summary = book.create_sheet("Resumen")
    for row in (("Campo", "Valor"), ("Ejecución", folder.name), ("Origen", origin),
                ("Periodo desde", start), ("Periodo hasta", end),
                ("Movimiento neto", "Abonos menos cargos dentro del periodo"),
                *sorted(classifications.items())):
        summary.append(row)

    accounts = book.create_sheet("Por cuenta")
    accounts.append(("Cuenta", "Moneda", "Movimientos", "Abonos", "Cargos", "Movimiento neto"))
    for row in groups:
        accounts.append(row)
        cells = accounts[accounts.max_row]
        for cell in cells[:2]:
            cell.data_type = "s"
        for cell in cells[3:]:
            cell.number_format = AMOUNT_FORMAT
    for sheet, widths in (
        (by_type, (28, 19, 12, 24)),
        (detail, (24, 16, 22, 30, 21, 19)),
        (summary, (25, 50)),
        (accounts, (19, 21, 17, 19, 19, 22)),
    ):
        sheet.freeze_panes = "A2"
        for cell, width in zip(sheet[1], widths):
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="16324F")
            sheet.column_dimensions[cell.column_letter].width = width
    accounts.auto_filter.ref = accounts.dimensions
    by_type.auto_filter.ref = by_type.dimensions
    detail.auto_filter.ref = detail.dimensions

    target = folder / REPORT
    temporary = target.with_suffix(".tmp.xlsx")
    try:
        book.save(temporary)
        saved = load_workbook(temporary, read_only=True, data_only=True)
        try:
            if (saved["Resumen"]["B2"].value != folder.name
                    or saved["Por cuenta"].max_row != len(groups) + 1
                    or saved["Detalle de operaciones"].max_row != len(movements) + 1
                    or sum(row[2] for row in saved["Por tipo"].iter_rows(min_row=2, values_only=True)) != len(movements)):
                raise ValueError("El Excel de análisis guardado no coincide con la ejecución")
        finally:
            saved.close()
        temporary.replace(target)
    finally:
        book.close()
        temporary.unlink(missing_ok=True)
    return target


def write_movement_attachment(analysis, movement_count):
    target = analysis.with_name(MOVEMENT_ATTACHMENT)
    temporary = target.with_suffix(".tmp.xlsx")
    book = load_workbook(analysis)
    try:
        detail = book["Detalle de operaciones"]
        if (book["Resumen"]["B2"].value != analysis.parent.name
                or detail.max_row != movement_count + 1):
            raise ValueError("El detalle de movimientos no corresponde a esta ejecución")
        for sheet in list(book.worksheets):
            if sheet != detail:
                book.remove(sheet)
        book.active = 0
        detail.title = "Movimientos"
        detail["C1"] = "ID"
        detail["D1"] = "Operación"
        detail.auto_filter.ref = None
        for (cell,) in detail.iter_rows(min_row=2, min_col=1, max_col=1):
            if "hh:mm" in cell.number_format and not (cell.value.second or cell.value.microsecond):
                cell.number_format = "yyyy-mm-dd hh:mm"
        table = Table(displayName="TablaMovimientos", ref=detail.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        detail.add_table(table)
        book.save(temporary)
        temporary.replace(target)
    finally:
        book.close()
        temporary.unlink(missing_ok=True)
    return target


def email_summary(origin, start, end, movements, rejected_rows, reference):
    source = {
        "ARCHIVO_LOCAL_HISTORICO": "Respaldo histórico de Altoro Mutual",
        "API_ARCHIVO_HISTORICO": "Respaldo histórico de Altoro Mutual",
        "WEB": "Altoro Mutual",
        "API_OPERATIVA": "API operativa de movimientos",
    }.get(origin, origin)
    operations = Counter(operation_label(row).casefold() for row in movements)
    deposits = operations.pop("deposit", 0)
    withdrawals = operations.pop("withdrawal", 0)
    other = sum(operations.values())
    breakdown = f"{deposits} depósitos · {withdrawals} retiros"
    if other:
        breakdown += f" · {other} de otro tipo"
    year, value = reference[0][0]
    inflation = f"{value:.2f}".replace(".", ",")
    return "\n".join([
        f"Reporte EP3 · {source}",
        f"Período {start} al {end}",
        f"{len(movements)} movimientos válidos · {rejected_rows} rechazados",
        breakdown,
        f"Banco Mundial · Inflación anual de Chile en {year}: {inflation} %",
        "",
        "Adjuntos",
        f"- {MOVEMENT_ATTACHMENT} (tabla filtrable de movimientos)",
        "- Reporte_TestFire_EP3.xlsx (reporte y referencia del Banco Mundial)",
        f"- {REPORT} (totales por cuenta y tipo)",
    ])
