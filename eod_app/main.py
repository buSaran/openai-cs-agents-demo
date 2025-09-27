"""FastAPI application for end-of-day (EOD) reporting.

This module provides two form workflows – one for Randell's and one for
Best One – that persist submissions to Excel workbooks.  It integrates
with a Gmail-powered helper to retrieve the latest Z-report figures that
pre-populate the forms and are used to calculate the final variances
stored in the spreadsheet.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Sequence, Set

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook, load_workbook

from .zreport_gmail_api import fetch_latest_zreport

APP_ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = APP_ROOT / "templates"

app = FastAPI(title="End of Day Reporter")
templates = Jinja2Templates(directory=str(TEMPLATE_DIR))

FILES = {
    "randells": APP_ROOT / "randells_reports.xlsx",
    "bestone": APP_ROOT / "bestone_reports.xlsx",
}


@dataclass(frozen=True)
class WorkbookStyle:
    headers: Sequence[str]
    column_colours: Mapping[str, str]
    variance_headers: Set[str]
    decimal_headers: Set[str]


PALE_YELLOW = "FFFCE9A1"
PALE_PINK = "FFF8D9E6"
PALE_RED = "FFF8B4B4"


RANDELL_HEADERS: Sequence[str] = (
    "Date",
    "Z-Report No",
    "Adjusted Cash",
    "Counted Cash",
    "Cash Variance",
    "Lottery Till",
    "Lottery Report",
    "Lottery Variance",
    "Lottery Machine",
    "Lottery Payout Variance",
    "PayPoint Till",
    "PayPoint Report",
    "PayPoint Variance",
    "PayPoint Start Tx",
    "PayPoint End Tx",
    "PayZone Till",
    "PayZone Report",
    "PayZone Variance",
    "PayZone Start Tx",
    "PayZone End Tx",
    "Payouts",
    "Card",
    "Notes",
    "Staff 1 Name",
    "Staff 1 Start",
    "Staff 1 End",
    "Staff 2 Name",
    "Staff 2 Start",
    "Staff 2 End",
    "Staff 3 Name",
    "Staff 3 Start",
    "Staff 3 End",
    "Staff 4 Name",
    "Staff 4 Start",
    "Staff 4 End",
)


BESTONE_HEADERS: Sequence[str] = (
    "Date",
    "Z-Report No",
    "Adjusted Cash",
    "Counted Cash",
    "Cash Variance",
    "Lottery Till",
    "Lottery Report",
    "Lottery Variance",
    "Lottery Machine",
    "Lottery Payout Variance",
    "PayPoint Till",
    "PayPoint Report",
    "PayPoint Variance",
    "PayPoint Start Tx",
    "PayPoint End Tx",
    "PlayStation Till",
    "PlayStation Report",
    "PlayStation Variance",
    "PlayStation Start Tx",
    "PlayStation End Tx",
    "Payouts",
    "PlayStation Payouts",
    "Card",
    "Notes",
    "Staff 1 Name",
    "Staff 1 Start",
    "Staff 1 End",
    "Staff 2 Name",
    "Staff 2 Start",
    "Staff 2 End",
    "Staff 3 Name",
    "Staff 3 Start",
    "Staff 3 End",
    "Staff 4 Name",
    "Staff 4 Start",
    "Staff 4 End",
)


RANDELL_STYLE = WorkbookStyle(
    headers=RANDELL_HEADERS,
    column_colours={
        "Adjusted Cash": PALE_YELLOW,
        "Counted Cash": PALE_YELLOW,
        "Cash Variance": PALE_YELLOW,
        "Payouts": PALE_YELLOW,
        "Lottery Till": PALE_PINK,
        "Lottery Report": PALE_PINK,
        "Lottery Variance": PALE_PINK,
        "Lottery Machine": PALE_PINK,
        "Lottery Payout Variance": PALE_PINK,
        "PayPoint Till": PALE_YELLOW,
        "PayPoint Report": PALE_YELLOW,
        "PayPoint Variance": PALE_YELLOW,
        "PayPoint Start Tx": PALE_YELLOW,
        "PayPoint End Tx": PALE_YELLOW,
        "PayZone Till": PALE_RED,
        "PayZone Report": PALE_RED,
        "PayZone Variance": PALE_RED,
        "PayZone Start Tx": PALE_RED,
        "PayZone End Tx": PALE_RED,
    },
    variance_headers={
        "Cash Variance",
        "Lottery Variance",
        "Lottery Payout Variance",
        "PayPoint Variance",
        "PayZone Variance",
    },
    decimal_headers={
        "Adjusted Cash",
        "Counted Cash",
        "Cash Variance",
        "Lottery Till",
        "Lottery Report",
        "Lottery Variance",
        "Lottery Machine",
        "Lottery Payout Variance",
        "PayPoint Till",
        "PayPoint Report",
        "PayPoint Variance",
        "PayZone Till",
        "PayZone Report",
        "PayZone Variance",
        "Payouts",
        "Card",
    },
)


BESTONE_STYLE = WorkbookStyle(
    headers=BESTONE_HEADERS,
    column_colours={
        "Adjusted Cash": PALE_YELLOW,
        "Counted Cash": PALE_YELLOW,
        "Cash Variance": PALE_YELLOW,
        "Payouts": PALE_YELLOW,
        "PlayStation Payouts": PALE_YELLOW,
        "Lottery Till": PALE_PINK,
        "Lottery Report": PALE_PINK,
        "Lottery Variance": PALE_PINK,
        "Lottery Machine": PALE_PINK,
        "Lottery Payout Variance": PALE_PINK,
        "PayPoint Till": PALE_YELLOW,
        "PayPoint Report": PALE_YELLOW,
        "PayPoint Variance": PALE_YELLOW,
        "PayPoint Start Tx": PALE_YELLOW,
        "PayPoint End Tx": PALE_YELLOW,
        "PlayStation Till": PALE_RED,
        "PlayStation Report": PALE_RED,
        "PlayStation Variance": PALE_RED,
        "PlayStation Start Tx": PALE_RED,
        "PlayStation End Tx": PALE_RED,
    },
    variance_headers={
        "Cash Variance",
        "Lottery Variance",
        "Lottery Payout Variance",
        "PayPoint Variance",
        "PlayStation Variance",
    },
    decimal_headers={
        "Adjusted Cash",
        "Counted Cash",
        "Cash Variance",
        "Lottery Till",
        "Lottery Report",
        "Lottery Variance",
        "Lottery Machine",
        "Lottery Payout Variance",
        "PayPoint Till",
        "PayPoint Report",
        "PayPoint Variance",
        "PlayStation Till",
        "PlayStation Report",
        "PlayStation Variance",
        "Payouts",
        "PlayStation Payouts",
        "Card",
    },
)


SHOP_STYLES: Mapping[str, WorkbookStyle] = {
    "randells": RANDELL_STYLE,
    "bestone": BESTONE_STYLE,
}


DEFAULT_ZREPORT: Dict[str, Any] = {
    "date": datetime.today().strftime("%d/%m/%Y"),
    "z_report_no": "",
    "till_cash": 0.0,
    "card": 0.0,
    "inst_lotto": 0.0,
    "lottery_till": 0.0,
    "paypoint_till": 0.0,
    "payzone_till": 0.0,
    "lottery_po": 0.0,
    "inst_po": 0.0,
}


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _ensure_workbook(path: Path, style: WorkbookStyle) -> Workbook:
    """Return an openpyxl workbook with headers guaranteed for the EOD sheet."""

    path = Path(path)
    if not path.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "EOD"
        ws.append(list(style.headers))
        wb.save(path)

    wb = load_workbook(path)
    if "EOD" not in wb.sheetnames:
        ws = wb.create_sheet("EOD")
        ws.append(list(style.headers))
    else:
        ws = wb["EOD"]
        if ws.max_row == 1 and ws.max_column == 1 and ws.cell(row=1, column=1).value is None:
            ws.delete_rows(1)
            ws.append(list(style.headers))
        else:
            current = [ws.cell(row=1, column=idx).value for idx in range(1, len(style.headers) + 1)]
            if current != list(style.headers):
                ws.insert_rows(1)
                for idx, header in enumerate(style.headers, start=1):
                    ws.cell(row=1, column=idx, value=header)
    return wb


def _safe_float(value: Optional[float]) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def get_zreport(shop: str) -> Dict[str, Any]:
    """Fetch the latest Z-report data for ``shop`` with sensible defaults."""
    try:
        data = fetch_latest_zreport(shop)
    except Exception as exc:  # pragma: no cover - defensive logging only
        print(f"⚠️ Failed to fetch Z-report for {shop}: {exc}")
        data = None

    if not data:
        return DEFAULT_ZREPORT.copy()

    merged = DEFAULT_ZREPORT.copy()
    merged.update({k: _safe_float(v) if isinstance(v, (int, float)) else v for k, v in data.items()})
    if not merged.get("date"):
        merged["date"] = DEFAULT_ZREPORT["date"]
    return merged


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def home(_: Request) -> HTMLResponse:
    """Landing page that links to the individual shop forms."""
    html = (
        "<html><head><title>EOD Home</title>"
        "<style>body{font-family:Arial;text-align:center;padding-top:100px;}"
        "a.button{display:inline-block;padding:20px 40px;margin:20px;font-size:22px;"
        "font-weight:bold;text-decoration:none;color:white;border-radius:10px;}"
        ".randells{background:#007bff;}.bestone{background:#28a745;}"
        "</style></head><body>"
        "<h1>Choose Shop</h1>"
        "<a class='button randells' href='/randells'>Randell&#39;s EOD</a>"
        "<a class='button bestone' href='/bestone'>Best One EOD</a>"
        "</body></html>"
    )
    return HTMLResponse(content=html)


def _tidy_workbook_for(shop_key: str) -> None:
    from .cleanup_excel import tidy_workbook

    style = SHOP_STYLES[shop_key]
    tidy_workbook(
        FILES[shop_key],
        sheet_name="EOD",
        headers=style.headers,
        column_colours=style.column_colours,
        variance_headers=style.variance_headers,
        decimal_headers=style.decimal_headers,
    )


@app.get("/randells", response_class=HTMLResponse)
async def get_randells_form(request: Request):
    z = get_zreport("Randell's")
    message = (
        f"✅ Z-Report Loaded ({z.get('date', '')})"
        if _safe_float(z.get("till_cash")) > 0
        else "⚠️ No Z-Report loaded"
    )
    return templates.TemplateResponse(
        "form_randells.html",
        {"request": request, "shop": "Randell's", "pref": z, "message": message},
    )


@app.post("/randells", response_class=HTMLResponse)
async def submit_randells(
    pin: str = Form(...),
    payouts: float = Form(...),
    counted_cash: float = Form(...),
    lottery_report: float = Form(...),
    lottery_machine: float = Form(...),
    paypoint_report: float = Form(...),
    pp_start_tx: str = Form(""),
    pp_end_tx: str = Form(""),
    payzone_report: float = Form(...),
    pz_start_tx: str = Form(""),
    pz_end_tx: str = Form(""),
    notes: str = Form(""),
    staff1_name: str = Form(""), staff1_start: str = Form(""), staff1_end: str = Form(""),
    staff2_name: str = Form(""), staff2_start: str = Form(""), staff2_end: str = Form(""),
    staff3_name: str = Form(""), staff3_start: str = Form(""), staff3_end: str = Form(""),
    staff4_name: str = Form(""), staff4_start: str = Form(""), staff4_end: str = Form(""),
):
    del pin  # Only used for validation; never persisted.
    z = get_zreport("Randell's")

    adjusted_cash = _safe_float(z.get("till_cash")) - payouts
    cash_var = adjusted_cash - counted_cash
    lot_var = lottery_report - _safe_float(z.get("lottery_till"))
    lot_payout_var = lottery_machine - (
        abs(_safe_float(z.get("inst_po"))) + abs(_safe_float(z.get("lottery_po")))
    )
    pp_var = paypoint_report - _safe_float(z.get("paypoint_till"))
    pz_var = payzone_report - _safe_float(z.get("payzone_till"))

    style = SHOP_STYLES["randells"]
    wb = _ensure_workbook(FILES["randells"], style)
    ws = wb["EOD"]
    ws.append(
        [
            z.get("date", DEFAULT_ZREPORT["date"]),
            z.get("z_report_no", ""),
            adjusted_cash,
            counted_cash,
            cash_var,
            _safe_float(z.get("lottery_till")),
            lottery_report,
            lot_var,
            lottery_machine,
            lot_payout_var,
            _safe_float(z.get("paypoint_till")),
            paypoint_report,
            pp_var,
            pp_start_tx,
            pp_end_tx,
            _safe_float(z.get("payzone_till")),
            payzone_report,
            pz_var,
            pz_start_tx,
            pz_end_tx,
            payouts,
            _safe_float(z.get("card")),
            notes,
            staff1_name,
            staff1_start,
            staff1_end,
            staff2_name,
            staff2_start,
            staff2_end,
            staff3_name,
            staff3_start,
            staff3_end,
            staff4_name,
            staff4_start,
            staff4_end,
        ]
    )
    wb.save(FILES["randells"])

    _tidy_workbook_for("randells")
    return HTMLResponse("<h3>✅ Report saved and Excel updated for Randell's</h3>")


@app.get("/bestone", response_class=HTMLResponse)
async def get_bestone_form(request: Request):
    z = get_zreport("Best One")
    message = (
        f"✅ Z-Report Loaded ({z.get('date', '')})"
        if _safe_float(z.get("till_cash")) > 0
        else "⚠️ No Z-Report loaded"
    )
    return templates.TemplateResponse(
        "form_bestone.html",
        {"request": request, "shop": "Best One", "pref": z, "message": message},
    )


@app.post("/bestone", response_class=HTMLResponse)
async def submit_bestone(
    pin: str = Form(...),
    payouts: float = Form(...),
    ps_payout: float = Form(0),
    counted_cash: float = Form(...),
    lottery_report: float = Form(...),
    lottery_machine: float = Form(...),
    paypoint_report: float = Form(...),
    pp_start_tx: str = Form(""),
    pp_end_tx: str = Form(""),
    ps_report: float = Form(...),
    ps_start_tx: str = Form(""),
    ps_end_tx: str = Form(""),
    notes: str = Form(""),
    staff1_name: str = Form(""), staff1_start: str = Form(""), staff1_end: str = Form(""),
    staff2_name: str = Form(""), staff2_start: str = Form(""), staff2_end: str = Form(""),
    staff3_name: str = Form(""), staff3_start: str = Form(""), staff3_end: str = Form(""),
    staff4_name: str = Form(""), staff4_start: str = Form(""), staff4_end: str = Form(""),
):
    del pin
    z = get_zreport("Best One")

    adjusted_cash = _safe_float(z.get("till_cash")) - (payouts + ps_payout)
    cash_var = adjusted_cash - counted_cash
    lot_var = lottery_report - _safe_float(z.get("lottery_till"))
    lot_payout_var = lottery_machine - (
        abs(_safe_float(z.get("inst_po"))) + abs(_safe_float(z.get("lottery_po")))
    )
    pp_var = paypoint_report - _safe_float(z.get("paypoint_till"))
    ps_var = ps_report - _safe_float(z.get("payzone_till"))

    style = SHOP_STYLES["bestone"]
    wb = _ensure_workbook(FILES["bestone"], style)
    ws = wb["EOD"]
    ws.append(
        [
            z.get("date", DEFAULT_ZREPORT["date"]),
            z.get("z_report_no", ""),
            adjusted_cash,
            counted_cash,
            cash_var,
            _safe_float(z.get("lottery_till")),
            lottery_report,
            lot_var,
            lottery_machine,
            lot_payout_var,
            _safe_float(z.get("paypoint_till")),
            paypoint_report,
            pp_var,
            pp_start_tx,
            pp_end_tx,
            _safe_float(z.get("payzone_till")),
            ps_report,
            ps_var,
            ps_start_tx,
            ps_end_tx,
            payouts,
            ps_payout,
            _safe_float(z.get("card")),
            notes,
            staff1_name,
            staff1_start,
            staff1_end,
            staff2_name,
            staff2_start,
            staff2_end,
            staff3_name,
            staff3_start,
            staff3_end,
            staff4_name,
            staff4_start,
            staff4_end,
        ]
    )
    wb.save(FILES["bestone"])

    _tidy_workbook_for("bestone")
    return HTMLResponse("<h3>✅ Report saved and Excel updated for Best One</h3>")


if __name__ == "__main__":  # pragma: no cover
    import uvicorn

    uvicorn.run("eod_app.main:app", host="0.0.0.0", port=8000, reload=True)
