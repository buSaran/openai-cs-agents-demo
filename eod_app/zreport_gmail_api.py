"""Utilities for fetching Z-report totals from Gmail messages."""
from __future__ import annotations

import email
import imaplib
import os
import re
import argparse
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Optional

from dotenv import load_dotenv

load_dotenv()

IMAP_SERVER = "imap.gmail.com"
IMAP_USER = os.getenv("GMAIL_USER")
IMAP_PASS = os.getenv("GMAIL_PASS")
SENDER = "vbralert@visualbusinessretail.co.uk"
ENABLE_GMAIL = os.getenv("EOD_ENABLE_GMAIL", "0").lower() in {"1", "true", "yes"}


@dataclass
class POSReport:
    z_report_no: Optional[str] = None
    report_date: Optional[str] = None
    cash: float = 0.0
    card: float = 0.0
    lottery_till: float = 0.0
    paypoint_till: float = 0.0
    payzone_till: float = 0.0
    inst_po: float = 0.0
    lottery_po: float = 0.0


_NUMBER_RE = r"([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)"


def _grab_num(pattern: str, text: str) -> float:
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return 0.0
    raw = match.group(1).replace(",", "")
    try:
        return float(raw)
    except ValueError:
        return 0.0


def _grab_str(pattern: str, text: str) -> Optional[str]:
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def _parse_body(body: str) -> POSReport:
    report = POSReport()
    report.z_report_no = _grab_str(r"Z-REPORT ID:\s*([A-Za-z0-9-]+)", body)
    report.report_date = _grab_str(r"Date:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", body)
    report.cash = _grab_num(r"CASH\s+" + _NUMBER_RE, body)
    report.card = _grab_num(r"CARD\s+" + _NUMBER_RE, body)
    report.lottery_till = _grab_num(r"LOTTERY\s+" + _NUMBER_RE, body)
    report.paypoint_till = _grab_num(r"PAY\s*POINT\s+" + _NUMBER_RE, body)
    report.payzone_till = _grab_num(r"(PAY\s*ZONE|PLAY\s*STATION)\s+" + _NUMBER_RE, body)
    report.inst_po = _grab_num(r"INSTANT PAYOUTS?\s+" + _NUMBER_RE, body)
    report.lottery_po = _grab_num(r"LOTTERY PAYOUTS?\s+" + _NUMBER_RE, body)
    return report


def _fetch_mailbox() -> Optional[imaplib.IMAP4_SSL]:
    if not ENABLE_GMAIL:
        print("ℹ️ Gmail fetch disabled (set EOD_ENABLE_GMAIL=1 to enable)")
        return None

    if not IMAP_USER or not IMAP_PASS:
        print("❌ Missing Gmail credentials (set GMAIL_USER and GMAIL_PASS)")
        return None

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(IMAP_USER, IMAP_PASS)
        return mail
    except imaplib.IMAP4.error as exc:  # pragma: no cover - defensive logging
        print(
            "❌ Gmail login rejected. Enable IMAP and create an App Password ("
            "https://support.google.com/mail/answer/185833) before setting EOD_ENABLE_GMAIL=1."
        )
        print(f"   Details: {exc}")
        return None
    except Exception as exc:  # pragma: no cover - defensive logging
        print(f"❌ Gmail login failed: {exc}")
        return None


def _iter_messages(shop: str):
    mailbox = _fetch_mailbox()
    if mailbox is None:
        return

    try:
        mailbox.select("inbox")
        status, data = mailbox.search(None, f'(FROM "{SENDER}")')
        if status != "OK":
            print("❌ Could not search inbox for Z-reports")
            return

        for msg_id in reversed(data[0].split()):
            status, msg_data = mailbox.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue
            raw_msg = msg_data[0][1]
            message = email.message_from_bytes(raw_msg)

            body = ""
            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_type() in {"text/plain", "text/html"}:
                        payload = part.get_payload(decode=True)
                        if payload:
                            body += payload.decode(errors="ignore")
            else:
                payload = message.get_payload(decode=True)
                if payload:
                    body = payload.decode(errors="ignore")

            upper = body.upper()
            if shop == "Randell's" and "RANDELL" not in upper:
                continue
            if shop == "Best One" and "BESTONE" not in upper:
                continue

            yield body
    finally:
        try:
            mailbox.logout()
        except Exception:  # pragma: no cover - defensive cleanup
            pass


def fetch_latest_zreport(shop: str) -> Optional[Dict[str, object]]:
    """Return a consolidated Z-report for ``shop``.

    The Gmail feed contains individual POS (till) summaries.  We merge the
    most recent POS1 and POS2 records for the requested shop into a single
    dictionary that matches what the FastAPI layer expects.  ``None`` is
    returned if no qualifying message is found.
    """

    pos1: Optional[POSReport] = None
    pos2: Optional[POSReport] = None
    latest_date: Optional[str] = None

    for body in _iter_messages(shop):
        report = _parse_body(body)
        pos_id = _grab_num(r"POS ID\s*:\s*" + _NUMBER_RE, body)
        latest_date = latest_date or report.report_date

        try:
            pos_index = int(pos_id)
        except (TypeError, ValueError):
            pos_index = 0

        if pos_index == 1 and pos1 is None:
            pos1 = report
        elif pos_index == 2 and pos2 is None:
            pos2 = report

        if pos1 and pos2:
            break

    if not pos1 and not pos2:
        print(f"❌ No Z-report emails found for {shop}")
        return None

    def combine(getattr_name: str) -> float:
        total = 0.0
        for pos in (pos1, pos2):
            if pos is not None:
                total += getattr(pos, getattr_name)
        return total

    date_value = (
        (pos1.report_date if pos1 else None)
        or (pos2.report_date if pos2 else None)
        or latest_date
        or datetime.today().strftime("%d/%m/%Y")
    )

    return {
        "date": date_value,
        "z_report_no": ", ".join(
            [value for value in (pos1.z_report_no if pos1 else None, pos2.z_report_no if pos2 else None) if value]
        ),
        "till_cash": combine("cash"),
        "card": combine("card"),
        "lottery_till": combine("lottery_till"),
        "paypoint_till": combine("paypoint_till"),
        "payzone_till": combine("payzone_till"),
        "inst_po": combine("inst_po"),
        "lottery_po": combine("lottery_po"),
    }


def _format_report(report: Optional[Dict[str, object]]) -> str:
    if not report:
        return "(no data)"

    def _to_float(value: object) -> float:
        try:
            return float(value or 0)
        except (TypeError, ValueError):
            return 0.0

    lines = [
        f"Date: {report.get('date', 'n/a')}",
        f"Z-Report No: {report.get('z_report_no', 'n/a') or 'n/a'}",
        f"Till Cash: {_to_float(report.get('till_cash')):.2f}",
        f"Card: {_to_float(report.get('card')):.2f}",
        f"Lottery Till: {_to_float(report.get('lottery_till')):.2f}",
        f"PayPoint Till: {_to_float(report.get('paypoint_till')):.2f}",
        f"PayZone Till: {_to_float(report.get('payzone_till')):.2f}",
        f"Instant Lotto Payouts: {_to_float(report.get('inst_po')):.2f}",
        f"Lottery Payouts: {_to_float(report.get('lottery_po')):.2f}",
    ]
    return "\n".join(lines)


def _cli(shop: str) -> None:
    print(f"🔍 Fetching latest Z-report for {shop}...")
    report = fetch_latest_zreport(shop)
    if report:
        print("✅ Success! Totals:")
        print(_format_report(report))
    else:
        print("⚠️ No report data returned. Double-check forwarding rules and credentials.")


def main(argv: Optional[list[str]] = None) -> None:
    parser = argparse.ArgumentParser(description="Test the Gmail Z-report fetcher")
    parser.add_argument(
        "--shop",
        choices=["Randell's", "Best One"],
        default="Randell's",
        help="Which shop to fetch (default: Randell's)",
    )
    args = parser.parse_args(argv)
    _cli(args.shop)


if __name__ == "__main__":  # pragma: no cover - manual testing hook
    main()


__all__ = ["fetch_latest_zreport"]
