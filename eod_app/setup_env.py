"""Helper CLI for creating the ``.env`` file used by the EOD app."""
from __future__ import annotations

import argparse
import sys
from getpass import getpass
from pathlib import Path
from typing import Optional


ENV_PATH = Path(__file__).resolve().parent / ".env"


def _prompt(label: str, default: Optional[str] = None, secret: bool = False) -> str:
    """Prompt the user for a value, optionally hiding the input."""

    prompt = f"{label}"
    if default:
        prompt += f" [{default}]"
    prompt += ": "

    if secret:
        value = getpass(prompt)
    else:
        value = input(prompt)

    value = value.strip()
    if not value and default is not None:
        return default
    return value


def _confirm_overwrite(force: bool) -> bool:
    if force or not ENV_PATH.exists():
        return True

    answer = input(f"⚠️ {ENV_PATH.name} already exists. Overwrite it? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def build_env_contents(gmail_user: str, app_password: str, enable_gmail: bool) -> str:
    lines = [
        f"EOD_ENABLE_GMAIL={'1' if enable_gmail else '0'}",
        f"GMAIL_USER={gmail_user}",
        f"GMAIL_PASS={app_password}",
        "",
        "# Never commit this file to source control.",
    ]
    return "\n".join(lines)


def write_env_file(contents: str) -> None:
    ENV_PATH.write_text(contents, encoding="utf-8")


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Create or update the eod_app .env file")
    parser.add_argument("--gmail-user", help="Gmail address that receives the POS alerts")
    parser.add_argument("--app-password", help="16-character Gmail App Password for the account")
    parser.add_argument(
        "--enable-gmail",
        choices=["1", "0", "true", "false", "yes", "no"],
        default=None,
        help="Whether to enable Gmail ingestion in the generated file",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite an existing .env without prompting",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Fail instead of prompting when required values are missing",
    )
    args = parser.parse_args(argv)

    if not _confirm_overwrite(args.force):
        print("Aborted. Existing .env was left untouched.")
        return 1

    gmail_user = (args.gmail_user or "").strip()
    app_password = (args.app_password or "").strip()

    if not gmail_user:
        if args.non_interactive:
            parser.error("--gmail-user is required when --non-interactive is set")
        gmail_user = _prompt("Gmail address for POS emails")

    if not app_password:
        if args.non_interactive:
            parser.error("--app-password is required when --non-interactive is set")
        app_password = _prompt("Paste the 16-character App Password", secret=True)

    enable_gmail: Optional[bool]
    if args.enable_gmail is None:
        if args.non_interactive:
            enable_gmail = True
        else:
            answer = _prompt("Enable Gmail fetching?", default="Y").lower()
            enable_gmail = answer in {"y", "yes", "1", "true"}
    else:
        enable_gmail = args.enable_gmail.lower() in {"1", "true", "yes"}

    contents = build_env_contents(gmail_user, app_password, enable_gmail)
    write_env_file(contents)

    print(f"✅ Wrote credentials to {ENV_PATH}")
    print("   Remember to keep this file private and never commit it to Git.")
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())


__all__ = ["build_env_contents", "write_env_file", "main"]
