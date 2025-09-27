"""Utility to launch the FastAPI backend on demand."""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
from pathlib import Path
DEFAULT_HOST = os.getenv("EOD_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.getenv("EOD_PORT", "8000"))


def _is_server_live(host: str, port: int) -> bool:
    try:
        with socket.create_connection((host, port), timeout=1):
            return True
    except OSError:
        return False


def ensure_backend(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT, timeout: float = 15.0) -> bool:
    """Start the FastAPI backend if it is not already running."""

    if _is_server_live(host, port):
        return True

    repo_root = Path(__file__).resolve().parent.parent
    command = [
        sys.executable,
        "-m",
        "uvicorn",
        "eod_app.main:app",
        "--host",
        host,
        "--port",
        str(port),
    ]

    popen_kwargs: dict[str, object] = {"cwd": str(repo_root)}
    if os.name == "nt":  # pragma: no cover - Windows specific flags
        creation_flags = subprocess.CREATE_NEW_CONSOLE
        if hasattr(subprocess, "DETACHED_PROCESS"):
            creation_flags |= subprocess.DETACHED_PROCESS  # type: ignore[attr-defined]
        popen_kwargs["creationflags"] = creation_flags
    else:
        popen_kwargs["start_new_session"] = True

    subprocess.Popen(command, **popen_kwargs)

    deadline = time.time() + timeout
    while time.time() < deadline:
        if _is_server_live(host, port):
            return True
        time.sleep(0.5)

    return _is_server_live(host, port)


def main() -> None:
    if ensure_backend():
        print(f"✅ Backend is running on http://{DEFAULT_HOST}:{DEFAULT_PORT}")
    else:
        raise SystemExit("Backend did not start within the timeout window.")


if __name__ == "__main__":  # pragma: no cover
    main()


__all__ = ["ensure_backend"]
