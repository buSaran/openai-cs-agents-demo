@echo off
rem Ensure the backend is running then open Randell's form.
pushd "%~dp0\.."
python -m eod_app.ensure_backend
start "" "http://127.0.0.1:8000/randells"
popd
