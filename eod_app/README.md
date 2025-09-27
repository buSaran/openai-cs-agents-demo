# End of Day app

The FastAPI app in this directory lets staff file end-of-day reports for Randell's and Best One.  By default the forms only use the values you key in manually.

## Optional Gmail integration

If you want the forms to preload the latest till totals from the daily POS emails, you can opt in to the Gmail fetcher:

1. Enable IMAP in Gmail (Settings → Forwarding and POP/IMAP → Enable IMAP).
2. Turn on 2-step verification for the account.
3. Create an [App Password](https://support.google.com/mail/answer/185833) named for this tool.
4. Create an `.env` file alongside `main.py` that contains:

   ```ini
   EOD_ENABLE_GMAIL=1
   GMAIL_USER=your.name@gmail.com
   GMAIL_PASS=the_app_password
   ```

   You can generate this file interactively by running `python -m eod_app.setup_env` and pasting the 16-character App Password
   (for example, `meouwmpimppuilbj`) when prompted.  The script writes `eod_app/.env` for you and never echoes the password back
   to the console.

The helper uses IMAP with the credentials above.  If Google blocks the login it means the account is still using a normal password or IMAP is disabled—follow the steps to create an app password and try again.  Leave `EOD_ENABLE_GMAIL` unset (or set it to 0) to skip Gmail entirely.

### Replacing a blocked Gmail account

Google occasionally locks accounts that look automated.  When that happens the quickest recovery path is to create a fresh Gmail address that is dedicated to the End of Day feed:

1. Sign up for a new Gmail account and forward the POS alert emails (from `vbralert@visualbusinessretail.co.uk`) to it.  Keeping this mailbox separate reduces the risk of future lockouts.
2. Enable IMAP, turn on 2-step verification, and generate a new App Password for that account just like above.
3. Update your `.env` file so `GMAIL_USER` is the new inbox and `GMAIL_PASS` is the new App Password.  You can temporarily set `EOD_ENABLE_GMAIL=0` while you are switching accounts to keep the forms usable.
4. Test the credentials locally with:

   ```bash
   python -m eod_app.zreport_gmail_api --shop "Randell's"
   ```

   The command will report whether the login succeeded and show the most recent totals it can see.  Run it again with `--shop "Best One"` to confirm both stores are covered.

Once the CLI check succeeds you can re-enable `EOD_ENABLE_GMAIL=1`, restart the FastAPI app, and the forms will pull in the new account automatically.

## Starting the app without a console window

The repository includes helper scripts that ensure the FastAPI backend is running before a browser tab opens.  On Windows you can double-click either batch file in `scripts\windows`:

* `open_randells_eod.bat`
* `open_bestone_eod.bat`

Each shortcut launches `python -m eod_app.ensure_backend`, which checks whether the server is already bound to `http://127.0.0.1:8000` and starts it if not.  The script waits for the port to come up and then opens the relevant form in your default browser, so you no longer need to keep a terminal window open manually.

