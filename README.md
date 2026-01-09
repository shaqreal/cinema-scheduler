# Cinema Scheduler (Windows)

This folder contains **two executables**:

- `scheduler_ui.exe` — the point-and-click app (recommended)
- `scheduler.exe` — the command-line version (optional)

Both do the same scheduling update; the UI just makes it easier to choose files.

---

## Quick start (recommended)

1. Put your **bookings export** Excel file anywhere (e.g., `WC Bookings 10-24.xlsx`).
2. Double‑click **`scheduler_ui.exe`**.
3. Click **Browse…** and select your bookings file.
4. In **Schedule output**, type a name (or click **Save As…**), e.g. `Schedule.xlsx`.
5. Click **Run**.

When it finishes, the schedule workbook will be saved to the path you chose.

---

## What you should give it

### Bookings export (input)
An Excel file (`.xlsx` / `.xlsm`) with one sheet of booking rows. The program is designed to work across different circuits as long as the export has:

- show title / film name
- week start date (Play Week)
- status (New / Hold / Final, etc.)
- a **Comments** column
- one or more **screen-unit columns** *after* Comments (for example: `Standard`, `Digital`, `ATMOS`, `MPX`, `DBOX`, `3D`, etc.)

**Important:** Every screen-unit column after **Comments** is treated as its own “section” of screens:
- the **first** one is the “standard” pool
- every following column is treated as a separate premium pool and is appended after standard  
  (they do **not** share screens with each other)

### Schedule workbook (output)
- If the file **does not exist**, it will be created.
- If the file **already exists**, it will be updated in place (existing sheets reused).

---

## Command-line usage (optional)

If you prefer CLI (or want to automate):

```bat
scheduler.exe "Bookings.xlsx" "Schedule.xlsx"
```

Same arguments as the UI.

---

## Tips

- If you want consistent screen header counts (e.g., always 7 screens for one site), start from a schedule workbook that already has your preferred headers. The program can expand headers when needed, but it will not delete columns automatically.
- If you see extra duplicate tabs for the same location, it’s usually a sheet-name normalization issue. Newer builds normalize trailing punctuation/spaces; older duplicates can be manually removed after verifying they’re not used.

---

## Troubleshooting

### Clicking Run opens another UI window
This happens when the UI is accidentally launching itself instead of the scheduler.

Fix:
- Make sure you have **both** `scheduler.exe` **and** `scheduler_ui.exe` in the same folder.
- Do not rename `scheduler_ui.exe` to `scheduler.exe`.

### “File is open” / permission errors
Close the schedule workbook in Excel before running, and make sure OneDrive isn’t locking the file.

---

## License
MIT — see `LICENSE`.
