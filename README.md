# Sandbaggers League Website

A static website (plain HTML, CSS, vanilla JavaScript — no frameworks, no
database) for the Sandbaggers golf league.

- **Home course:** Viera East Golf Club — Green tees, 9 holes
- **League night:** Thursday evenings
- **Hosting:** deployed live from this GitHub repo via Higgsfield. Every push
  to `main` redeploys the site automatically.

## How the data works

The site reads all of its golf results from a single public file:
**`data.json`**. That file is generated from the league's Excel workbook by
the script **`build_data.py`** — you never edit `data.json` by hand.

```
Excel workbook  ->  build_data.py  ->  data.json  ->  website
```

The workbook is the source of truth. It is **never** committed to GitHub (it
lives outside the repo, in Dropbox / on your Mac). `data.json` contains golf
results only — no email addresses or phone numbers.

## Weekly update routine

After entering a week's results in the Excel workbook:

```bash
cd ~/dev/sandbaggers-league
python3 build_data.py
git add data.json
git commit -m "Update results through week N"
git push
```

Higgsfield redeploys automatically after the push. That's the whole loop.

### Notes

- `build_data.py` automatically finds a file named
  `Sandbaggers League Workbook*.xlsx` in the current folder. Keep **only one**
  workbook there so it always reads the right file.
- To read a workbook stored somewhere else, pass its path:
  ```bash
  python3 build_data.py "/path/to/Sandbaggers League Workbook.xlsx"
  ```
- After running, `python3 build_data.py` prints a short summary (players,
  weeks, result rows). Glance at it to confirm the numbers look right before
  committing.

## First-time setup on a new computer

```bash
# 1. Get the code
mkdir -p ~/dev && cd ~/dev
git clone https://github.com/skarbekadam/sandbaggers-league.git
cd sandbaggers-league

# 2. Install the one library the script needs
python3 -m pip install openpyxl
# (if you see an "externally-managed-environment" error, add --user)

# 3. Put a copy of the workbook in this folder, then run:
python3 build_data.py
```

## Rules (important)

- **Never** commit the raw `.xlsx` workbook. It is excluded by `.gitignore`.
- **Never** publish player email addresses or phone numbers. `build_data.py`
  deliberately reads only name, GHIN, and status from the roster.
- Keep the code beginner-friendly and well-commented.
- Do **not** keep this repo inside a Dropbox (or other cloud-synced) folder —
  the sync process fights with git and can corrupt the repo. Keep it in a plain
  local folder like `~/dev/sandbaggers-league`.

## Files

| File            | What it is                                              |
|-----------------|---------------------------------------------------------|
| `build_data.py` | Reads the Excel workbook, writes `data.json`.           |
| `data.json`     | Public results file the website reads. Auto-generated.  |
| `.gitignore`    | Keeps the workbook, `.DS_Store`, and `.claude/` out of git. |
