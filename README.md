# Statsbook PDF Cleaner

**Print only what you actually need from a WFTDA StatsBook — no extra sheets, no filename stamped in the footer, no fighting with Page Setup.**

If you've ever stood in front of a printer at 8:47 AM on a game day, manually deselecting sheets and clicking "Clear header" for the seventh time that weekend, this one is for you. We know. We've been there. That's why this exists.

This is a small cross-platform tool that takes a StatsBook exported from [CRG scoreboard](https://github.com/rollerderby/scoreboard) and gives you back a clean, print-ready PDF — just the sheets officials need on paper, no auto-generated clutter.

(Derby is a nerd sport. Officiating is, somehow, the nerdier part of a nerd sport. If you're here reading a README file about StatsBook automation, you know exactly what we mean. Welcome.)

This software is not provided, endorsed, produced, or supported by the WFTDA, SweSports, or anyone else except one very tired THNSO with a Python interpreter and strong feelings about paper waste. It is not guaranteed to be free of bugs, edge cases, or mild disappointment — but it has been saving real NSOs real time at real tournaments.

**Honesty disclaimer:** this is an Apple household, so the Mac path is the only one that's been battle-tested at an actual game weekend. The Windows and Linux instructions should work (it's the same Python code and the same LibreOffice), but if you try them first and something breaks in a weird way, please open an issue — we'd rather know than pretend.

<!-- SCREENSHOT-1: Side-by-side comparison. Left: messy Excel print preview with 15 sheet tabs and "STATS-20260418...xlsx" in the footer. Right: clean PDF with just the 8 sheets and no footer junk. Caption: "Before and after. Yes, we know." -->

---

## What it does

Give it a `STATS-*.xlsx` from CRG. It will:

1. **Keep only the sheets you need on paper** — drops team breakdowns, per-player statlines, and all the other tabs that don't belong on the officials' bench.
2. **Strip every header and footer field** — no filename, no timestamp, no page counter, no path. Just the sheet name at the top, clean.
3. **Convert to PDF** via headless LibreOffice.
4. **Save it with a sensible name** in a `Statsbooks/` folder next to the script.
5. **Skip games that are already done**, so you can run it every time CRG updates without re-making the same PDFs.

Default sheets kept (18 pages per game):

| Sheet | Pages |
|---|---|
| IGRF | 2 |
| Score | 4 |
| Penalties | 2 |
| Penalties-Lineups | 2 |
| Expulsion-Suspension Form | 1 |
| Official Reviews | 2 |
| Penalty Box | 4 |
| Rosters | 1 |

The standalone "Lineups" sheet is dropped on purpose — the same info lives in "Penalties-Lineups" and nobody needs both on paper.

---

## What you need

- **Python 3.10+**
- **LibreOffice** (that's what does the PDF conversion)
- **Two tiny Python packages**: `openpyxl` and `pypdf`

That's it. No cloud, no account, no dependencies on WFTDA tools. Runs entirely on your own laptop.

### macOS

```bash
python3 --version                    # check if Python is already there
brew install python@3.11             # only if it isn't
brew install --cask libreoffice      # (or download from libreoffice.org)
pip3 install openpyxl pypdf
```

<!-- SCREENSHOT-2: Terminal showing `python3 --version` returning Python 3.11.x and `soffice --version` returning LibreOffice 7.x. Caption: "Both exist? You're good." -->

### Windows

1. Grab **Python 3** from [python.org/downloads](https://www.python.org/downloads/) and **tick "Add Python to PATH"** during install. If you miss this box, double-clicking `.bat` files won't work and you'll be back in PATH-editing hell. Don't skip it.
2. Install **LibreOffice** from [libreoffice.org](https://www.libreoffice.org/).
3. In Command Prompt:

   ```cmd
   pip install openpyxl pypdf
   ```

### Linux

Debian/Ubuntu:
```bash
sudo apt install python3 python3-pip libreoffice
pip3 install --user openpyxl pypdf
```

Fedora:
```bash
sudo dnf install python3 python3-pip libreoffice
pip3 install --user openpyxl pypdf
```

Arch, because you know what you're doing:
```bash
sudo pacman -S python python-pip libreoffice-fresh
pip install --user openpyxl pypdf
```

---

## Setup

1. **Clone or download** this repo.
2. **Drop it next to your CRG scoreboard folder**, like this:

   ```
   my-derby-stuff/
   ├── statsbook-pdf-cleaner/        ← this repo
   │   ├── statsbook-print.py
   │   ├── statsbook-print.command
   │   ├── statsbook-print.bat
   │   └── statsbook-print.sh
   └── crg-scoreboard_v2025.8/       ← your CRG install (auto-detected)
   ```

   The script looks for any folder matching `crg-scoreboard_v*` next to itself. When CRG releases a new version, it just works. No path edits, no config files.

3. **Mac / Linux only:** make the launcher executable once:

   ```bash
   chmod +x statsbook-print.command statsbook-print.sh
   ```

<!-- SCREENSHOT-3: Finder (or Explorer or Nautilus) window showing the parent folder with crg-scoreboard_v2025.8 and statsbook-pdf-cleaner side by side. Caption: "Keep them together. The tool finds the CRG folder automatically." -->

---

## Using it

1. Run your game in CRG scoreboard as usual.
2. When the game is over (or whenever you want a fresh PDF), hit **"Update"** in CRG so the latest data gets saved to the xlsx export.
3. Double-click the launcher for your OS:
   - **macOS:** `statsbook-print.command`
   - **Windows:** `statsbook-print.bat`
   - **Linux:** `statsbook-print.sh`
4. A terminal opens. You'll see something like:

   ```
   Using CRG scoreboard: crg-scoreboard_v2025.8
   Processing: STATS-20260418-HomeTeam-vs-AwayTeam.xlsx
     Sheets kept: IGRF, Score, Penalties, Penalties-Lineups, Expulsion-Suspension Form, Official Reviews, Penalty Box, Rosters
     → STATS-20260418-HomeTeam-vs-AwayTeam_PRINT.pdf (18 pages)

   Done! 1 game(s) processed.
   ```

5. Open the `Statsbooks/` folder. Your clean PDF is there. Print it.

<!-- SCREENSHOT-4: Terminal with a successful run on a real game (game name blurred/changed for privacy). Caption: "Takes about five seconds per game." -->

Run it again any time during the weekend — it only processes new or updated games, so it's fine to trigger it every time CRG updates.

### Command-line options

If you want to get fancy:

```
python3 statsbook-print.py [--crg PATH] [--output PATH]
```

- `--crg PATH` — point at a specific CRG folder (overrides auto-detection)
- `--output PATH` — save the PDFs somewhere other than `Statsbooks/`

You can also use the `CRG_PATH` environment variable instead of `--crg`.

---

## Making it yours

### Different sheets

Edit the `KEEP_SHEETS` set near the top of `statsbook-print.py`:

```python
KEEP_SHEETS = {
    "IGRF", "Score", "Penalties", "Penalties-Lineups",
    "Expulsion-Suspension Form", "Official Reviews", "Penalty Box", "Rosters",
}
```

Sheet names must match exactly — including hyphens and capitalization. CRG is particular.

### Keep the headers/footers after all

Delete or comment out the `# Clear headers/footers` block. The script will leave CRG's defaults alone.

### Put the tool somewhere that isn't next to CRG

Use the `--crg` flag or set `CRG_PATH`:

```bash
CRG_PATH=/path/to/crg-scoreboard_v2025.8 python3 statsbook-print.py
```

---

## How it works, roughly

Under the hood:

1. Auto-detects the newest `crg-scoreboard_v*/html/game-data/xlsx/` folder next to itself.
2. Finds LibreOffice (tries common install paths on each OS).
3. Lists all `STATS-*.xlsx` files in the CRG export folder.
4. For each one that doesn't already have a matching PDF:
   - Copies the xlsx to a temp directory (your original stays untouched).
   - Uses `openpyxl` to delete the sheets that aren't in `KEEP_SHEETS`.
   - Clears every header and footer field on every remaining sheet, keeping only the sheet name at the top via Excel's `&A` code.
   - Runs `soffice --headless --convert-to pdf` to turn the rewritten xlsx into a PDF.
   - Moves the PDF to `Statsbooks/`.

Five seconds per game, give or take.

---

## Troubleshooting

### `ERROR: Missing Python package`
`pip3 install openpyxl pypdf` (or `pip install` on Windows). Run it in the same terminal you'll run the script from.

### `ERROR: LibreOffice (soffice) not found`
Either it's not installed, or it's hiding in an unexpected spot. Install from [libreoffice.org](https://www.libreoffice.org/).

On macOS, a symlink usually fixes discovery:
```bash
sudo ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice /usr/local/bin/soffice
```

### `ERROR: Could not locate a CRG scoreboard folder`
Move the tool next to your `crg-scoreboard_v*` folder, or pass `--crg /path/to/crg-scoreboard_vX.Y`.

### `No statsbooks found`
CRG hasn't exported a StatsBook yet. Click **Update** in CRG and try again.

### LibreOffice conversion fails silently
LibreOffice sometimes keeps a background process running that blocks headless mode. Quit LibreOffice completely (on macOS that means `Cmd+Q` on the dock icon, not just closing the window) and try again.

### Double-clicking opens the launcher in a text editor
- **macOS:** `chmod +x statsbook-print.command` in Terminal. Finder might also ask once whether you want to open it — say yes.
- **Linux:** Right-click → Properties → Permissions → "Allow executing file as program." Some file managers also need "Executable text files" set to "Run" in their preferences.
- **Windows:** Right-click the `.bat` → Properties → make sure it opens with Windows Command Processor (`cmd.exe`). If not, "Open with" → pick `cmd.exe`.

### Still broken?
Open an issue with the error message. Please include your OS, Python version, LibreOffice version, and the first few lines of output from the script. Screenshots help. Apologies in advance for any dire warnings of awful consequences.

---

## Contributing

PRs welcome. Especially from anyone willing to help with:

- **Windows and Linux testing.** The code should work (we've been careful), but "should work" is not the same as "has been run at 7 AM by a stressed HNSO before day two". Real-world reports wanted.
- **Windows-specific fixes** for things we haven't thought of.
- **Non-default CRG folder layouts** — some leagues run CRG in places the auto-detect won't find.
- **A tiny GUI** for NSOs who would rather click than terminal.
- **Translations of the guide.** The Swedish version is in [`docs/swedish-guide.md`](docs/swedish-guide.md). Others welcome.

If you're an NSO who isn't comfortable with code but has a good bug report, that's genuinely useful too. Issues > silence.

---

## License

MIT — see [LICENSE](LICENSE). Fork it, rename it, run it at your tournament, put your own league's logo on it. Just don't sell it to WFTDA as official merch.

Originally built by WaspBee (THNSO, Sweden), because the forty-fifth time manually deselecting "Lineups" before printing was forty-four too many.

Free for every league and every official, forever.
