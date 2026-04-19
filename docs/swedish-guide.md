# Statsbook-utskrift — bara de nödvändiga sidorna, utan onödig text

En guide för att skriva ut WFTDA StatsBook rent och snabbt från CRG, utan onödiga flikar och utan den lilla metadata-raden (filnamn, datum, sidnummer) som annars hamnar i headers och footers.

**Målgrupp:** HNSO:er, TH:er och andra som ansvarar för paperwork före match. Guiden täcker **Mac, Windows och Linux**.

---

## Vad problemet är

När du laddar ner en statsbook från CRG ("Update" + exportera som .xlsx) och öppnar den i Excel eller LibreOffice, innehåller filen **15+ flikar**. När man skriver ut hela filen blir det:

- **För många sidor:** Du får flikar som Team-översikt, statistik per spelare och annat som inte ska finnas på pappersbunten för matchen.
- **Metadata i utskriften:** Filnamn, datum, sidnummer och ibland sökvägar trycks automatiskt i sidhuvud/sidfot. Det ser fult ut på paperworket och är onödig information för officials som ska jobba med pappret.
- **Manuellt klickande:** Varje helg måste någon sitta och välja ut rätt flikar och rensa headers/footers manuellt. Det tar tid och det är lätt att missa.

## Vad lösningen gör

En liten scriptfil som du dubbelklickar på (eller kör från terminal). Den:

1. Hittar senaste statsbook-filen som CRG har exporterat
2. Tar bort alla flikar som inte ska skrivas ut
3. Rensar sidhuvud/sidfot helt
4. Konverterar till PDF
5. Sparar PDF:en i en Statsbooks-mapp med rätt namn
6. Hoppar över matcher som redan är klara (så du kan köra den flera gånger om du vill)

Du får en ren PDF med exakt de flikar som ska i pappersbunten.

### Vilka flikar behålls

| Flik | Antal sidor |
|---|---|
| IGRF | 2 |
| Score | 4 |
| Penalties | 2 |
| Penalties-Lineups | 2 |
| Expulsion-Suspension Form | 1 |
| Official Reviews | 2 |
| Penalty Box | 4 |
| Rosters | 1 |

Summa: **18 sidor per match**. "Lineups" som separat flik tas bort — den informationen finns redan i "Penalties-Lineups".

---

## Mappstrukturen

Samma oavsett operativsystem. Scriptet förväntar sig:

```
Derby grejer/                       ← huvudmapp (byt namn om du vill)
├── statsbook-print.py              ← Python-scriptet (samma på alla OS)
├── statsbook-print.command         ← launcher för Mac (dubbelklick)
├── statsbook-print.bat             ← launcher för Windows (dubbelklick)
├── statsbook-print.sh              ← launcher för Linux (dubbelklick)
├── crg-scoreboard_v2025.8/         ← din befintliga CRG-scoreboard
│   └── html/
│       └── game-data/
│           └── xlsx/               ← här hamnar STATS-*.xlsx automatiskt
└── Statsbooks/                     ← hit sparas PDF:erna (skapas automatiskt)
```

Håll ihop Python-scriptet, CRG-mappen och Statsbooks-mappen i **samma föräldramapp**.

---

## Python-scriptet (samma på alla OS)

Skapa en fil som heter `statsbook-print.py` i din Derby-mapp. Öppna i valfri textredigerare och klistra in:

```python
#!/usr/bin/env python3
"""Statsbook-utskrift: CRG xlsx → rensad PDF med bara utskriftsflikarna."""
import sys, os, shutil, subprocess, tempfile
from pathlib import Path

script_dir = Path(__file__).resolve().parent
crg_folder = script_dir / "crg-scoreboard_v2025.8" / "html" / "game-data" / "xlsx"
output_folder = script_dir / "Statsbooks"
output_folder.mkdir(exist_ok=True)

# Flikar som ska behållas (resten tas bort)
KEEP_SHEETS = {
    "IGRF", "Score", "Penalties", "Penalties-Lineups",
    "Expulsion-Suspension Form", "Official Reviews", "Penalty Box", "Rosters"
}

try:
    import openpyxl
    from pypdf import PdfReader
except ImportError as e:
    print(f"FEL: Saknar Python-paket: {e}")
    print("Kör: pip3 install openpyxl pypdf")
    sys.exit(1)

# Hitta LibreOffice
def find_soffice():
    """Hitta soffice-kommandot på olika OS."""
    candidates = [
        "soffice",                                              # Linux/Mac med PATH
        "/Applications/LibreOffice.app/Contents/MacOS/soffice", # Mac standard
        r"C:\Program Files\LibreOffice\program\soffice.exe",    # Windows 64-bit
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",  # Windows 32-bit
    ]
    for c in candidates:
        try:
            subprocess.run([c, "--version"], capture_output=True, timeout=10)
            return c
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None

soffice = find_soffice()
if not soffice:
    print("FEL: Kunde inte hitta LibreOffice (soffice).")
    print("Installera LibreOffice från https://www.libreoffice.org/")
    sys.exit(1)

if not crg_folder.exists():
    print(f"FEL: Hittar inte CRG-mappen: {crg_folder}")
    print("Kontrollera att CRG-scoreboard ligger i samma föräldramapp som detta script.")
    sys.exit(1)

# Hitta xlsx-filer i CRG
xlsx_files = sorted(crg_folder.glob("STATS-*.xlsx"), key=os.path.getmtime, reverse=True)
if not xlsx_files:
    print("Inga statsbooks hittade i CRG.")
    sys.exit(1)

processed = 0
for xlsx_path in xlsx_files:
    output_pdf = output_folder / f"{xlsx_path.stem}_IGRF.pdf"
    if output_pdf.exists() and output_pdf.stat().st_mtime >= xlsx_path.stat().st_mtime:
        continue

    print(f"Bearbetar: {xlsx_path.name}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_xlsx = Path(tmpdir) / xlsx_path.name
        shutil.copy2(xlsx_path, tmp_xlsx)

        wb = openpyxl.load_workbook(tmp_xlsx)

        # Ta bort onödiga flikar
        for name in list(wb.sheetnames):
            if name not in KEEP_SHEETS:
                del wb[name]

        # Rensa headers/footers — behåll bara fliknamnet
        for ws in wb.worksheets:
            ws.oddHeader.left.text = "&A"
            ws.oddHeader.center.text = None
            ws.oddHeader.right.text = None
            ws.oddFooter.left.text = None
            ws.oddFooter.center.text = None
            ws.oddFooter.right.text = None
            ws.evenHeader.left.text = None
            ws.evenHeader.center.text = None
            ws.evenHeader.right.text = None
            ws.evenFooter.left.text = None
            ws.evenFooter.center.text = None
            ws.evenFooter.right.text = None

        wb.save(tmp_xlsx)
        print(f"  Flikar: {', '.join(wb.sheetnames)}")

        # Konvertera till PDF via LibreOffice
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, str(tmp_xlsx)],
            capture_output=True, text=True, timeout=120
        )

        tmp_pdf = Path(tmpdir) / f"{xlsx_path.stem}.pdf"
        if not tmp_pdf.exists():
            print(f"  FEL: Kunde inte konvertera till PDF")
            if result.stderr:
                print(f"  {result.stderr}")
            continue

        shutil.move(tmp_pdf, output_pdf)

        pages = len(PdfReader(output_pdf).pages)
        print(f"  → {output_pdf.name} ({pages} sidor)")
        processed += 1

if processed == 0:
    print("\nAlla matcher redan klara.")
else:
    print(f"\nKlart! {processed} match(er) bearbetade.")
```

Det här är **exakt samma kod på Mac, Windows och Linux**. Det som skiljer är bara *launchern* — den lilla filen som startar Python med rätt inställningar för ditt operativsystem.

---

## Installation per operativsystem

### Mac

**1. Python 3**

Öppna Terminal (Program → Verktygsprogram, eller via Spotlight). Kontrollera:

```bash
python3 --version
```

Om du får `Python 3.10.x` eller högre är du bra. Om inte:
- Installera [Homebrew](https://brew.sh)
- Kör sedan: `brew install python@3.11`

**2. LibreOffice**

Ladda ner gratis från [libreoffice.org](https://www.libreoffice.org/download/download/). Dra appen till Program-mappen som vanligt.

**3. Python-paketen**

```bash
pip3 install openpyxl pypdf
```

**4. Launcher — `statsbook-print.command`**

Skapa en fil som heter `statsbook-print.command` i samma mapp som Python-scriptet, med innehållet:

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 statsbook-print.py
```

Gör den körbar:

```bash
cd "/Users/dittnamn/Documents/Derby grejer"
chmod +x statsbook-print.command
```

Nu kan du **dubbelklicka** på `statsbook-print.command` för att köra scriptet.

---

### Windows

**1. Python 3**

Ladda ner från [python.org/downloads](https://www.python.org/downloads/). **Viktigt:** bocka för "Add Python to PATH" under installationen, annars fungerar inte dubbelklickandet.

Öppna Kommandotolken (Command Prompt) — tryck Win+R, skriv `cmd`, Enter. Kontrollera:

```cmd
python --version
```

Du ska få `Python 3.10.x` eller högre.

**2. LibreOffice**

Ladda ner från [libreoffice.org](https://www.libreoffice.org/download/download/) och installera.

**3. Python-paketen**

I Kommandotolken:

```cmd
pip install openpyxl pypdf
```

**4. Launcher — `statsbook-print.bat`**

Skapa en fil som heter `statsbook-print.bat` i samma mapp som Python-scriptet, med innehållet:

```bat
@echo off
cd /d "%~dp0"
python statsbook-print.py
pause
```

`pause`-raden gör att fönstret stannar öppet efter att scriptet körts klart, så du hinner se om det blev fel. Ta bort den raden om du hellre vill att fönstret stängs direkt.

Nu kan du **dubbelklicka** på `statsbook-print.bat`.

---

### Linux

**1. Python 3**

De flesta Linux-distributioner har Python 3 förinstallerat. Kontrollera i terminal:

```bash
python3 --version
```

Om du behöver installera (Debian/Ubuntu):

```bash
sudo apt install python3 python3-pip
```

Fedora:

```bash
sudo dnf install python3 python3-pip
```

**2. LibreOffice**

De flesta distributioner har det förinstallerat. Annars:

```bash
sudo apt install libreoffice       # Debian/Ubuntu
sudo dnf install libreoffice       # Fedora
sudo pacman -S libreoffice-fresh   # Arch
```

**3. Python-paketen**

```bash
pip3 install openpyxl pypdf
```

Om din distribution har låst `pip3` (PEP 668), använd:

```bash
pip3 install --user openpyxl pypdf
```

Eller skapa en virtuell miljö:

```bash
python3 -m venv ~/.venvs/statsbook
~/.venvs/statsbook/bin/pip install openpyxl pypdf
```

Och anpassa launchern så den använder `~/.venvs/statsbook/bin/python` istället för `python3`.

**4. Launcher — `statsbook-print.sh`**

Skapa en fil som heter `statsbook-print.sh` i samma mapp som Python-scriptet, med innehållet:

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 statsbook-print.py
```

Gör den körbar:

```bash
cd ~/Documents/Derby\ grejer
chmod +x statsbook-print.sh
```

På de flesta Linux-desktops (GNOME, KDE, Cinnamon) kan du högerklicka → Properties → Permissions → "Allow executing file as program", sen dubbelklicka på filen.

Beroende på ditt filhanterarprogram kan du behöva välja "Run" eller "Run in terminal" när du dubbelklickar första gången. Om dubbelklick inte fungerar, kör från terminal:

```bash
./statsbook-print.sh
```

---

## Så här använder du det (alla OS)

1. **Öppna CRG-scoreboard** och kör matchen som vanligt.
2. När matchen är klar, klicka **"Update"** i CRG så att den senaste datan sparas ner till xlsx-filen.
3. **Dubbelklicka** på launcher-filen för ditt OS (`.command`, `.bat`, eller `.sh`). Ett terminalfönster öppnas och scriptet körs.
4. När det står "Klart!" är du klar.
5. **Öppna `Statsbooks/`-mappen** — där ligger PDF:en, redo att skriva ut.

Du kan köra scriptet flera gånger under helgen — det bearbetar bara nya eller uppdaterade filer.

---

## Hur det funkar rad-för-rad (för den nyfikne)

Scriptet är kort och läsbart. Här är vad som händer:

1. **Hittar CRG-mappen** relativt scriptets egen plats — så länge allt ligger i samma föräldramapp fungerar det på alla OS.
2. **Söker efter LibreOffice** på flera möjliga platser (Mac, Windows, Linux) tills det hittar en fungerande installation.
3. **Letar efter `STATS-*.xlsx`**-filer i CRG:s exportmapp, sorterade med nyaste först.
4. **För varje fil som inte redan har en färdig PDF:**
   - Kopierar filen till en tillfällig plats (så originalet inte ändras)
   - Öppnar den med `openpyxl` (Python-bibliotek för Excel)
   - Går igenom alla flikar och raderar de som **inte** finns i `KEEP_SHEETS`-listan
   - Går igenom kvarvarande flikar och **tömmer alla sidhuvuden och sidfötter**, utom att fliknamnet visas överst (det är `&A`-koden i Excel)
   - Sparar den rensade Excel-filen
   - Anropar LibreOffice i headless-läge för att konvertera till PDF
   - Flyttar PDF:en till `Statsbooks/`-mappen med rätt namn

Hela processen tar några sekunder per match.

---

## Anpassa för din förening

**Andra flikar?** Redigera `KEEP_SHEETS`-listan i `statsbook-print.py`:

```python
KEEP_SHEETS = {
    "IGRF", "Score", "Penalties", "Penalties-Lineups",
    "Expulsion-Suspension Form", "Official Reviews", "Penalty Box", "Rosters"
}
```

Lägg till eller ta bort flikar efter behov. Fliknamnen måste stavas **exakt** som de heter i CRG:s statsbook (inklusive bindestreck och versaler).

**Annan CRG-version?** Byt ut `crg-scoreboard_v2025.8` i denna rad:

```python
crg_folder = script_dir / "crg-scoreboard_v2025.8" / "html" / "game-data" / "xlsx"
```

**Behåll headers/footers?** Ta bort hela `# Rensa headers/footers`-blocket.

---

## Felsökning

### Alla OS

**"Inga statsbooks hittade i CRG"**
CRG har inte exporterat några filer ännu. Klicka "Update" i CRG och kör sedan om scriptet.

**"FEL: Saknar Python-paket"**
Paketen är inte installerade. Kör `pip3 install openpyxl pypdf` (Mac/Linux) eller `pip install openpyxl pypdf` (Windows).

**"FEL: Kunde inte hitta LibreOffice (soffice)"**
LibreOffice är inte installerat, eller ligger på en ovanlig plats. Installera det eller kör en manuell test: öppna Terminal/Kommandotolken och skriv `soffice --version`.

**"FEL: Kunde inte konvertera till PDF"**
LibreOffice är installerat men kraschar på filen. Prova att öppna xlsx-filen manuellt i LibreOffice — om det fungerar där men inte via scriptet, starta om datorn (LibreOffice har ibland en bakgrundsprocess som blockerar).

**Scriptet kör men PDF:en är tom eller fel**
Kontrollera att CRG verkligen har sparat ner en ny STATS-*.xlsx. Om filen inte är uppdaterad använder scriptet den befintliga PDF:en. Radera den gamla PDF:en från `Statsbooks/` och kör om.

### Mac-specifika problem

**Scriptet öppnar i TextEdit istället för att köras**
Filen är inte markerad som körbar. Kör `chmod +x statsbook-print.command` i Terminal.

**"soffice: command not found"**
Skapa en symlänk:
```bash
sudo ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice /usr/local/bin/soffice
```

### Windows-specifika problem

**"'python' is not recognized as an internal or external command"**
Python lades inte till i PATH vid installationen. Antingen installera om med "Add Python to PATH" ibockad, eller lägg till det manuellt via Systeminställningar → Avancerade systeminställningar → Miljövariabler.

**.bat-filen öppnas i Notepad istället för att köras**
Högerklicka på filen → Egenskaper → kontrollera att "Öppna med" är satt till "Windows Command Processor" (cmd.exe).

### Linux-specifika problem

**Dubbelklick öppnar filen i textredigerare**
Olika filhanterare har olika default-beteende. Högerklicka → Properties → Permissions → "Allow executing file as program". I GNOME Files kan du också behöva ändra "Executable Text Files" i Preferences till "Run them" eller "Ask what to do".

**"error: externally-managed-environment" vid pip install**
Nyare Debian/Ubuntu blockerar systemwide pip. Använd `pip3 install --user openpyxl pypdf` eller en virtuell miljö (se Linux-installationen ovan).

---

## Delning och anpassning

Detta är ett kort script som kan delas fritt inom derbycommunityn. Om andra ligor vill använda det, peka dem gärna hit. Förbättringar välkomnas.

*Skapat ursprungligen av WaspBee (HNSO/THNSO) för WaspBees egen användning. Öppen för vidareutveckling.*
