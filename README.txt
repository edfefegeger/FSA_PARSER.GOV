FSA Registry Parser - Instructions
================================

Installation on Linux (Ubuntu/Debian)
sudo apt install python3-venv     # if not already installed
chmod +x run.sh
./run.sh                          # automatically creates .venv and installs dependencies
The `pip install` command without a venv does not work in Ubuntu 24.04 (due to PEP 668 protection) – this is expected behavior.
Manually:  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python fsa_parser.py

Installation on Windows (one-time setup)
1. Install Python 3.9+ from python.org (check the "Add Python to PATH" box during installation).
2. Open the command prompt in the parser folder and run:
pip install -r requirements.txt

Execution (weekly)
- Linux:   ./run.sh   (same arguments apply: ./run.sh --days 90)
- Windows: double-click run.bat
or via console:  python fsa_parser.py
- Change threshold for a single run:  python fsa_parser.py --days 90

Output
The file appears in the "exports" folder: fsa_60_days_YYYY-MM-DD.xlsx
Old exports are not overwritten (if run again on the same day,
suffixes like _2, _3, etc., are added). Documents expiring within 30 days are highlighted in red. SETTINGS (config.ini, open with Notepad)
- days_threshold    - how many days before expiration to include the document
- product_keywords  - words/word roots for product filtering
- document_types    - certificates, declarations, or both
- format            - xlsx / csv / both
- columns           - column composition and order
- urgent_days       - threshold for red highlighting
Run frequency is simply how often you execute run.bat.

IF THE SITE DROPS THE CONNECTION / BLOCKS YOUR IP
The parser stops automatically at the first sign of a block (it doesn't hammer the site with retries),
saves everything found to the 'cache' folder, and writes a ..._partial.xlsx file.
Wait (several hours in the event of a block) and run it again: it will resume from where
it stopped. To avoid triggering a block:
- do not reduce the pauses between requests (delay, batch_pause);
- set max_requests = 300: the parser will stop automatically after 300 requests; run
it several times throughout the day—each run continues from the previous one;
- recurring weekly runs fetch applicant data from the cache (30-day validity) and place
minimal load on the site.
You can delete the 'cache' folder to force a fresh download of everything.

IF SOMETHING GOES WRONG
- Empty "Head's Full Name" or "e-mail" fields: run `python fsa_parser.py --debug`
and send the `debug_*.json` files to the developer—these contain the site's responses,
allowing the field logic to be fixed in minutes.
- Network errors: the parser automatically retries requests; in case of repeated failures,
increase the `delay` in config.ini.
- The FSA site might change its API; in that case, the parser will need an update.
