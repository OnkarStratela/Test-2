# Beer-Pour RFID Verification Test

Test harness for a **beer-pour-machine** where a tagged cup is slid over one
of two antennas. For every cup-slide ("trial") the harness drives the
arbitrated dual-antenna scanner (`rfid_gc_live`), records which antenna
the arbitrator picked, how fast it picked, and whether the *other*
antenna also saw the tag at any point during the slide -- the cross-read
failure mode the arbitrator is designed to prevent.

The operator does **not** declare which antenna the cup will be slid over.
In real-world use the system has no advance knowledge of that -- a user
can put their cup on either antenna -- so the test mirrors that: just
slide the cup and let the system make the call.

## What the test measures

Each trial is **one row** in `beer-pour-results.xlsx` with these columns (in order):

| Column                       | What it means                                                                       |
|------------------------------|-------------------------------------------------------------------------------------|
| `session_id`                 | Timestamp of the test session (groups all trials from one run together).            |
| `scenario`                   | Physical setup under test (e.g. `drip_tray_empty_cup_empty`).                       |
| `power_mw`                   | TX power used for this trial (mW).                                                  |
| `swap_rate_3s_hz`            | Per-EPC antenna swaps observed within the first 3 s, in Hz. `0.0` = stable answer; rising values = the arbitrator was flipping the tag between antennas. |
| `winning_antenna`            | Per-EPC home antenna (last 6 hex chars → `antN`). E.g. with two cups: `6E6F76 -> ant0, 6E6FD6 -> ant1`. |
| `cross_reads`                | Number of EPCs that appeared on **both** antennas during the trial (same-tag leakage). **0 is the happy path.** Two different tags, each only on its own antenna, does **not** count. |
| `best_rssi_winner_dbm`       | Strongest (closest-to-zero) RSSI on the winning antenna across the trial.           |
| `detected_epcs`              | Every distinct EPC seen during the trial, regardless of antenna.                    |
| `tag_photo`                  | Embedded thumbnail of `images/tags/<tag>.png`.                                      |
| `result`                     | One-word verdict (colour-coded): **PASS** (verified ≤3 s, clean) / **SLOW** (verified + clean, >3 s) / **DIRTY** (same EPC seen on both antennas) / **FAIL** (no attribution). |
| `notes`                      | Free-text comment you typed at the post-trial prompt.                               |

## Test axes

The three axes are selected from menus and persist across trials until
you change them with `p` / `s` / `t`. The test does **not** ask which
antenna the cup will be slid over -- the harness derives the winner
from the data.

| Axis              | Values                                                                       |
|-------------------|------------------------------------------------------------------------------|
| Scenario          | `drip_tray_empty_cup_empty`, `drip_tray_half_full_cup_empty` (editable in `SCENARIOS` in `beer_pour_logger.py`) |
| Power (mW)        | `30`, `175`, `316` (editable in `POWER_LEVELS_MW`)                            |
| Tag type          | Auto-discovered from `images/tags/*.png`. Drop e.g. `images/tags/foam.png` to register a new tag. |

After every trial the harness prints a verdict line and asks:

1. **`Save this trial? [Y/n]`** — ENTER (or `y`) keeps the row, `n` discards it.
2. **`Comment for notes column (ENTER to skip)`** — optional free-text
   note that goes into the `notes` column for this row (e.g.
   "cup slid too slowly", "tag was loose", "this is the one to show in
   the demo"). Press ENTER for no comment.

The per-`(scenario, power, tag)` trial counter only advances when you
choose to save, so the spreadsheet stays sequential even if you re-do a
trial that didn't go well.

**Results always append to the same `beer-pour-results.xlsx`** for as
long as the schema matches. Every run writes new rows under the
existing header — `session_id` is stamped per run so you can still tell
which run a row came from. If a newer version of this script changes
the column layout, the existing file is renamed to
`beer-pour-results_archive_<timestamp>.xlsx` so your old data is
preserved alongside, and a fresh `beer-pour-results.xlsx` is created
for the new layout.

## Files

| File                       | Purpose |
|----------------------------|---------|
| `rfid_gc_live.c`           | Arbitrated dual-antenna scanner. Same antenna-winner logic as the standalone binary — the harness just spawns one instance per trial and parses its sweep output. |
| `rfid_standard.c`          | Unarbitrated baseline scanner (every read printed, no per-window decision). Useful for sanity-checking the radio without the arbitrator in the loop. |
| `compile_gc.sh`            | Builds `rfid_gc_live` against `SRC/`. |
| `compile_standard.sh`      | Builds `rfid_standard` against `SRC/`. |
| `run_gc.sh`                | Compile-and-run wrapper for `rfid_gc_live` (interactive use, no logging). |
| `run_standard.sh`          | Compile-and-run wrapper for `rfid_standard`. |
| `run_test.sh`              | **One-shot test runner.** Checks SRC, fixes USB perms, compiles `rfid_gc_live`, launches the Python logger. Preferred entry point for a test session. |
| `beer_pour_logger.py`      | Python test harness. Drives `rfid_gc_live` one trial at a time, captures every window, computes the metrics above, appends to `beer-pour-results.xlsx` (with operator confirmation). |
| `requirements.txt`         | Python deps for the logger: `openpyxl`, `Pillow`. |
| `images/scenarios/`        | Optional photos of each scenario — embedded next to the trial row when present. |
| `images/tags/`             | Photos of each tag type. **The filename (without `.png`) is the tag name** that shows up in the menu. |
| `SRC/`                     | CAEN RFID Light library sources/headers — do not modify. |

## Hardware

- Raspberry Pi (CM4 or any model with USB)
- CAEN R3100C-Lepton3 25 dBm RFID reader on `/dev/ttyACM0` (USB)
- 2× UHF antennas on `Source_0` (antenna 0) and `Source_1` (antenna 1), with
  the layout the arbitrator was tuned for: 150 mm centre-to-centre,
  cup-on-antenna distance ~5–7 cm, opposite-antenna distance ~15 cm.

## Setup

```bash
# 1. Make scripts executable (first time only)
chmod +x compile_gc.sh compile_standard.sh run_gc.sh run_standard.sh run_test.sh

# 2. Install Python deps for the logger.
#    On Pi OS Bookworm + later, pip3 is locked down by PEP 668 — use apt:
sudo apt install -y python3-openpyxl python3-pil
#    (or, if those packages aren't available on your distro:
#       pip3 install --break-system-packages -r requirements.txt )

# 3. Build the C reader (or just let run_test.sh do it for you)
./compile_gc.sh
```

## Run a test session

Easiest path — one-shot runner (compiles + USB perms + launches the logger):

```bash
./run_test.sh
```

Or run the Python directly after compiling:

```bash
./compile_gc.sh
python3 beer_pour_logger.py
```

Sample session:

```
=== Beer-pour RFID verification test session 20260527-143200 ===
Results : /home/stratela/Test-2/beer-pour-results.xlsx
Reader  : rfid_gc_live
Trial   : 5.0 s  (verify deadline: 3.0 s)

Available scenarios:
  1. drip_tray_empty_cup_empty
  2. drip_tray_half_full_cup_empty
Select scenarios [1..2]: 1

Available power levels (mW):
  1. 30
  2. 175
  3. 316
Select power levels (mW) [1..3]: 1

Available tag types:
  1. foam
Select tag types [1..1]: 1

[drip_tray_empty_cup_empty | 30 mW | foam]
  ENTER = start trial   'p' = power   's' = scenario   't' = tag   'q' = quit
> <ENTER>

[Trial #1] Launching reader at 30 mW... (reader takes ~1-2 s to come up)
    ===== Dual-Antenna RFID Live Scanner (arbitrated) =====
    [GC] Reader: R3100C  Serial: ...
    [GC] Ready. Empty sweeps print []. Tagged sweeps prepend [TX …]. Ctrl+C to stop.
[Trial #1] GO! Slide the foam cup over either antenna now (5.0 s).
    [TX=30 mW] [(0)(-58.3) E2801160600002054E1A1234,                                      ]
    [TX=30 mW] [(0)(-58.4) E2801160600002054E1A1234,                                      ]
    [TX=30 mW] [(0)(-58.5) E2801160600002054E1A1234,                                      ]
    []
    [TX=30 mW] [(0)(-58.6) E2801160600002054E1A1234,                                      ]
[Trial #1] PASS -- verified in 1.00 s, homes [E1A1234->ant0], 0 leaking EPC(s), best RSSI -58.3 dBm
    Save this trial? [Y/n]: <ENTER>
    Comment for notes column (ENTER to skip): cup slid smoothly over ant0
    logged to beer-pour-results.xlsx
```

The reader's sweep lines are forwarded verbatim to your terminal (same
ANSI-coloured output as running `rfid_gc_live` by hand), and at the end
the harness prints a single-line verdict.

### Verdict short-codes

| Verdict   | Meaning |
|-----------|---------|
| **PASS**  | `verified == yes`, `within_3s == yes`, `cross_reads == 0`. The happy path. |
| **SLOW**  | Verified + clean, but `ttv_s > 3.0`. The arbitrator got it right but missed the 3-second deadline (e.g. the operator hadn't actually placed the cup yet). |
| **DIRTY** | Verified, but at least one EPC was attributed to both antennas during the trial (same tag leaking). Two different tags on two different antennas is **not** DIRTY. |
| **FAIL**  | No attribution at all in the trial. |

Press `'q'` (or Ctrl-C) at the menu to end the session.

## How trials are isolated

`rfid_gc_live` keeps its decision-window stats in memory for the life of
the process. To make every trial a clean slate, the harness **spawns a
fresh `rfid_gc_live` subprocess for every trial** and kills it (SIGINT →
graceful disconnect) when the trial duration elapses. The cost is ~1–2 s
of reader startup at the start of each trial; the harness displays
`GO!` only once the reader has actually printed `[GC] Ready`, so you
know when to slide.

## Reading the spreadsheet

`beer-pour-results.xlsx` is created automatically on the first trial
and appended to thereafter — every run writes new rows into the single
`Trials` sheet. The header row is dark-navy / white, frozen so it stays visible
while scrolling, and rows are striped white/off-white for readability.
The `result` verdict is colour-coded as a soft pill (green PASS / amber
SLOW / orange DIRTY / red FAIL), and the `tag_photo` column holds an
embedded thumbnail of the tag used.

If a newer version of this script changes the column layout, the
existing file is renamed to `results_archive_<timestamp>.xlsx` and a
fresh `beer-pour-results.xlsx` is created with the new layout. Your
old data is preserved alongside, untouched.

### Column reference

See the [What the test measures](#what-the-test-measures) table at the
top of this document for the full column list and definitions.

## Tweaks

- Trial duration & verify deadline: `TRIAL_DURATION_S` / `VERIFY_DEADLINE_S` in `beer_pour_logger.py`.
- Add a scenario: append it to `SCENARIOS`, drop a matching `images/scenarios/<name>.png` (optional).
- Add a power level: append to `POWER_LEVELS_MW`. The C binary accepts anything in `1..316` mW.
- Add a tag type: drop `images/tags/<name>.png`. The next session will list it in the tag menu.
- Arbitration thresholds (floor RSSI, RSSI dominance margin, count dominance, min reads): defined in `rfid_gc_live.c`, recompile with `./compile_gc.sh` after editing.

## Troubleshooting

- **`'rfid_gc_live' not found or not executable`** — run `./compile_gc.sh` (or just `./run_test.sh`).
- **`ModuleNotFoundError: openpyxl`** — `sudo apt install -y python3-openpyxl python3-pil` (or `pip3 install --break-system-packages -r requirements.txt`).
- **`Could not connect (code …)`** — check USB, then `sudo chmod 666 /dev/ttyACM0` (or add yourself to `dialout` group and re-login).
- **All trials show `FAIL`** — make sure you're sliding the cup quickly after the `GO!` prompt; the trial clock starts when `[GC] Ready` appears.
- **Lots of `DIRTY` results** — the same EPC is showing up on both antennas (`cross_read_epcs` column tells you which). Likely cause: one tag sitting between antennas, or `GC_RSSI_MARGIN_DB10` in `rfid_gc_live.c` is too low. If you deliberately have one tag per antenna and each EPC only ever appears on its home antenna, you should see **PASS**, not DIRTY.
- **`beer-pour-results.xlsx` write fails** — Excel locks the file when it's open. Close it and retry the trial.
