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

For each trial the harness records:

| Metric                       | What it means                                                                       |
|------------------------------|-------------------------------------------------------------------------------------|
| `result`                     | One-word verdict: **PASS** (verified ≤3 s, no cross-reads) / **SLOW** (verified + clean but >3 s) / **DIRTY** (verified but the other antenna also got hits) / **FAIL** (no attribution at all). |
| `verified`                   | Did any antenna ever report a tag during the trial?                                 |
| `ttv_s`                      | Time to verification — seconds from "GO!" to the first window with any attribution. |
| `within_3s`                  | `ttv_s ≤ 3.0` — the product-spec deadline.                                         |
| `winning_antenna`            | `0` or `1` — the antenna with the most attributed windows (= system's answer).      |
| `n_hits_winner`              | How many of the trial's windows attributed to the winning antenna (consistency).    |
| `cross_reads`                | Windows that attributed any tag to the OTHER antenna. **0 means the arbitrator held one consistent decision**; anything >0 means it flip-flopped, which is what the per-window dominance rule is meant to prevent. |
| `clean`                      | `yes` iff `cross_reads == 0`.                                                       |
| `best_rssi_winner_dbm`       | Strongest (closest-to-zero) RSSI on the winning antenna across the trial.           |
| `detected_epcs`              | Every distinct EPC seen during the trial, regardless of antenna.                    |

Each trial also stores the full per-window trace in a second sheet
(`Windows`) so you can drill in: see exactly which antenna got which
EPC at every 1-second decision window.

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

After every trial the harness prints a verdict line and asks
**`Save this trial? [Y/n]`** — press ENTER (or `y`) to append the row,
`n` to discard it. The per-`(scenario, power, tag)` trial counter only
advances when you choose to save, so the spreadsheet stays sequential
even if you re-do a trial that didn't go well (e.g. you slid the cup
too slowly).

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
| `beer_pour_logger.py`      | Python test harness. Drives `rfid_gc_live` one trial at a time, captures every window, computes the metrics above, appends to `results.xlsx` (with operator confirmation). |
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
Results : /home/stratela/Test-2/results.xlsx
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
[Trial #1] PASS -- verified in 1.00 s, winner ant0 (4/5 windows), 0 cross-read(s), best RSSI -58.3 dBm
    Save this trial? [Y/n]: <ENTER>
    logged to results.xlsx
```

The reader's sweep lines are forwarded verbatim to your terminal (same
ANSI-coloured output as running `rfid_gc_live` by hand), and at the end
the harness prints a single-line verdict.

### Verdict short-codes

| Verdict   | Meaning |
|-----------|---------|
| **PASS**  | `verified == yes`, `within_3s == yes`, `cross_reads == 0`. The happy path. |
| **SLOW**  | Verified + clean, but `ttv_s > 3.0`. The arbitrator got it right but missed the 3-second deadline (e.g. the operator hadn't actually placed the cup yet). |
| **DIRTY** | Verified, but the other antenna also got attributed at some point. Investigate: physical alignment / arbitration thresholds. |
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

`results.xlsx` is created automatically on the first trial and appended
to thereafter. The header row is frozen, the verdict (`result`) is
colour-coded for at-a-glance scanning, and the photos are embedded as
thumbnails next to each row.

If a previous version of this script left a `results.xlsx` with a
different column layout, the harness automatically backs it up to
`results.xlsx.bak` and starts a fresh file with the new schema.

### Sheet `Trials` — one row per trial

| Column                       | Description |
|------------------------------|-------------|
| `session_id`                 | Timestamp of the session (e.g. `20260527-143200`). |
| `trial_num`                  | 1-based counter, **independent per (scenario, power, tag) combination**. Only advances when you confirm-save the trial. |
| `scenario`                   | Name from `SCENARIOS`. |
| `power_mw`                   | TX power for this trial. |
| `tag`                        | Tag-name (file in `images/tags/`). |
| `start_time`                 | Wall-clock at "GO!". |
| `duration_s`                 | Actual elapsed time of the trial. |
| `n_windows`                  | Number of 1-second decision windows the C binary emitted during the trial. |
| `result`                     | PASS / SLOW / DIRTY / FAIL (colour-coded). |
| `verified`                   | `yes` / `no`. |
| `ttv_s`                      | Time to verification (s). Blank if not verified. |
| `within_3s`                  | `yes` / `no`. |
| `winning_antenna`            | `0` / `1` / blank. |
| `n_hits_winner`              | Windows that attributed to the winning antenna. |
| `cross_reads`                | Windows that attributed to the OTHER antenna. |
| `clean`                      | `yes` iff `cross_reads == 0`. |
| `best_rssi_winner_dbm`       | Strongest RSSI on the winning antenna across the trial. |
| `detected_epcs`              | All unique EPCs detected, comma-separated. |
| `notes`                      | Free-text — type here in Excel after a trial. |
| `scenario_photo`             | Embedded thumbnail of `images/scenarios/<scenario>.png` (if present). |
| `tag_photo`                  | Embedded thumbnail of `images/tags/<tag>.png`. |

### Sheet `Windows` — one row per 1-second decision window

| Column           | Description |
|------------------|-------------|
| `session_id`     | Matches the trials sheet. |
| `trial_num`      | Matches the trials sheet. |
| `scenario`       | Matches the trials sheet. |
| `power_mw`       | Matches the trials sheet. |
| `tag`            | Matches the trials sheet. |
| `window_idx`     | 1-based index of this window within the trial. |
| `t_offset_s`     | Seconds from "GO!" to this window's print line. |
| `ant0_epcs`      | Pipe-separated EPCs the arbitrator attributed to antenna 0. |
| `ant0_rssis_dbm` | Pipe-separated max RSSIs (winning value) per EPC, in dBm. |
| `ant1_epcs`      | Pipe-separated EPCs attributed to antenna 1. |
| `ant1_rssis_dbm` | Pipe-separated RSSIs for antenna 1. |

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
- **Lots of `DIRTY` results** — the arbitrator is letting the opposite antenna through. Likely cause: a tag is sitting roughly equidistant between both antennas, or `GC_RSSI_MARGIN_DB10` in `rfid_gc_live.c` is too low for your physical layout.
- **`results.xlsx` write fails** — Excel locks the file when it's open. Close it and retry the trial.
