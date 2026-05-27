# Beer-Pour RFID Verification Test

Test harness for a **beer-pour-machine** where a tagged cup is slid over
one of two antennas. For every cup-slide ("trial") the harness drives the
arbitrated dual-antenna scanner (`rfid_gc_live`), measures how quickly the
correct antenna verifies the tag, watches for any cross-reads on the other
antenna, and appends one row to a master Excel workbook with the scenario
and tag photos embedded next to it.

## What the test measures

For each trial the harness records:

| Metric                       | What it means                                                                       |
|------------------------------|-------------------------------------------------------------------------------------|
| `verified`                   | Did the expected antenna ever report a tag during the trial?                        |
| `ttv_s`                      | Time to verification — seconds from "GO!" to the first window with a tag on that antenna. |
| `within_3s`                  | `ttv_s <= 3.0` — the product-spec deadline.                                         |
| `cross_reads`                | Number of decision windows that attributed any tag to the **other** antenna. With the arbitrator working as designed this should be `0`. Anything >0 is a false attribution to investigate. |
| `n_hits_correct_ant`         | How many of the trial's windows saw a tag on the correct antenna (consistency).     |
| `best_rssi_correct_ant_dbm`  | Strongest (closest-to-zero) RSSI on the correct antenna across the trial.           |
| `detected_epcs`              | Every distinct EPC seen during the trial, regardless of antenna.                    |

Each trial also stores the full per-window trace in a second sheet
(`Windows`) so you can drill in: see exactly which antenna saw what at
each 1-second window.

## Test axes

The four axes are selected from menus and persist across trials until you
change them with `p` / `s` / `t` / `a`.

| Axis              | Values                                                                       |
|-------------------|------------------------------------------------------------------------------|
| Scenario          | `drip_tray_empty_cup_empty`, `drip_tray_half_full_cup_empty` (editable in `SCENARIOS` in `beer_pour_logger.py`) |
| Power (mW)        | `30`, `175`, `316` (editable in `POWER_LEVELS_MW`)                            |
| Tag type          | Auto-discovered from `images/tags/*.png`. Drop e.g. `images/tags/foam.png` to register a new tag. |
| Expected antenna  | `0` (Source_0) or `1` (Source_1) — which antenna you'll slide the cup over.   |

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
| `beer_pour_logger.py`      | Python test harness. Drives `rfid_gc_live` one trial at a time, captures every window, computes the metrics above, appends to `results.xlsx`. |
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

Which antenna will you slide the cup over? [0/1]: 0

[drip_tray_empty_cup_empty | 30 mW | foam | ant0]
  ENTER = start trial   'p' = power   's' = scenario   't' = tag   'a' = antenna   'q' = quit
> <ENTER>

[Trial #1] Launching reader at 30 mW... (reader takes ~1-2 s to come up)
    ===== Dual-Antenna RFID Live Scanner (arbitrated) =====
    [GC] Reader: R3100C  Serial: ...
    [GC] Ready. Empty sweeps print []. Tagged sweeps prepend [TX …]. Ctrl+C to stop.
[Trial #1] GO! Slide the foam cup over antenna 0 now (5.0 s).
    [TX=30 mW] [(0)(-58.3) E2801160600002054E1A1234,                                      ]
    [TX=30 mW] [(0)(-58.4) E2801160600002054E1A1234,                                      ]
    [TX=30 mW] [(0)(-58.5) E2801160600002054E1A1234,                                      ]
    []
    [TX=30 mW] [(0)(-58.6) E2801160600002054E1A1234,                                      ]
[Trial #1] DONE -- verified in 1.00 s [OK], 0 cross-read(s), 5 window(s), best RSSI -58.3 dBm
    logged to results.xlsx
```

The reader's sweep lines are forwarded verbatim to your terminal (same
ANSI-coloured output as running `rfid_gc_live` by hand), and at the end
the harness prints a single-line verdict.

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
to thereafter. Two sheets:

### Sheet `Trials` — one row per trial

| Column                       | Description |
|------------------------------|-------------|
| `session_id`                 | Timestamp of the session (e.g. `20260527-143200`). |
| `trial_num`                  | 1-based counter, **independent per (scenario, power, tag, antenna) combination**. |
| `scenario`                   | Name from `SCENARIOS`. |
| `power_mw`                   | TX power for this trial. |
| `tag`                        | Tag-name (file in `images/tags/`). |
| `expected_antenna`           | `0` or `1`. |
| `start_time`                 | Wall-clock at "GO!". |
| `duration_s`                 | Actual elapsed time of the trial. |
| `n_windows`                  | Number of 1-second windows the C binary emitted during the trial. |
| `verified`                   | `yes` / `no` — did the expected antenna ever attribute a tag? |
| `ttv_s`                      | Time to verification (s). Blank if not verified. |
| `within_3s`                  | `yes` / `no` — did `ttv_s <= 3.0`? |
| `cross_reads`                | Windows that attributed any tag to the other antenna. |
| `n_hits_correct_ant`         | Windows that attributed to the expected antenna. |
| `best_rssi_correct_ant_dbm`  | Strongest RSSI on the expected antenna across the trial. |
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
- **All trials show `NOT VERIFIED`** — make sure you're sliding the cup quickly after the `GO!` prompt; the trial clock starts when `[GC] Ready` appears.
- **`cross_reads > 0` on every trial** — the arbitrator is letting the opposite antenna through. Likely cause: a tag is sitting roughly equidistant between both antennas, or `GC_RSSI_MARGIN_DB10` in `rfid_gc_live.c` is too low for your physical layout.
- **`results.xlsx` write fails** — Excel locks the file when it's open. Close it and retry the trial.
