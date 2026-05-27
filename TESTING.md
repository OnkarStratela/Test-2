# Beer-pour RFID reliability test

Test harness for the dual-antenna beer-pour rig. A user slides a reusable
plastic mug (tagged on the bottom) onto a drip tray; the RFID antennas
sit underneath the tray. We want to know:

1. **Does the reader miss the mug?**  *(miss-read)*
2. **Does a mug ever get attributed to the wrong antenna?**  *(cross-read)*

Both questions are asked across this matrix:

| Dimension     | Values                                                       |
|---------------|--------------------------------------------------------------|
| Tag type      | `foam` (`images/tags/foam.png`) — add more without code changes |
| Mug layout    | `ant0_only`, `ant1_only`, `both`                             |
| TX power      | `30 mW`, `175 mW`, `316 mW`                                  |
| Condition     | `dry` (no liquid), `wet` (beer pouring into mug during trial) |
| Trial length  | 3 s, with a 250 ms decision window → ~12 data points / trial |

The arbitration logic from the production code (`rfid_gc_live.c`) is
preserved exactly: RSSI floor, signal-strength margin, read-count
dominance. This harness just orchestrates *which* power and layout each
trial runs at and computes the higher-level miss/cross-read rates.

---

## Files

| File              | Purpose |
|-------------------|---------|
| `rfid_gc_live.c`  | C reader: arbitrates the two antennas, emits one human line and one machine line (`[GC_WIN] …`) per 250 ms window. |
| `compile_gc.sh`   | Builds `rfid_gc_live` against `SRC/`. |
| `run_gc.sh`       | Raw-reader runner (no test logging) — useful for quick hardware checks. |
| `test_logger.py`  | **The test harness.** Spawns `rfid_gc_live` per trial, parses every `[GC_WIN]` line into a window record, computes per-trial metrics, and appends to `results.xlsx`. |
| `run_test.sh`     | One-shot entry point: dependency checks → compile → launch `test_logger.py`. |
| `requirements.txt`| Python deps: `openpyxl`, `Pillow`. |
| `images/tags/`    | Tag reference photos. Embedded in `results.xlsx` next to each trial row. |
| `images/setups/`  | (Optional) Phone photos of the test rig. |
| `results.xlsx`    | The single master report (auto-created on first run, appended forever). Tracked in git so the Pi and your laptop stay in sync. |
| `SRC/`            | CAEN RFID Light library. Do not modify. |

---

## Hardware

- Raspberry Pi + carrier board (any model with USB; CM4 used during development)
- CAEN R3100C-Lepton3 25 dBm reader on `/dev/ttyACM0`
- Two UHF antennas wired to `Source_0` and `Source_1`, 150 mm centre-to-centre, mounted under the drip tray
- Tagged mug(s) for testing

---

## Setup (Pi)

```bash
cd ~/Test-2/test/Test-2   # adjust to where you pulled the repo

# 1) Build the C reader (also installs the standard build-essential if needed)
chmod +x compile_gc.sh run_test.sh
./compile_gc.sh

# 2) Install Python deps via apt (avoids PEP 668)
sudo apt install -y python3-openpyxl python3-pil

# 3) USB permissions
sudo usermod -a -G dialout "$USER"
# Log out and back in once.
```

If you skip step 2, `./run_test.sh` will offer to run those apt commands
for you on first launch.

---

## Run a test session

```bash
cd ~/Test-2/test/Test-2
./run_test.sh
```

Sample session — picks the tag type, then walks one scenario at a time:

```text
=== Beer-pour RFID test session 20260527-141205 ===
Results: /home/stratela/Test-2/test/Test-2/results.xlsx
Trial duration: 3.0 s    Decision window: 250 ms    Powers: 30 mW, 175 mW, 316 mW

Tag type for this session: foam (Foam tag (CP15710-01))

Available scenarios for tag 'foam':
   1. foam | ant0_only | 30 mW  | dry
   2. foam | ant0_only | 30 mW  | wet
   3. foam | ant0_only | 175 mW | dry
   4. foam | ant0_only | 175 mW | wet
   ...
  18. foam | both      | 316 mW | wet
Select scenario [1-18, 's' to switch tag, 'q' to quit]: 1

------------------------------------------------------------------------
  Scenario: foam | ant0_only | 30 mW | dry
  Setup   : Place ONE mug centred over antenna 0. Antenna 1 must be empty.
  Liquid  : Make sure the drip tray and the mug are completely DRY. No droplets.
------------------------------------------------------------------------

  > Press ENTER once the setup is in place ...
[Trial #1] starting reader for 3.0 s (power=30 mW, window=250 ms) ...
    [GC] Connecting...
    [GC] Reader: A949P  Serial: 12345
    [GC] Ready. One line every 0.250 s: ...
[Trial #1] LIVE — recording 3.0 s of windows now.
    win= 0  ant0=E2801160600002054E1A1234  ant1=-  drop(floor/low/amb)=0/2/0
    win= 1  ant0=E2801160600002054E1A1234  ant1=-  drop(floor/low/amb)=0/0/0
    ...
[Trial #1] DONE — 12 window(s) in 3.0 s

    Total windows   : 12
    Ant0 attribution:  91.67%   inferred EPC: E2801160600002054E1A1234
    Ant1 attribution:   0.00%   inferred EPC: -
    Miss-reads      : 1
    Cross-reads     : 0
    Scan rate 0/1   : 132.4 Hz / 130.1 Hz
    Mean RSSI 0/1   : -52.1 dBm / - dBm
    Dropped         : floor=0  low_count=2  ambiguous=0

  Append this trial to results.xlsx? [Y/n] y
    logged to results.xlsx
  Run another trial in THIS scenario? [y/N] n
```

`'s'` at the scenario menu lets you switch tag type mid-session.
`'q'` (or Ctrl-C) ends cleanly — every accepted trial is already saved.

---

## What gets recorded

A single file `results.xlsx` at the repo root, appended across every
session.

### Sheet `Trials` — one row per 3-second trial

| Column                | Description |
|-----------------------|-------------|
| session_id            | Timestamp of the session (`YYYYMMDD-HHMMSS`). |
| trial_uid             | Unique id encoding scenario + trial number. |
| started_at            | Wall-clock start of the trial. |
| tag_type / layout / power_mw / condition | The four matrix dimensions. |
| expected_ant0 / expected_ant1 | What the layout says each antenna SHOULD see. |
| total_windows         | Number of 250 ms windows recorded (~12 for a 3-s trial). |
| ant0_attribution_%    | % of windows where ant0's slot was non-empty. |
| ant1_attribution_%    | Same for ant1. |
| ant0_inferred_epc     | Most common EPC on ant0 across the trial. |
| ant1_inferred_epc     | Same for ant1. |
| miss_read_count       | Windows where an *expected* antenna was empty. |
| cross_read_count      | Windows where a mug was attributed to the wrong antenna. |
| mean_rate_0_hz / mean_rate_1_hz | Average scans/sec per antenna (~120-140 Hz typical). |
| mean_rssi_ant0_dbm / mean_rssi_ant1_dbm | Average max-RSSI of attributed reads. |
| drop_floor_total / drop_low_total / drop_amb_total | Counts of arbitration rejections (RSSI below floor, too few reads, ambiguous). |
| tag_reference_photo   | Embedded thumbnail of `images/tags/<tag>.png`. |
| notes                 | Free-text — currently blank, edit by hand if you want. |

Cells with `ant0/1_attribution_%` below the threshold *for an antenna
that was supposed to see a mug* are highlighted red. So is any non-zero
`cross_read_count`.

### Sheet `Windows` — one row per 250 ms decision window

The raw data exactly as the C binary emitted it (`[GC_WIN] …`), one row
per window across every trial. Use this for plotting / forensic
analysis. Each row references its trial via `trial_uid`.

---

## Adding a second tag type later

1. Photograph the new tag; drop it as `images/tags/<id>.png` (e.g.
   `images/tags/sticker.png`).
2. Append an entry to `TAG_TYPES` near the top of `test_logger.py`:

   ```python
   TAG_TYPES: List[TagType] = [
       TagType(id="foam", label="Foam tag (CP15710-01)", notes="..."),
       TagType(id="sticker", label="Generic UHF sticker", notes="Paper inlay"),
   ]
   ```

3. Re-run `./run_test.sh`. The session menu now offers the new tag too.

No new sheets, no new columns — the same `results.xlsx` grows.

---

## Syncing `results.xlsx` between the Pi and your laptop

`results.xlsx` is tracked in git. After a session on the Pi:

```bash
cd ~/Test-2
git pull --rebase                       # pick up any code changes first
git add test/Test-2/results.xlsx
git commit -m "Test results $(date +%Y-%m-%d_%H-%M)"
git push
```

On your laptop:

```bash
git pull
```

Don't have `results.xlsx` open in Excel during the pull — Excel locks
the file and the pull will fail.

---

## Tuning

- **Powers / trial length / decision window** — edit the constants at
  the top of `test_logger.py` (`POWERS_MW`, `TRIAL_DURATION_S`,
  `DECISION_WINDOW_MS`). No C recompile needed; the harness passes
  them to `rfid_gc_live` per trial.
- **Arbitration thresholds** (RSSI floor, dominance margin, count
  ratio, min reads) — edit the `GC_*` macros at the top of
  `rfid_gc_live.c`, then `./compile_gc.sh`.
- **Reliability thresholds for the red highlight in Excel** —
  `ATTRIBUTION_RATE_MIN_PCT` / `CROSS_READ_RATE_MAX_PCT` in
  `test_logger.py`.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `rfid_gc_live not found or not executable` | Run `./compile_gc.sh` first. |
| `[GC] ERROR: Could not connect (code -3)` | USB perms — `sudo usermod -a -G dialout $USER`, log out, log in. Or one-shot: `sudo chmod 666 /dev/ttyACM0`. |
| `ModuleNotFoundError: openpyxl` | `sudo apt install -y python3-openpyxl python3-pil`. |
| `error: externally-managed-environment` from pip | That's PEP 668 on Pi OS Bookworm; use apt as above. |
| Reader never reaches LIVE | Reader took longer than 3 s to start. Increase `TRIAL_DURATION_S` to 5 s in `test_logger.py`. |
| All RSSI columns blank | Reader is fine but found no tags. Try 316 mW + the dry condition to verify the hardware before re-running. |
| Cross-reads stay non-zero in `both` layout at high power | The arbitrator's RSSI margin (6 dB by default) isn't enough at that power. Raise `GC_RSSI_MARGIN_DB10` in `rfid_gc_live.c`, recompile, re-test. |
