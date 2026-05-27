# Beer-pour RFID reliability harness

What the test measures (from your spec):

1. **Dry-setup reliability** — drip tray empty, mugs empty.
2. **Power sweep** — `30 / 175 / 316 mW`.
3. **Wet-setup reliability** — beer being poured into the mug for the
   duration of the trial.
4. **Tag types** — start with the foam tag (`CP15710-01`); add more by
   editing `config.yaml`.

For each of those scenarios the harness measures two things you care about:

- **Miss-reads** — an antenna that *should* see a mug attributed nothing
  in a decision window.
- **Cross-reads** — a mug appeared on the *wrong* antenna's slot (either
  in a single-mug layout, or with the wrong EPC in the two-mug layout).

All results are appended to a single Excel file you can open from
Windows: `bench/results/report.xlsx`.

---

## Architecture

```
Windows PC                        Linux box (CAEN reader plugged in)
-----------                       --------------------------------
run_test_matrix.py  -- SSH --->   ./rfid_gc_live <power>
                                    --duration 3
                                    --window-ms 250
                                    --csv /tmp/...csv
                                    --quiet
            <-- SFTP (CSV) ----   /tmp/.trial_*.csv

bench/results/report.xlsx (Summary, Trials, Raw, Tags, Config)
bench/results/session_<ts>/*.csv  (raw CSV per trial, retained)
```

The C code (`rfid_gc_live.c`) gained four backward-compatible flags so the
test harness can drive it deterministically:

| Flag                  | Purpose                                                    |
|-----------------------|------------------------------------------------------------|
| `--duration <sec>`    | Auto-stop after N seconds (no Ctrl+C needed)               |
| `--window-ms <ms>`    | Override the 1000 ms decision window (we use **250 ms**)   |
| `--csv <file>`        | Write one row per decision window in a fixed schema         |
| `--quiet`             | Suppress the colour stdout (CSV is the only output)        |

The CSV now also includes drop counters (`n_dropped_below_floor`,
`n_dropped_low_count`, `n_dropped_ambiguous`) — these used to be silently
thrown away. They let you see *why* a tag was rejected by the arbitrator.

---

## One-time setup

### Linux box (where the reader is)

1. Re-build the binary with the new flags:

   ```bash
   cd ~/Test-2          # path used in config.yaml -> ssh.remote_workdir
   ./compile_gc.sh
   ```

2. Make sure your Windows user can SSH in without typing a password:

   ```bash
   # on the Linux box, append your Windows pubkey to ~/.ssh/authorized_keys
   ```

3. If the reader needs `sudo` for `/dev/ttyACM0`, either:
   - add yourself to `dialout` so it doesn't (recommended), **or**
   - set `ssh.use_sudo: true` in `config.yaml` and configure passwordless
     sudo for the binary.

### Windows PC (where you run the harness)

1. Install Python 3.11+ and from `bench/`:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

2. Edit `bench/config.yaml`:
   - `ssh.host`, `ssh.user`, `ssh.key_path` — your Linux box
   - `ssh.remote_workdir` — where you ran `compile_gc.sh`
   - everything else has sensible defaults

---

## Running a session

```powershell
cd C:\Users\it241\Desktop\Test-2\test\Test-2\bench
.\.venv\Scripts\Activate.ps1
python run_test_matrix.py
```

If you have more than one tag type, you'll be asked which one this
session is for. (You can also pass `--tag-type foam` to skip the prompt.)

The harness then walks the matrix
**layout × power × condition** in this order:

```
ant0_only,  30 mW, dry  -> ant0_only,  30 mW, wet
ant0_only, 175 mW, dry  -> ant0_only, 175 mW, wet
ant0_only, 316 mW, dry  -> ant0_only, 316 mW, wet
ant1_only, 30 mW, dry   -> ... (and so on for all combinations)
both,      30 mW, dry   -> ...
```

For each cell:

```
  Tag type : Foam tag (CP15710-01)
  Layout   : Mug A over antenna 0 only
  Power    : 30 mW
  Condition: Dry (drip tray empty, mug empty)
  Setup    : Place ONE mug centred over antenna 0. Antenna 1 must be empty.
  Liquid   : Make sure the drip tray and the mug are completely DRY.
```

then prompts:

```
  Run a trial for this cell? [Y/n]
  >>> Press ENTER when the setup is in place (and pour, if wet) <<<
```

The 3-second trial runs, the CSV is pulled back, a summary prints:

```
  Trial UID            : foam_ant0_only_30mW_dry_20260527-123412
  Decision windows     : 12
  Ant0 attribution     :  91.67%   inferred EPC: E2801160600002054E1A1234
  Ant1 attribution     :   0.00%   inferred EPC: -
  Miss-reads           : 1
  Cross-reads          : 0
  Scan rate ant0/ant1  : 132.4 Hz / 130.1 Hz
  Mean RSSI ant0/ant1  : -52.1 dBm / - dBm
  Dropped below floor  : 0
  Dropped low count    : 2
  Dropped ambiguous    : 0
  CSV saved to         : results\session_20260527-123405\foam_ant0_only_30mW_dry_20260527-123412.csv

  Append this trial to the report? [Y/n]
  Run another trial in THIS cell? [y/N]
```

You decide live whether to keep each trial and whether to repeat the cell.
That's the "ask after every test run" behaviour you asked for.

### Adding a photo per scenario

After accepting the first trial of a cell, the harness prints the
exact filename it expects for the setup photo, e.g.

```
  If you want to attach a setup photo, drop it as:
    C:\...\bench\photos\setups\foam_ant0_only_30mW_dry.jpg  (or .png / .jpeg)
  Photo ready? [y/N]
```

Drop the phone photo there with that exact filename, then answer `y`.
The photo is embedded into the Summary sheet next to the matching cell.
You can also drop photos into that folder **before** running — the
harness picks them up automatically.

---

## The report

`bench/results/report.xlsx` is the single canonical report. Every session
appends to it; nothing is overwritten.

| Sheet     | What's in it |
|-----------|--------------|
| `Summary` | One row per `(tag_type, layout, power, condition)`. Aggregates across all trials ever recorded for that cell. Cells that violate the thresholds in `config.yaml` are highlighted red. Includes embedded setup photo. |
| `Trials`  | One row per individual trial — the *only* sheet that gets new rows on every run. |
| `Raw`     | One row per 250 ms decision window across every trial. Use this for plotting / detailed analysis. |
| `Tags`    | One row per tag type with notes and an embedded reference image. |
| `Config`  | Last session's metadata: SSH host, decision-window, thresholds, session ID. |

### Adding a new tag type later

1. Photograph the new tag, drop it as `bench/photos/tags/<id>.png`.
2. Add an entry to `tag_types:` in `config.yaml`:

   ```yaml
   - id: sticker
     label: "Generic UHF sticker"
     reference_image: "photos/tags/sticker.png"
     notes: "Paper-substrate dry inlay."
   ```

3. Re-run the harness; the prompt now offers `foam` and `sticker`. Pick
   `sticker`, run your trials, and the same `report.xlsx` grows with the
   new tag-type rows.

---

## Re-using existing CSVs without re-running

If the SSH run worked but the workbook write didn't (e.g. Excel had the
file open and locked it), the per-trial CSVs in
`bench/results/session_<ts>/` are the raw record — nothing is lost. Open
them directly in Excel or re-import with a one-off script: every CSV
follows the same schema documented at the top of `rfid_gc_live.c`.

---

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| SSH connect error | Check `config.yaml` `ssh:` block. Test with `ssh user@host` from PowerShell first. |
| `cannot open CSV file ...` from the binary | `remote_workdir` not writable by the SSH user. |
| Trial returns 0 windows | Reader didn't start fast enough for 3 s. Increase `trial_duration_s` to 5 s. |
| All RSSI columns blank | Reader connected but found no tags (mug too far or below the floor). Try `316` mW with the dry config to verify hardware. |
| Cross-reads stuck non-zero in `both` layout at high power | Expected — the arbitrator's RSSI margin is tight. Raise `GC_RSSI_MARGIN_DB10` in `rfid_gc_live.c`, recompile, re-test. |
