"""Beer-pour-machine RFID verification test logger.

Drives the existing ``rfid_gc_live`` C binary one trial at a time and records
every trial to a single master Excel workbook (``results.xlsx``).

The C binary is the arbitrated dual-antenna scanner from ``rfid_gc_live.c``:
it runs both antennas at maximum scan rate, accumulates per-EPC stats for one
decision window, and at every window boundary emits one line on stdout:

- ``[]`` when no tag was decisively attributed in that window
- ``[TX=N mW] [(0)(-XX.X) EPC, (1)(-YY.Y) EPC]`` otherwise (ANSI-coloured),
  with antenna 0 in the left slot, antenna 1 in the right slot. Empty slots
  render as pure whitespace so the comma column never shifts.

The arbitration logic (zero-cross-read guarantee) is in the C binary; this
script is purely a measurement harness. For every trial it:

1. Spawns a fresh ``rfid_gc_live <power_mw>`` subprocess so cross-trial
   stats can't leak between cups.
2. Waits for ``[GC] Ready`` -- that's the cue to slide the cup. The moment
   we see that line is ``t = 0`` for the trial.
3. Streams the binary's output to the operator's terminal (so they still
   see the live ``[TX=...] [...]`` lines, exactly like running the binary
   by hand) AND parses each window line in the background.
4. After ``TRIAL_DURATION_S`` seconds, sends SIGINT so the C binary exits
   cleanly (CAENRFID_Disconnect()), then computes:

   - ``ttv_s``       time of the first window that put any tag on the
                     EXPECTED antenna (= "verification time")
   - ``within_3s``   ttv_s <= 3.0
   - ``cross_reads`` number of windows that put a tag on the OTHER antenna
                     (these are the false attributions the arbitrator is
                     meant to prevent)

5. Appends one row to ``Trials`` (with the scenario + tag photos embedded
   as thumbnails) and one row per window to ``Windows``.

The four test axes are picked from menus and kept across trials until you
change them ('p' = power, 's' = scenario, 't' = tag, 'a' = antenna,
'q' = quit). Tag types are discovered from ``images/tags/*.png`` so adding
a new tag is "drop a photo in the folder" -- nothing to edit in code.

Usage::

    python3 beer_pour_logger.py
"""

from __future__ import annotations

import datetime as dt
import os
import re
import signal
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage


# ────────────────────────── paths & config ──────────────────────────


SCRIPT_DIR     = Path(__file__).resolve().parent
RFID_BINARY    = SCRIPT_DIR / "rfid_gc_live"
IMAGES_DIR     = SCRIPT_DIR / "images"
TAGS_DIR       = IMAGES_DIR / "tags"
SCENARIOS_DIR  = IMAGES_DIR / "scenarios"
RESULTS_XLSX   = SCRIPT_DIR / "results.xlsx"
THUMB_DIR      = SCRIPT_DIR / ".thumbs"

# Fixed set of physical scenarios under test on the beer-pour machine.
# Extend this list (and drop a matching <name>.png into images/scenarios/)
# when you add a scenario; everything else picks it up automatically.
SCENARIOS: List[str] = [
    "drip_tray_empty_cup_empty",
    "drip_tray_half_full_cup_empty",
]

# Power levels (mW) we sweep through. Passed verbatim as the positional
# argument to rfid_gc_live.
POWER_LEVELS_MW: List[int] = [30, 175, 316]

# The reader's decision window is 1000 ms, so each trial captures
# roughly TRIAL_DURATION_S window lines. 5 s gives 5 windows -- enough
# to comfortably detect a tag that needs to verify in 3 s and still
# observe a couple of windows of behaviour after that.
TRIAL_DURATION_S: float = 5.0

# A tag must be attributed to the expected antenna within this many
# seconds for the trial to pass. Matches the product spec: "verification
# happens within 3 s of the user sliding a tagged cup over the antenna".
VERIFY_DEADLINE_S: float = 3.0

# Embedded-thumbnail height for the photo columns.
THUMB_HEIGHT_PX  = 80
ROW_HEIGHT_PT    = 64


# ──────────────────────── output line parsing ───────────────────────
#
# rfid_gc_live emits its sweep lines with ANSI colour codes (cyan for the
# [TX=...] prefix, yellow for the (N) antenna labels, green/red for the
# EPCs). We strip the colours first, then run a small set of regexes on
# the plain text. The line shapes after stripping are:
#
#   "[GC] Ready. Empty sweeps print []. ..."         -> start of trial
#   "[]"                                              -> empty window
#   "[TX=30 mW] [(0)(-58.3) EPC0  ,   (1)(-61.7) EPC1  ]"
#   "[TX=30 mW] [(0)(-58.3) EPC0  ,                                  ]"
#   "[TX=30 mW] [                  ,   (1)(-61.7) EPC1               ]"
#
# Slot padding is just trailing whitespace, so strip() on each slot string
# is enough; the comma is the only separator.

ANSI_RE          = re.compile(r"\x1b\[[0-9;]*m")
READY_RE         = re.compile(r"\[GC\]\s+Ready\.")
EMPTY_WINDOW_RE  = re.compile(r"^\s*\[\]\s*$")
WINDOW_RE        = re.compile(
    r"^\s*\[TX=(?P<power>\d+)\s*mW\]\s+\[(?P<inner>.*)\]\s*$"
)
SLOT_HEADER_RE   = re.compile(r"^\((?P<ant>\d+)\)(?P<tags>.+)$")
TAG_PAIR_RE      = re.compile(r"\((?P<rssi>-?\d+\.\d+)\)\s+(?P<epc>[0-9A-Fa-f]+)")


def _parse_slot(slot: str) -> List[Tuple[str, float]]:
    """Return [(EPC, rssi_dbm), ...] for one slot, or [] if it's empty."""
    if not slot:
        return []
    m = SLOT_HEADER_RE.match(slot)
    if not m:
        return []
    return [
        (epc, float(rssi))
        for rssi, epc in TAG_PAIR_RE.findall(m.group("tags"))
    ]


@dataclass
class WindowRead:
    """One decision-window line from the C binary, with arrival timestamp.

    ``t_offset_s`` is seconds since the trial's ``[GC] Ready`` line."""

    idx: int
    t_offset_s: float
    raw: str
    is_empty: bool
    ant0: List[Tuple[str, float]] = field(default_factory=list)
    ant1: List[Tuple[str, float]] = field(default_factory=list)

    def tags_on(self, antenna: int) -> List[Tuple[str, float]]:
        return self.ant0 if antenna == 0 else self.ant1


def parse_window_line(idx: int, t_offset_s: float, line: str) -> Optional[WindowRead]:
    """Turn one stdout line from rfid_gc_live into a WindowRead, or None
    if the line isn't a sweep line at all (e.g. the Ready banner)."""
    clean = ANSI_RE.sub("", line).rstrip("\r\n")
    if EMPTY_WINDOW_RE.match(clean):
        return WindowRead(idx=idx, t_offset_s=t_offset_s, raw=clean, is_empty=True)
    m = WINDOW_RE.match(clean)
    if not m:
        return None
    inner = m.group("inner")
    if "," not in inner:
        return None
    left, right = inner.split(",", 1)
    return WindowRead(
        idx=idx,
        t_offset_s=t_offset_s,
        raw=clean,
        is_empty=False,
        ant0=_parse_slot(left.strip()),
        ant1=_parse_slot(right.strip()),
    )


# ─────────────────────────── data types ────────────────────────────


@dataclass
class TrialResult:
    scenario: str
    power_mw: int
    tag: str
    expected_antenna: int            # 0 or 1
    trial_num: int
    start_time: dt.datetime          # wall-clock at [GC] Ready
    duration_s: float                # actual measured trial length
    windows: List[WindowRead] = field(default_factory=list)

    # ----- derived metrics ---------------------------------------------------

    @property
    def n_windows(self) -> int:
        return len(self.windows)

    @property
    def detected_epcs(self) -> List[str]:
        seen: List[str] = []
        for w in self.windows:
            for epc, _ in w.ant0 + w.ant1:
                if epc not in seen:
                    seen.append(epc)
        return seen

    @property
    def ttv_s(self) -> Optional[float]:
        """Time (since trial start) of the first window with ANY tag on
        the expected antenna. ``None`` if it never happened."""
        for w in self.windows:
            if w.tags_on(self.expected_antenna):
                return round(w.t_offset_s, 3)
        return None

    @property
    def verified(self) -> bool:
        return self.ttv_s is not None

    @property
    def within_deadline(self) -> Optional[bool]:
        if self.ttv_s is None:
            return False
        return self.ttv_s <= VERIFY_DEADLINE_S

    @property
    def n_hits_correct_ant(self) -> int:
        return sum(1 for w in self.windows if w.tags_on(self.expected_antenna))

    @property
    def cross_reads(self) -> int:
        """Number of windows where ANY tag was attributed to the antenna
        the cup is NOT meant to be on. With the arbitrator working as
        designed this should be 0 -- every >0 value is a false attribution
        we want to investigate."""
        other = 1 - self.expected_antenna
        return sum(1 for w in self.windows if w.tags_on(other))

    @property
    def best_rssi_correct_ant(self) -> Optional[float]:
        """Strongest (closest to 0 dBm) RSSI we ever saw on the expected
        antenna across the whole trial, or None if nothing was attributed
        there."""
        best: Optional[float] = None
        for w in self.windows:
            for _, rssi in w.tags_on(self.expected_antenna):
                if best is None or rssi > best:
                    best = rssi
        return best


# ─────────────────────────── workbook I/O ───────────────────────────


TRIALS_HEADERS = [
    "session_id",
    "trial_num",
    "scenario",
    "power_mw",
    "tag",
    "expected_antenna",
    "start_time",
    "duration_s",
    "n_windows",
    "verified",
    "ttv_s",
    f"within_{int(VERIFY_DEADLINE_S)}s",
    "cross_reads",
    "n_hits_correct_ant",
    "best_rssi_correct_ant_dbm",
    "detected_epcs",
    "notes",
    "scenario_photo",
    "tag_photo",
]

TRIALS_WIDTHS = [
    16, 10, 32, 9, 14, 18, 22, 11, 11,
    10, 9, 11, 12, 18, 23, 60, 30, 22, 16,
]

WINDOWS_HEADERS = [
    "session_id",
    "trial_num",
    "scenario",
    "power_mw",
    "tag",
    "window_idx",
    "t_offset_s",
    "ant0_epcs",
    "ant0_rssis_dbm",
    "ant1_epcs",
    "ant1_rssis_dbm",
]

WINDOWS_WIDTHS = [16, 10, 32, 9, 14, 11, 11, 36, 24, 36, 24]


def ensure_workbook() -> None:
    """Create ``results.xlsx`` with the Trials + Windows sheets if it
    doesn't exist; otherwise upgrade missing header columns in place so
    older files keep working when this script grows new metrics."""
    if not RESULTS_XLSX.exists():
        wb = Workbook()
        trials = wb.active
        trials.title = "Trials"
        trials.append(TRIALS_HEADERS)
        for col, w in enumerate(TRIALS_WIDTHS, start=1):
            trials.column_dimensions[get_column_letter(col)].width = w

        windows = wb.create_sheet("Windows")
        windows.append(WINDOWS_HEADERS)
        for col, w in enumerate(WINDOWS_WIDTHS, start=1):
            windows.column_dimensions[get_column_letter(col)].width = w

        wb.save(RESULTS_XLSX)
        return

    wb = load_workbook(RESULTS_XLSX)
    changed = False

    def upgrade(sheet: str, headers: List[str], widths: List[int]) -> None:
        nonlocal changed
        if sheet not in wb.sheetnames:
            return
        ws = wb[sheet]
        for i, h in enumerate(headers, start=1):
            if not ws.cell(row=1, column=i).value:
                ws.cell(row=1, column=i, value=h)
                changed = True
        for col, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col)].width = w

    upgrade("Trials", TRIALS_HEADERS, TRIALS_WIDTHS)
    upgrade("Windows", WINDOWS_HEADERS, WINDOWS_WIDTHS)

    if changed:
        wb.save(RESULTS_XLSX)


def _thumb_for(src_dir: Path, name: str) -> Optional[Path]:
    """Return a cached thumbnail of ``<src_dir>/<name>.png``, regenerating
    it when the source photo has been updated. Returns None if there's
    no source photo on disk (in that case the cell is just left blank)."""
    src = src_dir / f"{name}.png"
    if not src.exists():
        return None
    THUMB_DIR.mkdir(exist_ok=True)
    dst = THUMB_DIR / f"{src_dir.name}__{name}_thumb.png"
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
        return dst
    img = PILImage.open(src)
    img.thumbnail((10_000, THUMB_HEIGHT_PX))
    img.save(dst, "PNG")
    return dst


def append_trial(session_id: str, t: TrialResult) -> None:
    wb = load_workbook(RESULTS_XLSX)
    trials = wb["Trials"]
    windows = wb["Windows"]

    scenario_col = get_column_letter(TRIALS_HEADERS.index("scenario_photo") + 1)
    tag_col      = get_column_letter(TRIALS_HEADERS.index("tag_photo") + 1)
    next_row = trials.max_row + 1

    trials.append([
        session_id,
        t.trial_num,
        t.scenario,
        t.power_mw,
        t.tag,
        t.expected_antenna,
        t.start_time.strftime("%Y-%m-%d %H:%M:%S"),
        round(t.duration_s, 2),
        t.n_windows,
        "yes" if t.verified else "no",
        t.ttv_s if t.ttv_s is not None else "",
        ("yes" if t.within_deadline else "no") if t.verified else "no",
        t.cross_reads,
        t.n_hits_correct_ant,
        t.best_rssi_correct_ant if t.best_rssi_correct_ant is not None else "",
        ", ".join(t.detected_epcs),
        "",
        "",
        "",
    ])
    trials.row_dimensions[next_row].height = ROW_HEIGHT_PT

    s_thumb = _thumb_for(SCENARIOS_DIR, t.scenario)
    if s_thumb is not None:
        img = XLImage(str(s_thumb))
        img.anchor = f"{scenario_col}{next_row}"
        trials.add_image(img)

    t_thumb = _thumb_for(TAGS_DIR, t.tag)
    if t_thumb is not None:
        img = XLImage(str(t_thumb))
        img.anchor = f"{tag_col}{next_row}"
        trials.add_image(img)

    for w in t.windows:
        windows.append([
            session_id,
            t.trial_num,
            t.scenario,
            t.power_mw,
            t.tag,
            w.idx,
            round(w.t_offset_s, 3),
            "|".join(epc for epc, _ in w.ant0),
            "|".join(f"{r:.1f}" for _, r in w.ant0),
            "|".join(epc for epc, _ in w.ant1),
            "|".join(f"{r:.1f}" for _, r in w.ant1),
        ])

    wb.save(RESULTS_XLSX)


# ────────────────────── rfid_gc_live subprocess ─────────────────────


def discover_tags() -> List[str]:
    """Tag types are whatever PNGs live in images/tags/, sorted. Adding
    a new tag = drop ``<name>.png`` in there and restart the harness."""
    if not TAGS_DIR.exists():
        return []
    return sorted(
        p.stem for p in TAGS_DIR.glob("*.png") if not p.name.startswith(".")
    )


def run_one_trial(scenario: str,
                  power_mw: int,
                  tag: str,
                  expected_antenna: int,
                  trial_num: int) -> TrialResult:
    """Launch one ``rfid_gc_live`` subprocess for a single trial. The
    operator should be ready to slide the cup the moment the harness
    prints ``GO!``; that moment is t = 0 for the trial."""
    print(
        f"\n[Trial #{trial_num}] Launching reader at {power_mw} mW... "
        f"(reader takes ~1-2 s to come up)"
    )

    proc = subprocess.Popen(
        [str(RFID_BINARY), str(power_mw)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        bufsize=1,
        universal_newlines=True,
    )

    windows: List[WindowRead] = []
    reader_ready = threading.Event()
    start_time_holder: List[Tuple[dt.datetime, float]] = []  # (wallclock, perf_counter)

    def reader_thread() -> None:
        assert proc.stdout is not None
        win_idx = 0
        for line in proc.stdout:
            sys.stdout.write("    " + line)
            sys.stdout.flush()
            clean = ANSI_RE.sub("", line)
            if not reader_ready.is_set() and READY_RE.search(clean):
                start_time_holder.append((dt.datetime.now(), time.perf_counter()))
                reader_ready.set()
                print(
                    f"[Trial #{trial_num}] GO! Slide the {tag} cup over "
                    f"antenna {expected_antenna} now ({TRIAL_DURATION_S:.1f} s)."
                )
                sys.stdout.flush()
                continue
            if not reader_ready.is_set():
                continue
            t_offset = time.perf_counter() - start_time_holder[0][1]
            win_idx += 1
            w = parse_window_line(win_idx, t_offset, line)
            if w is None:
                win_idx -= 1
                continue
            windows.append(w)

    t = threading.Thread(target=reader_thread, daemon=True)
    t.start()

    # Wait for the reader to come up before starting the trial clock. If
    # it never gets ready (e.g. USB unplugged) we abort after a generous
    # timeout so the harness doesn't hang the operator.
    if not reader_ready.wait(timeout=15.0):
        print(f"[Trial #{trial_num}] ERROR: reader never became ready; "
              f"check USB / power and try again.")
        _shutdown(proc)
        t.join(timeout=2)
        return TrialResult(
            scenario=scenario, power_mw=power_mw, tag=tag,
            expected_antenna=expected_antenna, trial_num=trial_num,
            start_time=dt.datetime.now(), duration_s=0.0,
        )

    # Let the C binary run for the trial budget, then stop it.
    time.sleep(TRIAL_DURATION_S)
    duration_s = time.perf_counter() - start_time_holder[0][1]

    _shutdown(proc)
    t.join(timeout=2)

    return TrialResult(
        scenario=scenario,
        power_mw=power_mw,
        tag=tag,
        expected_antenna=expected_antenna,
        trial_num=trial_num,
        start_time=start_time_holder[0][0],
        duration_s=duration_s,
        windows=windows,
    )


def _shutdown(proc: subprocess.Popen) -> None:
    """Send SIGINT so the C binary's signal handler runs
    CAENRFID_Disconnect() cleanly; escalate if it ignores us."""
    try:
        proc.send_signal(signal.SIGINT)
        proc.wait(timeout=3)
    except Exception:
        try:
            proc.terminate()
            proc.wait(timeout=2)
        except Exception:
            proc.kill()


# ───────────────────────────── menu loop ─────────────────────────────


def _pick(label: str, options: List[str]) -> int:
    while True:
        print(f"\nAvailable {label}:")
        for i, s in enumerate(options, 1):
            print(f"  {i}. {s}")
        choice = input(f"Select {label} [1..{len(options)}]: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return int(choice) - 1
        print(f"  Invalid choice, try again.")


def _pick_antenna() -> int:
    while True:
        choice = input(
            "\nWhich antenna will you slide the cup over? [0/1]: "
        ).strip()
        if choice in {"0", "1"}:
            return int(choice)
        print("  Must be 0 or 1.")


def main() -> int:
    if not RFID_BINARY.is_file() or not os.access(RFID_BINARY, os.X_OK):
        print(f"ERROR: '{RFID_BINARY.name}' not found or not executable.")
        print("Build it first:  ./compile_gc.sh   (or ./run_test.sh which does it for you)")
        return 1

    tags_available = discover_tags()
    if not tags_available:
        print(f"ERROR: no tag photos found in {TAGS_DIR}.")
        print("Drop a PNG named e.g. 'foam.png' in there to register a tag type.")
        return 1

    ensure_workbook()

    session_id = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    print(f"=== Beer-pour RFID verification test session {session_id} ===")
    print(f"Results : {RESULTS_XLSX}")
    print(f"Reader  : {RFID_BINARY.name}")
    print(f"Trial   : {TRIAL_DURATION_S:.1f} s  (verify deadline: "
          f"{VERIFY_DEADLINE_S:.1f} s)")

    scenario_idx     = _pick("scenarios", SCENARIOS)
    power_idx        = _pick("power levels (mW)", [str(p) for p in POWER_LEVELS_MW])
    tag_idx          = _pick("tag types", tags_available)
    expected_antenna = _pick_antenna()

    # Per-(scenario, power, tag, antenna) trial counters so each unique
    # combination has its own 1, 2, 3 sequence in the spreadsheet.
    counters: Dict[Tuple[str, int, str, int], int] = {}

    try:
        while True:
            scenario = SCENARIOS[scenario_idx]
            power_mw = POWER_LEVELS_MW[power_idx]
            tag      = tags_available[tag_idx]
            key      = (scenario, power_mw, tag, expected_antenna)

            print(
                f"\n[{scenario} | {power_mw} mW | {tag} | ant{expected_antenna}]"
            )
            print(
                "  ENTER = start trial   'p' = power   's' = scenario   "
                "'t' = tag   'a' = antenna   'q' = quit"
            )
            try:
                choice = input("> ").strip().lower()
            except (KeyboardInterrupt, EOFError):
                print("\nBye.")
                return 0

            if choice == "q":
                print("Bye.")
                return 0
            if choice == "p":
                power_idx = _pick("power levels (mW)",
                                  [str(p) for p in POWER_LEVELS_MW])
                continue
            if choice == "s":
                scenario_idx = _pick("scenarios", SCENARIOS)
                continue
            if choice == "t":
                tag_idx = _pick("tag types", tags_available)
                continue
            if choice == "a":
                expected_antenna = _pick_antenna()
                continue

            counters[key] = counters.get(key, 0) + 1
            result = run_one_trial(
                scenario=scenario,
                power_mw=power_mw,
                tag=tag,
                expected_antenna=expected_antenna,
                trial_num=counters[key],
            )

            # Operator-readable verdict line. Use plain ASCII so it
            # renders the same way on the Pi terminal and in CI logs.
            if result.verified:
                deadline = "OK" if result.within_deadline else "LATE"
                print(
                    f"[Trial #{result.trial_num}] DONE -- verified in "
                    f"{result.ttv_s:.2f} s [{deadline}], "
                    f"{result.cross_reads} cross-read(s), "
                    f"{result.n_windows} window(s), "
                    f"best RSSI {result.best_rssi_correct_ant} dBm"
                )
            else:
                print(
                    f"[Trial #{result.trial_num}] DONE -- NOT VERIFIED, "
                    f"{result.cross_reads} cross-read(s), "
                    f"{result.n_windows} window(s)"
                )

            try:
                append_trial(session_id, result)
                print(f"    logged to {RESULTS_XLSX.name}")
            except Exception as exc:
                print(f"    ERROR writing to {RESULTS_XLSX.name}: {exc}")
                print("    (trial was NOT saved -- fix the issue and retry)")
    except KeyboardInterrupt:
        print("\nBye.")
        return 0


if __name__ == "__main__":
    sys.exit(main())
