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

The harness does NOT ask the operator which antenna the cup is going to be
slid over -- in real-world use the system has no advance knowledge of that.
Instead it lets the arbitrator make the call and then characterises the
quality of that call:

- ``winning_antenna``  the antenna with the most attributions across the
                       trial (i.e. the system's "answer")
- ``cross_reads``      number of windows that ALSO put a tag on the OTHER
                       antenna -- a non-zero value means the arbitrator
                       was flip-flopping during the slide, which is what
                       the per-window dominance rule is supposed to
                       prevent.
- ``ttv_s``            seconds from "GO!" to the first window that
                       attributed anything (= time-to-verification)
- ``within_3s``        ``ttv_s <= 3.0`` (the product spec deadline)
- ``result``           PASS / SLOW / DIRTY / FAIL one-word verdict

After every trial the operator is asked whether to keep that row -- press
ENTER (or 'y') to append, 'n' to discard.

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
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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

# A tag must be attributed within this many seconds for the trial to
# pass. Matches the product spec: "verification happens within 3 s of
# the user sliding a tagged cup over the antenna".
VERIFY_DEADLINE_S: float = 3.0

# Photo thumbnail height (pixels) embedded in the workbook. Row height
# is in points; one row comfortably holds an 80-px-tall thumbnail.
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
    def winning_antenna(self) -> Optional[int]:
        """The antenna with the most attributed windows across the trial.
        This is the system's "answer" -- whichever antenna the arbitrator
        decided the cup was closest to. ``None`` if the trial saw no
        attribution on either antenna."""
        n0 = sum(1 for w in self.windows if w.ant0)
        n1 = sum(1 for w in self.windows if w.ant1)
        if n0 == 0 and n1 == 0:
            return None
        return 0 if n0 >= n1 else 1

    @property
    def ttv_s(self) -> Optional[float]:
        """Time (since trial start) of the first window that put any tag
        on any antenna. ``None`` if it never happened."""
        for w in self.windows:
            if w.ant0 or w.ant1:
                return round(w.t_offset_s, 3)
        return None

    @property
    def verified(self) -> bool:
        return self.ttv_s is not None

    @property
    def within_deadline(self) -> bool:
        return self.ttv_s is not None and self.ttv_s <= VERIFY_DEADLINE_S

    @property
    def n_hits_winner(self) -> int:
        if self.winning_antenna is None:
            return 0
        return sum(1 for w in self.windows if w.tags_on(self.winning_antenna))

    @property
    def cross_reads(self) -> int:
        """Number of windows that attributed any tag to the OTHER antenna
        (= not the winner). With the arbitrator's dominance rule working
        as designed this should be 0; anything >0 means the trial was
        flip-flopping between antennas, which is the failure mode we
        care about avoiding in the beer-pour use case."""
        if self.winning_antenna is None:
            return 0
        other = 1 - self.winning_antenna
        return sum(1 for w in self.windows if w.tags_on(other))

    @property
    def clean(self) -> bool:
        return self.cross_reads == 0

    @property
    def best_rssi_winner(self) -> Optional[float]:
        """Strongest (closest-to-zero) RSSI we ever saw on the winning
        antenna across the trial, or None if nothing was attributed."""
        if self.winning_antenna is None:
            return None
        best: Optional[float] = None
        for w in self.windows:
            for _, rssi in w.tags_on(self.winning_antenna):
                if best is None or rssi > best:
                    best = rssi
        return best

    @property
    def result(self) -> str:
        """One-word verdict for the summary column:

        - ``PASS``  verified, within 3 s, no cross-reads
        - ``SLOW``  verified + clean but took longer than 3 s
        - ``DIRTY`` verified but the other antenna also got hits at some
                    point during the trial (= flip-flop, what the arbitrator
                    is meant to prevent)
        - ``FAIL``  never got any attribution
        """
        if not self.verified:
            return "FAIL"
        if not self.clean:
            return "DIRTY"
        if not self.within_deadline:
            return "SLOW"
        return "PASS"


# ─────────────────────────── workbook I/O ───────────────────────────


TRIALS_HEADERS: List[str] = [
    "session_id",
    "trial_num",
    "scenario",
    "power_mw",
    "tag",
    "start_time",
    "duration_s",
    "n_windows",
    "result",
    "verified",
    "ttv_s",
    f"within_{int(VERIFY_DEADLINE_S)}s",
    "winning_antenna",
    "n_hits_winner",
    "cross_reads",
    "clean",
    "best_rssi_winner_dbm",
    "detected_epcs",
    "notes",
    "scenario_photo",
    "tag_photo",
]

TRIALS_WIDTHS: List[int] = [
    16, 10, 32, 9, 14, 22, 11, 11, 9, 10, 9,
    11, 17, 14, 12, 8, 21, 60, 30, 22, 16,
]

WINDOWS_HEADERS: List[str] = [
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

WINDOWS_WIDTHS: List[int] = [16, 10, 32, 9, 14, 11, 11, 36, 24, 36, 24]


# Columns whose values are best displayed centred (counts, numbers, flags).
_CENTERED_TRIALS = {
    "trial_num", "power_mw", "duration_s", "n_windows",
    "result", "verified", "ttv_s",
    f"within_{int(VERIFY_DEADLINE_S)}s",
    "winning_antenna", "n_hits_winner", "cross_reads", "clean",
    "best_rssi_winner_dbm",
}
_CENTERED_WINDOWS = {
    "trial_num", "power_mw", "window_idx", "t_offset_s",
}

# Result-column colour coding. Light fills + readable dark fonts so the
# verdict is visible at a glance when scanning the sheet.
RESULT_FILLS = {
    "PASS":  PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    "SLOW":  PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    "DIRTY": PatternFill(start_color="FFD9B3", end_color="FFD9B3", fill_type="solid"),
    "FAIL":  PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
}
RESULT_FONTS = {
    "PASS":  Font(bold=True, color="006100"),
    "SLOW":  Font(bold=True, color="9C5700"),
    "DIRTY": Font(bold=True, color="9C5700"),
    "FAIL":  Font(bold=True, color="9C0006"),
}


def _apply_header_style(ws: Any) -> None:
    """Bold white text on a dark-blue fill, centred, wrapped, with a
    1-pixel border. Also freezes the header row so it stays visible
    when scrolling."""
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="305496", end_color="305496",
                              fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center",
                             wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def _style_data_row(ws: Any, row: int, headers: List[str],
                    centered_cols: set, result_col: Optional[int] = None,
                    result_value: Optional[str] = None) -> None:
    thin_border = Border(
        left=Side(style="thin", color="E0E0E0"),
        right=Side(style="thin", color="E0E0E0"),
        top=Side(style="thin", color="E0E0E0"),
        bottom=Side(style="thin", color="E0E0E0"),
    )
    centered = Alignment(horizontal="center", vertical="center")
    left     = Alignment(horizontal="left",   vertical="center",
                         wrap_text=True)

    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx)
        cell.border = thin_border
        cell.alignment = centered if name in centered_cols else left

    if result_col is not None and result_value in RESULT_FILLS:
        cell = ws.cell(row=row, column=result_col)
        cell.fill = RESULT_FILLS[result_value]
        cell.font = RESULT_FONTS[result_value]
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=False)


def ensure_workbook() -> None:
    """Create ``results.xlsx`` if it doesn't exist. If an existing file
    has different headers (older schema from a previous version of this
    script), back it up to ``results.xlsx.bak`` and start fresh so the
    new schema isn't silently misaligned with old rows."""
    if RESULTS_XLSX.exists():
        try:
            wb = load_workbook(RESULTS_XLSX)
            if "Trials" in wb.sheetnames:
                trials = wb["Trials"]
                existing = [
                    trials.cell(row=1, column=i + 1).value
                    for i in range(len(TRIALS_HEADERS))
                ]
                if existing != TRIALS_HEADERS:
                    bak = RESULTS_XLSX.with_suffix(".xlsx.bak")
                    print(
                        f"NOTE: existing {RESULTS_XLSX.name} has an older "
                        f"schema; backing up to {bak.name} and starting fresh."
                    )
                    if bak.exists():
                        bak.unlink()
                    RESULTS_XLSX.rename(bak)
        except Exception:
            pass

    if RESULTS_XLSX.exists():
        # Headers match -- just re-apply formatting in case it was stripped.
        wb = load_workbook(RESULTS_XLSX)
        for s in ("Trials", "Windows"):
            if s in wb.sheetnames:
                _apply_header_style(wb[s])
        wb.save(RESULTS_XLSX)
        return

    wb = Workbook()
    trials = wb.active
    trials.title = "Trials"
    trials.append(TRIALS_HEADERS)
    for col, w in enumerate(TRIALS_WIDTHS, start=1):
        trials.column_dimensions[get_column_letter(col)].width = w
    _apply_header_style(trials)

    windows = wb.create_sheet("Windows")
    windows.append(WINDOWS_HEADERS)
    for col, w in enumerate(WINDOWS_WIDTHS, start=1):
        windows.column_dimensions[get_column_letter(col)].width = w
    _apply_header_style(windows)

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
    result_col   = TRIALS_HEADERS.index("result") + 1
    next_row     = trials.max_row + 1

    trials.append([
        session_id,
        t.trial_num,
        t.scenario,
        t.power_mw,
        t.tag,
        t.start_time.strftime("%Y-%m-%d %H:%M:%S"),
        round(t.duration_s, 2),
        t.n_windows,
        t.result,
        "yes" if t.verified else "no",
        t.ttv_s if t.ttv_s is not None else "",
        "yes" if t.within_deadline else "no",
        t.winning_antenna if t.winning_antenna is not None else "",
        t.n_hits_winner,
        t.cross_reads,
        "yes" if t.clean else "no",
        t.best_rssi_winner if t.best_rssi_winner is not None else "",
        ", ".join(t.detected_epcs),
        "",
        "",
        "",
    ])
    trials.row_dimensions[next_row].height = ROW_HEIGHT_PT

    _style_data_row(trials, next_row, TRIALS_HEADERS,
                    _CENTERED_TRIALS,
                    result_col=result_col, result_value=t.result)

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
        _style_data_row(windows, windows.max_row, WINDOWS_HEADERS,
                        _CENTERED_WINDOWS)

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
                    f"either antenna now ({TRIAL_DURATION_S:.1f} s)."
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

    if not reader_ready.wait(timeout=15.0):
        print(f"[Trial #{trial_num}] ERROR: reader never became ready; "
              f"check USB / power and try again.")
        _shutdown(proc)
        t.join(timeout=2)
        return TrialResult(
            scenario=scenario, power_mw=power_mw, tag=tag,
            trial_num=trial_num,
            start_time=dt.datetime.now(), duration_s=0.0,
        )

    time.sleep(TRIAL_DURATION_S)
    duration_s = time.perf_counter() - start_time_holder[0][1]

    _shutdown(proc)
    t.join(timeout=2)

    return TrialResult(
        scenario=scenario,
        power_mw=power_mw,
        tag=tag,
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


def _prompt_save() -> bool:
    """Ask whether to keep this trial. ENTER (or 'y'/'yes') keeps it;
    'n'/'no' discards. Repeats on anything else; Ctrl-C / EOF discards."""
    while True:
        try:
            ans = input("    Save this trial? [Y/n]: ").strip().lower()
        except (KeyboardInterrupt, EOFError):
            return False
        if ans in {"", "y", "yes"}:
            return True
        if ans in {"n", "no"}:
            return False
        print("    Type 'y' (or ENTER) to save, 'n' to discard.")


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

    scenario_idx = _pick("scenarios", SCENARIOS)
    power_idx    = _pick("power levels (mW)", [str(p) for p in POWER_LEVELS_MW])
    tag_idx      = _pick("tag types", tags_available)

    # Per-(scenario, power, tag) trial counters so each unique
    # combination has its own 1, 2, 3 sequence in the spreadsheet.
    counters: Dict[Tuple[str, int, str], int] = {}

    try:
        while True:
            scenario = SCENARIOS[scenario_idx]
            power_mw = POWER_LEVELS_MW[power_idx]
            tag      = tags_available[tag_idx]
            key      = (scenario, power_mw, tag)

            print(
                f"\n[{scenario} | {power_mw} mW | {tag}]"
            )
            print(
                "  ENTER = start trial   'p' = power   's' = scenario   "
                "'t' = tag   'q' = quit"
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
            if choice != "":
                # Anything other than ENTER / known shortcut: re-prompt.
                continue

            tentative_num = counters.get(key, 0) + 1
            result = run_one_trial(
                scenario=scenario,
                power_mw=power_mw,
                tag=tag,
                trial_num=tentative_num,
            )

            # Single-line verdict, ASCII-only so it renders the same in
            # all terminals / log files.
            if result.verified:
                win = result.winning_antenna
                rssi = (f"{result.best_rssi_winner:.1f} dBm"
                        if result.best_rssi_winner is not None else "?")
                print(
                    f"[Trial #{tentative_num}] {result.result} -- "
                    f"verified in {result.ttv_s:.2f} s, "
                    f"winner ant{win} ({result.n_hits_winner}/{result.n_windows} windows), "
                    f"{result.cross_reads} cross-read(s), "
                    f"best RSSI {rssi}"
                )
            else:
                print(
                    f"[Trial #{tentative_num}] {result.result} -- "
                    f"no attribution in {result.n_windows} window(s)"
                )

            if not _prompt_save():
                print("    discarded (not saved).")
                continue

            # Confirmed save -> commit the per-key counter and append.
            counters[key] = tentative_num
            try:
                append_trial(session_id, result)
                print(f"    logged to {RESULTS_XLSX.name}")
            except Exception as exc:
                print(f"    ERROR writing to {RESULTS_XLSX.name}: {exc}")
                print("    (trial was NOT saved -- fix the issue and retry)")
                # Roll the counter back so the next save reuses this num.
                counters[key] = tentative_num - 1
    except KeyboardInterrupt:
        print("\nBye.")
        return 0


if __name__ == "__main__":
    sys.exit(main())
