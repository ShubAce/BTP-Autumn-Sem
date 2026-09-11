r"""
=============================================================================
  Automated Multi-Root Video-Processing System  —  BTP IIT Kharagpur
=============================================================================
  Processes ALL highway folders found inside a top-level video directory.

  Folder hierarchy expected inside EACH highway folder:
    <highway>/          e.g.  NH-19 1st and 2nd May/
      <section>/              LHS/  or  RHS/  (or any name)
        <date>/               1st May/  2nd May/  3rd May/ …
          <session>/          0000-0100/  0100-0200/  …   (HHMM-HHMM)
            *.TS / *.mp4 …    actual video files

  The "messy" outer layers (highway name, section, date) can vary in name
  and depth.  The ONLY strict rule is that the leaf folders holding videos
  must be named in HHMM-HHMM format (the time-session folders).

  Usage (run from f:\BTP\Autmn Sem\):
    python automate_all_videos.py
       --video-root   "f:\BTP\video"
       --output-dir   "f:\BTP\Autmn Sem\outputs_all"
       --model        yolo11n
       --conf         0.20  --imgsz 640  --skip-frames 2

  Features:
    * Auto-discovers every highway folder under --video-root
    * Handles arbitrary nesting above the session (HHMM-HHMM) folders
    * Per (highway / section / date) calibration — interactive on first
      video of that group, headless thereafter
    * Single unified Master Excel:
        outputs_all/master_dataset_report.xlsx
          Sheet 1: Processing Manifest  (one row per video, all highways)
          Sheet 2: All Vehicles Raw Data (all detection rows)
    * Per-highway checkpoint JSON for fault-tolerant resume
    * Ctrl-C safe: saves current state before exiting
=============================================================================
"""

from __future__ import annotations

import argparse
import csv
import datetime
import json
import os
import re
import signal
import subprocess
import sys
import time
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    XLSX_OK = True
except ImportError:
    XLSX_OK = False
    print("[WARN] openpyxl not available — Excel output disabled")

VIDEO_EXTS = {".ts", ".mp4", ".avi", ".mkv", ".mov", ".mpeg", ".mpg"}

# Regex that matches time-session folder names like 0000-0100, 1300-1400, 2300-2400
SESSION_RE = re.compile(r"^\d{4}-\d{4}$")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def natural_key(text: str) -> list:
    """Natural alphanumeric sort (handles '1st May' < '2nd May', '0000-0100' < '0100-0200')."""
    return [int(c) if c.isdigit() else c.lower()
            for c in re.split(r"(\d+)", str(text))]


def is_session_folder(name: str) -> bool:
    """Return True if this folder name looks like a time-session (e.g. 0000-0100)."""
    return bool(SESSION_RE.match(name))


# ─────────────────────────────────────────────────────────────────────────────
# Video record
# ─────────────────────────────────────────────────────────────────────────────

class DatasetVideo:
    """Represents a single traffic video with its metadata and processing state."""

    def __init__(self, full_path: Path, highway_root: Path, video_root: Path):
        self.full_path    = full_path.resolve()
        self.highway_root = highway_root.resolve()
        self.video_root   = video_root.resolve()

        # Highway name = the folder directly under video_root
        self.highway = highway_root.name

        # Relative to highway_root  →  parts give us the hierarchy
        try:
            rel = self.full_path.relative_to(self.highway_root)
        except ValueError:
            rel = Path(self.full_path.name)

        parts = rel.parts  # e.g. ("LHS", "1st May", "0000-0100", "file.TS")

        # Walk from the end to find the session folder (HHMM-HHMM)
        # Everything above it becomes section + date (however many levels there are)
        session_idx = None
        for i, p in enumerate(parts[:-1]):         # exclude the filename
            if is_session_folder(p):
                session_idx = i
                break

        if session_idx is not None:
            above = parts[:session_idx]             # e.g. ("LHS", "1st May")
            self.section = above[0]  if len(above) >= 1 else "Root"
            self.date    = above[1]  if len(above) >= 2 else "General"
            # Any extra intermediate levels are joined into date for uniqueness
            if len(above) > 2:
                self.date = "_".join(above[1:])
            self.session  = parts[session_idx]
        else:
            # Fallback: no recognisable session folder — use depth-based assignment
            self.section  = parts[0] if len(parts) > 1 else "Root"
            self.date     = parts[1] if len(parts) > 2 else "General"
            self.session  = parts[2] if len(parts) > 3 else "Unknown"

        self.filename     = self.full_path.name
        self.rel_path_str = str(rel).replace("\\", "/")

        # Unique key: highway + rel_path so we can mix multiple highways
        self.global_key = f"{self.highway}/{self.rel_path_str}"

        # Processing state
        self.status            = "Pending"
        self.vehicles_detected = 0
        self.mean_ctr_off      = 0.0
        self.std_ctr_off       = 0.0
        self.error_reason      = ""
        self.timestamp         = ""
        self.output_vid        = ""
        self.output_csv        = ""
        self.output_xlsx       = ""

    # ── serialisation ────────────────────────────────────────────────────────

    def to_dict(self) -> dict:
        return {
            "global_key"       : self.global_key,
            "rel_path"         : self.rel_path_str,
            "full_path"        : str(self.full_path),
            "highway"          : self.highway,
            "section"          : self.section,
            "date"             : self.date,
            "session"          : self.session,
            "filename"         : self.filename,
            "status"           : self.status,
            "vehicles_detected": self.vehicles_detected,
            "mean_ctr_off"     : round(self.mean_ctr_off, 4),
            "std_ctr_off"      : round(self.std_ctr_off, 4),
            "error_reason"     : self.error_reason,
            "timestamp"        : self.timestamp,
            "output_vid"       : self.output_vid,
            "output_csv"       : self.output_csv,
            "output_xlsx"      : self.output_xlsx,
        }

    def from_dict(self, d: dict):
        self.status             = d.get("status",             "Pending")
        self.vehicles_detected  = d.get("vehicles_detected",  0)
        self.mean_ctr_off       = d.get("mean_ctr_off",       0.0)
        self.std_ctr_off        = d.get("std_ctr_off",        0.0)
        self.error_reason       = d.get("error_reason",       "")
        self.timestamp          = d.get("timestamp",          "")
        self.output_vid         = d.get("output_vid",         "")
        self.output_csv         = d.get("output_csv",         "")
        self.output_xlsx        = d.get("output_xlsx",        "")


# ─────────────────────────────────────────────────────────────────────────────
# Discovery
# ─────────────────────────────────────────────────────────────────────────────

def discover_videos_in_highway(highway_root: Path,
                               video_root: Path) -> List[DatasetVideo]:
    """
    Recursively walk a single highway folder.
    Any folder whose name matches HHMM-HHMM is treated as a session folder;
    all video files inside it are collected.
    Non-session folders are traversed recursively without limit.
    """
    results: List[DatasetVideo] = []

    def _walk(directory: Path):
        try:
            entries = sorted(directory.iterdir(), key=lambda p: natural_key(p.name))
        except PermissionError:
            return
        for entry in entries:
            if entry.is_dir():
                if is_session_folder(entry.name):
                    # This IS a session folder — collect videos directly inside
                    for vf in sorted(entry.iterdir(),
                                     key=lambda p: natural_key(p.name)):
                        if vf.is_file() and vf.suffix.lower() in VIDEO_EXTS:
                            results.append(
                                DatasetVideo(vf, highway_root, video_root))
                else:
                    # Non-session folder → recurse
                    _walk(entry)
            # (files directly in non-session dirs are ignored)

    _walk(highway_root)
    return results


def discover_all(video_root: Path) -> Dict[str, List[DatasetVideo]]:
    """
    Find all highway folders directly under video_root and discover their videos.
    Returns dict: { highway_name: [DatasetVideo, …] }
    """
    video_root = video_root.resolve()
    if not video_root.exists():
        sys.exit(f"[ERROR] video-root does not exist: {video_root}")

    highways = sorted(
        [p for p in video_root.iterdir() if p.is_dir()],
        key=lambda p: natural_key(p.name),
    )
    if not highways:
        sys.exit(f"[ERROR] No subdirectories found in {video_root}")

    result: Dict[str, List[DatasetVideo]] = {}
    for hw in highways:
        vids = discover_videos_in_highway(hw, video_root)
        print(f"[DISCOVERY] {hw.name}: {len(vids)} video(s) found")
        result[hw.name] = vids

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Checkpoint  (one JSON per highway, plus unified Excel)
# ─────────────────────────────────────────────────────────────────────────────

class CheckpointManager:

    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.master_excel_path = self.output_dir / "master_dataset_report.xlsx"

    def _ckpt_path(self, highway: str) -> Path:
        safe = re.sub(r'[\\/:*?"<>|]', "_", highway)
        return self.output_dir / f"checkpoint_{safe}.json"

    # ── load ─────────────────────────────────────────────────────────────────

    def load(self, highway: str,
             discovered: List[DatasetVideo]) -> List[DatasetVideo]:
        """Merge discovered videos with saved checkpoint for one highway."""
        ckpt_path = self._ckpt_path(highway)
        ckpt_data: dict = {}

        if ckpt_path.exists():
            try:
                with open(ckpt_path, "r", encoding="utf-8") as f:
                    ckpt_data = json.load(f).get("videos", {})
                print(f"[CHECKPOINT] {highway}: loaded {len(ckpt_data)} saved entries")
            except Exception as e:
                print(f"[WARN] Could not load checkpoint for {highway}: {e}")

        # Also try to restore from existing master Excel (handles case where
        # only the Excel survived but the JSON was lost/corrupted)
        if self.master_excel_path.exists() and XLSX_OK:
            try:
                wb = openpyxl.load_workbook(
                    self.master_excel_path, read_only=True, data_only=True)
                if "Processing Manifest" in wb.sheetnames:
                    ws = wb["Processing Manifest"]
                    for row in ws.iter_rows(min_row=3, values_only=True):
                        # Columns: #, Highway, Section, Date, Session, Filename,
                        #          RelPath, GlobalKey, Status, VehiclesDetected,
                        #          MeanCtr, StdCtr, Error, Timestamp, FullPath,
                        #          AnnotatedPath
                        if not row or len(row) < 9:
                            continue
                        gkey   = row[7]   # global_key column
                        status = row[8]
                        if not gkey or not gkey.startswith(highway + "/"):
                            continue
                        if status not in ("Completed", "Failed"):
                            continue
                        rel_key = "/".join(str(gkey).split("/")[1:])
                        if rel_key in ckpt_data and ckpt_data[rel_key].get("status") in ("Completed", "Failed"):
                            continue
                        ckpt_data[rel_key] = {
                            "status"            : status,
                            "vehicles_detected" : row[9] or 0,
                            "mean_ctr_off"      : row[10] or 0.0,
                            "std_ctr_off"       : row[11] or 0.0,
                            "error_reason"      : row[12] or "",
                            "timestamp"         : row[13] or "",
                            "output_vid"        : row[15] or "",
                        }
                wb.close()
            except Exception as e:
                print(f"[WARN] Excel restore for {highway}: {e}")

        merged: List[DatasetVideo] = []
        for vid in discovered:
            if vid.rel_path_str in ckpt_data:
                vid.from_dict(ckpt_data[vid.rel_path_str])
                if vid.status == "In Progress":
                    vid.status = "Pending"
                    vid.error_reason = "Interrupted — retrying"
            merged.append(vid)
        return merged

    # ── save (JSON) ───────────────────────────────────────────────────────────

    def save_checkpoint(self, highway: str, videos: List[DatasetVideo]):
        data = {
            "highway"      : highway,
            "last_updated" : datetime.datetime.now().isoformat(),
            "videos"       : {v.rel_path_str: v.to_dict() for v in videos},
        }
        ckpt_path = self._ckpt_path(highway)
        with open(ckpt_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    # ── save (master Excel) ───────────────────────────────────────────────────

    def save_master_excel(self, all_videos: List[DatasetVideo]):
        """Write / overwrite the unified master Excel with ALL videos."""
        if not XLSX_OK:
            return

        # Preserve raw detection rows already in the file
        existing_raw_rows: list = []
        existing_raw_keys: set  = set()
        if self.master_excel_path.exists():
            try:
                old_wb = openpyxl.load_workbook(
                    self.master_excel_path, read_only=True, data_only=True)
                if "All Vehicles Raw Data" in old_wb.sheetnames:
                    old_ws = old_wb["All Vehicles Raw Data"]
                    rows = list(old_ws.iter_rows(values_only=True))
                    if rows:
                        existing_raw_rows = rows[1:]        # skip header
                        # Key = (highway, video_filename)  column indices 1, 4
                        existing_raw_keys = {
                            (row[1], row[5])
                            for row in existing_raw_rows
                            if row and len(row) > 5 and row[1] and row[5]
                        }
                old_wb.close()
            except Exception as e:
                print(f"[WARN] Could not read old raw data: {e}")

        # Styles
        thin       = Side(style="thin", color="CCCCCC")
        bdr        = Border(left=thin, right=thin, top=thin, bottom=thin)
        H_FILL     = PatternFill("solid", start_color="0D1B40")
        C_FILL     = PatternFill("solid", start_color="E2F0D9")
        F_FILL     = PatternFill("solid", start_color="FCE4D6")
        P_FILL     = PatternFill("solid", start_color="FFF2CC")
        N_FILL     = PatternFill("solid", start_color="F2F2F2")
        CENTER_ALN = Alignment(horizontal="center")
        LEFT_ALN   = Alignment(horizontal="left")

        wb = openpyxl.Workbook()

        # ── Sheet 1: Processing Manifest ─────────────────────────────────────
        ws = wb.active
        ws.title = "Processing Manifest"

        manifest_headers = [
            "#", "Highway", "Section", "Date", "Time/Session",
            "Video Filename", "Relative Path", "Global Key",
            "Status", "Vehicles Detected",
            "Mean Center Off.(m)", "Std Center Off.(m)",
            "Error / Failure Reason", "Timestamp",
            "Full Path", "Annotated Video Path",
        ]
        col_widths = [5, 28, 12, 12, 15, 25, 35, 45, 14, 16, 18, 18, 30, 20, 45, 45]

        # Title row
        ws.merge_cells(f"A1:{openpyxl.utils.get_column_letter(len(manifest_headers))}1")
        tc = ws["A1"]
        tc.value = "Automated Video-Processing Master Manifest — BTP IIT Kharagpur"
        tc.font      = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
        tc.fill      = H_FILL
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        # Header row
        ws.append(manifest_headers)
        ws.row_dimensions[2].height = 22
        for c_i, (h, cw) in enumerate(zip(manifest_headers, col_widths), 1):
            cell = ws.cell(row=2, column=c_i)
            cell.font      = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
            cell.fill      = H_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = bdr
            ws.column_dimensions[openpyxl.utils.get_column_letter(c_i)].width = cw

        # Data rows
        for i, v in enumerate(all_videos, 1):
            row_data = [
                i, v.highway, v.section, v.date, v.session,
                v.filename, v.rel_path_str, v.global_key,
                v.status, v.vehicles_detected,
                v.mean_ctr_off if v.status == "Completed" else "",
                v.std_ctr_off  if v.status == "Completed" else "",
                v.error_reason, v.timestamp,
                str(v.full_path), v.output_vid,
            ]
            ws.append(row_data)
            fill = (C_FILL if v.status == "Completed" else
                    F_FILL if v.status == "Failed"    else
                    P_FILL if v.status == "In Progress" else N_FILL)
            for c_i in range(1, len(manifest_headers) + 1):
                cell            = ws.cell(row=2 + i, column=c_i)
                cell.font       = Font(name="Calibri", size=10)
                cell.fill       = fill
                cell.border     = bdr
                cell.alignment  = (LEFT_ALN if c_i in (6, 7, 8, 13, 15, 16)
                                   else CENTER_ALN)

        ws.freeze_panes = "A3"
        ws.auto_filter.ref = (
            f"A2:{openpyxl.utils.get_column_letter(len(manifest_headers))}2"
        )

        # ── Sheet 2: All Vehicles Raw Data ────────────────────────────────────
        ws2 = wb.create_sheet("All Vehicles Raw Data")
        raw_headers = [
            "#", "Highway", "Section", "Date", "Time/Session", "Video",
            "Frame", "Timestamp(s)", "Vehicle ID", "Vehicle Type",
            "Left Wheel Off(m)", "Right Wheel Off(m)", "Center Off(m)",
            "LW BEV-X(m)", "RW BEV-X(m)", "Confidence",
        ]
        ws2.append(raw_headers)
        for c_i, h in enumerate(raw_headers, 1):
            cell            = ws2.cell(row=1, column=c_i)
            cell.font       = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
            cell.fill       = H_FILL
            cell.alignment  = Alignment(horizontal="center")
            cell.border     = bdr
            ws2.column_dimensions[
                openpyxl.utils.get_column_letter(c_i)].width = 15

        raw_row_idx = 2

        # Write back previously saved raw rows
        for row_vals in existing_raw_rows:
            ws2.append(list(row_vals))
            for c_i in range(1, len(raw_headers) + 1):
                cell            = ws2.cell(row=raw_row_idx, column=c_i)
                cell.font       = Font(name="Calibri", size=9)
                cell.border     = bdr
                cell.alignment  = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.number_format = "0.0000"
            raw_row_idx += 1

        # Append new rows from this run's CSVs
        for v in all_videos:
            if v.status != "Completed":
                continue
            if not v.output_csv or not Path(v.output_csv).exists():
                continue
            if (v.highway, v.filename) in existing_raw_keys:
                continue   # already in the file
            try:
                with open(v.output_csv, "r", encoding="utf-8") as csv_f:
                    reader = csv.DictReader(csv_f)
                    for r in reader:
                        row_vals = [
                            raw_row_idx - 1,
                            v.highway, v.section, v.date, v.session, v.filename,
                            r.get("Frame",                ""),
                            r.get("Timestamp(s)",         ""),
                            r.get("VehicleID",            ""),
                            r.get("VehicleType",          ""),
                            _float(r.get("LeftWheelOffset(m)",   0)),
                            _float(r.get("RightWheelOffset(m)",  0)),
                            _float(r.get("CenterOffset(m)",      0)),
                            _float(r.get("LeftWheelX_BEV(m)",    0)),
                            _float(r.get("RightWheelX_BEV(m)",   0)),
                            _float(r.get("Confidence",           0)),
                        ]
                        ws2.append(row_vals)
                        for c_i in range(1, len(raw_headers) + 1):
                            cell           = ws2.cell(row=raw_row_idx, column=c_i)
                            cell.font      = Font(name="Calibri", size=9)
                            cell.border    = bdr
                            cell.alignment = Alignment(horizontal="center")
                            if isinstance(cell.value, float):
                                cell.number_format = "0.0000"
                        raw_row_idx += 1
            except Exception as e:
                print(f"[WARN] CSV read failed for {v.filename}: {e}")

        ws2.freeze_panes = "A2"

        try:
            wb.save(self.master_excel_path)
            print(f"[EXCEL] Saved → {self.master_excel_path}")
        except PermissionError:
            alt = self.master_excel_path.with_stem(
                self.master_excel_path.stem + f"_bak_{int(time.time())}")
            try:
                wb.save(alt)
                print(f"[EXCEL] Master file locked — saved backup to {alt}")
            except Exception as e2:
                print(f"[WARN] Could not write Excel at all: {e2}")
        except Exception as e:
            print(f"[WARN] Excel write error: {e}")


def _float(val, default=0.0) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


# ─────────────────────────────────────────────────────────────────────────────
# Calibration helpers
# ─────────────────────────────────────────────────────────────────────────────

def calib_matches(path: Path, rect_width: float, rect_length: float) -> bool:
    try:
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return (abs(float(d.get("rw", -1)) - rect_width)  < 1e-6 and
                abs(float(d.get("rl", -1)) - rect_length) < 1e-6)
    except Exception:
        return False


def calib_path_for(vid: DatasetVideo, out_root: Path) -> Tuple[Path, bool]:
    """
    Return (calib_json_path, is_headless).
    Calibrations are keyed by (highway, section, date) so each distinct
    camera view gets its own homography.
    """
    safe_hw  = re.sub(r'[\\/:*?"<>|]', "_", vid.highway)
    safe_sec = re.sub(r'[\\/:*?"<>|]', "_", vid.section)
    safe_dt  = re.sub(r'[\\/:*?"<>|]', "_", vid.date)
    fname    = f"calib_{safe_hw}_{safe_sec}_{safe_dt}.json"
    path     = out_root / fname
    return path, path.exists()


# ─────────────────────────────────────────────────────────────────────────────
# Single-video processing
# ─────────────────────────────────────────────────────────────────────────────

def process_video(vid: DatasetVideo, out_root: Path,
                  model: str, conf: float, iou: float, imgsz: int,
                  skip_frames: int, no_riders: bool,
                  road_width: float, lane_width: float,
                  rect_width: float, rect_length: float,
                  force_calib: bool, save_video: bool) -> Tuple[bool, str]:

    calib_json, already_calibrated = calib_path_for(vid, out_root)
    is_headless = already_calibrated and not force_calib

    if not is_headless:
        print("\n" + "=" * 70, flush=True)
        print(f"  [INTERACTIVE CALIBRATION]", flush=True)
        print(f"  Highway : {vid.highway}", flush=True)
        print(f"  Section : {vid.section}  |  Date: {vid.date}", flush=True)
        print(f"  Video   : {vid.filename}", flush=True)
        print(f"  Calib will be saved to: {calib_json.name}", flush=True)
        print(f"  (Reused headlessly for all {vid.section}/{vid.date} videos of this highway)", flush=True)
        print("=" * 70 + "\n", flush=True)

    # Output directory: outputs_all/<highway>/<section>/<date>/<session>/<stem>/
    out_dir = (out_root / vid.highway / vid.section / vid.date /
               vid.session / vid.full_path.stem)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_vid_path = out_dir / f"{vid.full_path.stem}_annotated.mp4"

    vid.output_vid  = str(out_vid_path) if save_video else "Not Saved"
    vid.output_csv  = str(out_dir / f"{vid.full_path.stem}_offsets.csv")
    vid.output_xlsx = str(out_dir / f"{vid.full_path.stem}_offsets.xlsx")

    cmd = [
        sys.executable, "pipeline.py",
        "--source",      str(vid.full_path),
        "--model",       model,
        "--output",      str(out_vid_path),
        "--conf",        str(conf),
        "--iou",         str(iou),
        "--imgsz",       str(imgsz),
        "--skip-frames", str(skip_frames),
        "--road-width",  str(road_width),
        "--lane-width",  str(lane_width),
        "--rect-width",  str(rect_width),
        "--rect-length", str(rect_length),
        "--calib-file",  str(calib_json),
    ]
    if is_headless:
        cmd.append("--headless")
    if not save_video:
        cmd.append("--no-video")
    if force_calib or not already_calibrated:
        cmd.append("--force-calib")
    if no_riders:
        cmd.append("--no-riders")

    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"\n[{ts}] RUNNING: {vid.highway}/{vid.rel_path_str}", flush=True)

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
    )
    output_lines: List[str] = []
    if proc.stdout:
        for line in iter(proc.stdout.readline, ""):
            s = line.strip()
            if any(s.startswith(tag) for tag in
                   ("[PROGRESS]", "[INFO]", "[CALIB]", "[DONE]",
                    "[XLSX]", "[CSV]")):
                print(f"   --> {s}", flush=True)
            output_lines.append(s)
    proc.wait()

    if proc.returncode != 0:
        err = ("\n".join(output_lines[-20:]) if output_lines
               else f"Exit code {proc.returncode}")
        return False, err

    # Parse output CSV for summary stats
    csv_path = Path(vid.output_csv)
    if csv_path.exists():
        try:
            with open(csv_path, "r", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
            vid.vehicles_detected = len(rows)
            if rows:
                coffs = [_float(r.get("CenterOffset(m)", 0)) for r in rows]
                import numpy as np
                vid.mean_ctr_off = float(np.mean(coffs))
                vid.std_ctr_off  = float(np.std(coffs))
        except Exception as e:
            print(f"[WARN] CSV parse error ({vid.filename}): {e}")

    return True, ""


# ─────────────────────────────────────────────────────────────────────────────
# Progress summary
# ─────────────────────────────────────────────────────────────────────────────

def print_summary(all_videos: List[DatasetVideo], current: int = -1):
    total     = len(all_videos)
    completed = sum(1 for v in all_videos if v.status == "Completed")
    failed    = sum(1 for v in all_videos if v.status == "Failed")
    pending   = sum(1 for v in all_videos if v.status == "Pending")
    in_prog   = sum(1 for v in all_videos if v.status == "In Progress")
    pct       = completed / total * 100 if total else 0.0

    # Per-highway breakdown
    highways = {}
    for v in all_videos:
        hw = v.highway
        if hw not in highways:
            highways[hw] = {"total": 0, "done": 0, "fail": 0}
        highways[hw]["total"] += 1
        if v.status == "Completed": highways[hw]["done"] += 1
        if v.status == "Failed":    highways[hw]["fail"] += 1

    print("\n" + "=" * 70)
    print("  MULTI-ROOT PIPELINE SUMMARY")
    print("=" * 70)
    print(f"  Grand Total : {total}  |  Done: {completed} ({pct:.1f}%)"
          f"  |  Failed: {failed}  |  Pending: {pending}")
    if in_prog:
        print(f"  In Progress : {in_prog}")
    print()
    for hw, s in sorted(highways.items()):
        p = s["done"] / s["total"] * 100 if s["total"] else 0
        print(f"    {hw[:45]:<45}  {s['done']:4}/{s['total']:4} ({p:5.1f}%)"
              f"  fail={s['fail']}")
    if 0 <= current < total:
        v = all_videos[current]
        print(f"\n  Current: [{current+1}/{total}] {v.highway}/{v.rel_path_str}")
    elif completed == total and total > 0:
        print("\n  ALL VIDEOS COMPLETED SUCCESSFULLY!")
    print("=" * 70 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# Post-run analysis
# ─────────────────────────────────────────────────────────────────────────────

def run_post_analysis(output_dir: Path):
    csv_files = list(output_dir.glob("**/*_offsets.csv"))
    if not csv_files:
        print("[ANALYSIS] No CSV files found — skipping analysis.")
        return
    analysis_dir = output_dir / "combined_analysis"
    cmd = ([sys.executable, "analyze.py", "--csv"]
           + [str(c) for c in csv_files]
           + ["--output-dir", str(analysis_dir)])
    print(f"\n[ANALYSIS] Running combined analysis on {len(csv_files)} CSV files...")
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    print(res.stdout)
    print(f"[ANALYSIS] Plots saved to: {analysis_dir.resolve()}")


# ─────────────────────────────────────────────────────────────────────────────
# main()
# ─────────────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Multi-Root Automated Video Processor — BTP IIT Kharagpur",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    ap.add_argument("--video-root", default=r"f:\BTP\video",
                    help="Top-level folder containing all highway sub-folders")
    ap.add_argument("--output-dir", default=r"f:\BTP\Autmn Sem\outputs_all",
                    help="Root output directory (all results go here)")
    ap.add_argument("--model",        default="yolo11n",
                    help="YOLO model (yolo11n for CPU speed, yolo11x for GPU accuracy)")
    ap.add_argument("--conf",         type=float, default=0.20)
    ap.add_argument("--iou",          type=float, default=0.50)
    ap.add_argument("--imgsz",        type=int,   default=640)
    ap.add_argument("--skip-frames",  type=int,   default=2,
                    help="Process every Nth frame (2 = 2x speedup)")
    ap.add_argument("--no-riders",    action="store_true")
    ap.add_argument("--road-width",   type=float, default=7.0)
    ap.add_argument("--lane-width",   type=float, default=3.5)
    ap.add_argument("--rect-width",   type=float, default=7.0)
    ap.add_argument("--rect-length",  type=float, default=8.0)
    ap.add_argument("--save-video",   action="store_true",
                    help="Save annotated MP4 output (off by default to save disk)")
    ap.add_argument("--retry-failed", action="store_true",
                    help="Retry all previously failed videos")
    ap.add_argument("--force-calib",  action="store_true",
                    help="Force interactive recalibration for ALL (highway/section/date) groups")
    ap.add_argument("--only-highway", default="",
                    help="Process only this highway folder name (substring match). "
                         "E.g. --only-highway NH-18")
    ap.add_argument("--skip-analysis", action="store_true",
                    help="Skip post-processing analysis step")
    args = ap.parse_args()

    video_root = Path(args.video_root)
    out_root   = Path(args.output_dir)
    out_root.mkdir(parents=True, exist_ok=True)

    # ── 1. Discover all highways & their videos ───────────────────────────────
    by_highway = discover_all(video_root)

    if args.only_highway:
        by_highway = {k: v for k, v in by_highway.items()
                      if args.only_highway.lower() in k.lower()}
        if not by_highway:
            sys.exit(f"[ERROR] --only-highway '{args.only_highway}' matched nothing.")
        print(f"[FILTER] Processing only: {list(by_highway.keys())}")

    # ── 2. Load checkpoints ───────────────────────────────────────────────────
    chk_mgr = CheckpointManager(out_root)
    highway_videos: Dict[str, List[DatasetVideo]] = {}
    for hw, discovered in by_highway.items():
        highway_videos[hw] = chk_mgr.load(hw, discovered)

    # Flat list (for summary / master Excel), ordered highway → section → date → session → file
    all_videos: List[DatasetVideo] = [
        v for hw in sorted(highway_videos) for v in highway_videos[hw]
    ]

    print_summary(all_videos)

    # ── 3. Ctrl-C handler ─────────────────────────────────────────────────────
    def sig_handler(sig, frame):
        print("\n[INTERRUPT] Saving state before exit...")
        for hw, vids in highway_videos.items():
            chk_mgr.save_checkpoint(hw, vids)
        chk_mgr.save_master_excel(all_videos)
        print_summary(all_videos)
        sys.exit(0)

    signal.signal(signal.SIGINT, sig_handler)

    # ── 4. Processing loop ────────────────────────────────────────────────────
    for global_idx, vid in enumerate(all_videos):
        if vid.status == "Completed":
            continue
        if vid.status == "Failed" and not args.retry_failed:
            continue

        # Mark in-progress immediately
        vid.status    = "In Progress"
        vid.timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        chk_mgr.save_checkpoint(vid.highway, highway_videos[vid.highway])
        chk_mgr.save_master_excel(all_videos)

        success, err = process_video(
            vid        = vid,
            out_root   = out_root,
            model      = args.model,
            conf       = args.conf,
            iou        = args.iou,
            imgsz      = args.imgsz,
            skip_frames= args.skip_frames,
            no_riders  = args.no_riders,
            road_width = args.road_width,
            lane_width = args.lane_width,
            rect_width = args.rect_width,
            rect_length= args.rect_length,
            force_calib= args.force_calib,
            save_video = args.save_video,
        )

        vid.timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if success:
            vid.status       = "Completed"
            vid.error_reason = ""
            print(f"[OK] {vid.highway}/{vid.rel_path_str}"
                  f" -> {vid.vehicles_detected} vehicles"
                  f" | mean={vid.mean_ctr_off:.3f}m")
        else:
            vid.status       = "Failed"
            vid.error_reason = err
            print(f"[FAIL] {vid.highway}/{vid.rel_path_str} -> {err[:120]}")

        # Persist after every video
        chk_mgr.save_checkpoint(vid.highway, highway_videos[vid.highway])
        chk_mgr.save_master_excel(all_videos)
        print_summary(all_videos, current=global_idx)

    # ── 5. Finish ─────────────────────────────────────────────────────────────
    print("\n[DONE] All videos processed!")
    chk_mgr.save_master_excel(all_videos)

    if not args.skip_analysis:
        run_post_analysis(out_root)

    print_summary(all_videos)


if __name__ == "__main__":
    main()
