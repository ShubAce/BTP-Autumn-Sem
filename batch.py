"""
=============================================================================
  Batch Video Processor  —  BTP IIT Kharagpur
=============================================================================
  Runs pipeline.py on every video in the input directory.
  All output videos + Excel files go into output/<video_stem>/

  Usage:
    python batch.py --video-dir "f:/BTP/videos" --output-dir outputs
                    --calib-file calib.json [--model yolo11x] [--conf 0.30]

  The first video will trigger interactive calibration → saved to calib.json
  All subsequent videos reuse that calibration automatically.

  After processing all videos, runs statistical analysis across ALL CSVs.
=============================================================================
"""

from __future__ import annotations
import argparse, subprocess, sys
from pathlib import Path

VIDEO_EXTS = {".ts", ".mp4", ".avi", ".mkv", ".mov", ".mpeg", ".mpg"}

def find_videos(video_dir: str) -> list[Path]:
    vdir = Path(video_dir)
    vids = sorted([p for p in vdir.iterdir()
                   if p.suffix.lower() in VIDEO_EXTS])
    if not vids:
        sys.exit(f"[ERROR] No video files found in {video_dir}")
    print(f"[BATCH] Found {len(vids)} video(s):")
    for v in vids:
        print(f"         • {v.name}")
    return vids

def run_pipeline(video: Path, out_root: Path, calib_file: str,
                 model: str, conf: float,
                 road_width: float, lane_width: float,
                 rect_width: float, rect_length: float,
                 force_calib: bool = False) -> Path:
    stem    = video.stem
    out_dir = out_root / stem
    out_dir.mkdir(parents=True, exist_ok=True)
    out_vid = str(out_dir / f"{stem}_annotated.mp4")

    cmd = [
        sys.executable, "pipeline.py",
        "--source",       str(video),
        "--model",        model,
        "--output",       out_vid,
        "--conf",         str(conf),
        "--road-width",   str(road_width),
        "--lane-width",   str(lane_width),
        "--rect-width",   str(rect_width),
        "--rect-length",  str(rect_length),
    ]
    if calib_file:
        cmd += ["--calib-file", calib_file]
    if force_calib:
        cmd += ["--force-calib"]
    cmd += ["--fps", "24"]

    print(f"\n{'='*58}")
    print(f"  Processing: {video.name}")
    print(f"  Output dir: {out_dir}")
    print(f"{'='*58}")
    subprocess.run(cmd, check=True)
    return out_dir

def run_analysis(csv_files: list[Path], out_dir: Path):
    if not csv_files:
        print("[BATCH] No CSV files to analyse.")
        return
    analysis_dir = out_dir / "combined_analysis"
    cmd = [
        sys.executable, "analyze.py",
        "--csv"] + [str(c) for c in csv_files] + [
        "--output-dir", str(analysis_dir),
    ]
    print(f"\n[BATCH] Running combined statistical analysis…")
    subprocess.run(cmd, check=True)
    print(f"[BATCH] Analysis saved → {analysis_dir}")

def main():
    ap = argparse.ArgumentParser(
        description="Batch Processor — BTP IIT Kharagpur",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    ap.add_argument("--video-dir",    default=r"f:\BTP\videos",
                    help="Directory containing .TS / .mp4 video files")
    ap.add_argument("--output-dir",   default="outputs",
                    help="Root output directory")
    ap.add_argument("--calib-file",   default="calib.json",
                    help="Calibration JSON (saved from first video, reused for rest)")
    ap.add_argument("--model",        default="yolo11x",
                    help="YOLO model name")
    ap.add_argument("--conf",         type=float, default=0.30)
    ap.add_argument("--road-width",   type=float, default=7.0)
    ap.add_argument("--lane-width",   type=float, default=3.5)
    ap.add_argument("--rect-width",   type=float, default=7.0)
    ap.add_argument("--rect-length",  type=float, default=8.0)
    ap.add_argument("--skip-analysis", action="store_true",
                    help="Skip the post-batch statistical analysis")
    args = ap.parse_args()

    videos  = find_videos(args.video_dir)
    out_root = Path(args.output_dir)

    csv_files: list[Path] = []
    for index, video in enumerate(videos):
        out_dir = run_pipeline(
            video        = video,
            out_root     = out_root,
            calib_file   = args.calib_file,
            model        = args.model,
            conf         = args.conf,
            road_width   = args.road_width,
            lane_width   = args.lane_width,
            rect_width   = args.rect_width,
            rect_length  = args.rect_length,
            force_calib  = (index == 0),
        )
        # collect CSV files from this run
        csv_files.extend(out_dir.glob("*_offsets.csv"))

    if not args.skip_analysis:
        run_analysis(csv_files, out_root)

    print(f"\n[BATCH] All done!  Results in: {out_root.resolve()}")

if __name__ == "__main__":
    main()
