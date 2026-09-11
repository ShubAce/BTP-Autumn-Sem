# Wheel Wander & Lateral Offset Measurement — v4.0

**BTP | IIT Kharagpur | Computer Vision for Traffic Engineering**

---

## Overview

An end-to-end computer vision pipeline to automatically measure:

- **Left-wheel offset** from the right road boundary
- **Right-wheel offset** from the right road boundary
- **Vehicle center offset** from the right road boundary

for trucks, buses, cars, motorcycles and other vehicles from roadside camera footage.

---

## Files

| File               | Purpose                                      |
| ------------------ | -------------------------------------------- |
| `pipeline.py`      | Main real-time processing pipeline           |
| `analyze.py`       | Post-processing statistical analysis         |
| `batch.py`         | Process all videos in a folder automatically |
| `requirements.txt` | Python dependencies                          |

---

## Setup

```powershell
pip install -r requirements.txt
```

YOLOv11 weights are auto-downloaded on first run from Ultralytics.

---

## Running a Single Video

```powershell
python pipeline.py `
  --source "f:\BTP\videos\094147-00015-M.TS" `
  --road-width 7.0 `
  --lane-width 3.5 `
  --rect-width 7.0 `
  --rect-length 8.0 `
  --model yolo11x `
  --output outputs\094147_annotated.mp4 `
  --calib-file calib.json
```

### Interactive Setup (first run)

When the video opens, you will see three setup phases:

**Phase 1 — Calibration (4-point homography)**

- Click **4 corners** of a known road rectangle in order:
    - P1 = Far Left, P2 = Far Right, P3 = Near Left, P4 = Near Right
- Use lane markings, kerb edges, or road paint as reference points
- The calibration is saved to `calib.json` and **reused automatically** for all future videos from the same camera

**Phase 2 — Right Boundary Line**

- Click **2 points** along the rightmost road edge
- This becomes the reference (offset = 0)

**Phase 3 — Measurement Gate**

- Click **2 points** to draw a line across the road
- Offset measurements are recorded when vehicles cross this gate

### Keyboard Controls

| Key         | Action                         |
| ----------- | ------------------------------ |
| `SPACE`     | Pause / Resume                 |
| `E`         | Export Excel + CSV immediately |
| `S`         | Save current frame as PNG      |
| `R`         | Redo calibration               |
| `Q` / `ESC` | Quit (auto-exports on exit)    |

---

## Batch Processing All Videos

```powershell
python batch.py `
  --video-dir "f:\BTP\videos" `
  --output-dir outputs `
  --calib-file calib.json `
  --model yolo11x
```

- First video opens interactively for calibration → saved to `calib.json`
- All subsequent videos use saved calibration automatically
- After all videos, runs combined statistical analysis

The automated hierarchy runner uses a separate 3-lane geometry for `RHS / 2nd May`
(`10.5 m` road width with `3.5 m` lanes). Its first video is calibrated separately;
the other section/date groups continue to use the existing 2-lane defaults.

---

## Statistical Analysis Only

```powershell
python analyze.py `
  --csv outputs\*\*_offsets.csv `
  --output-dir analysis_output
```

Produces:

- `analysis_output/statistics.csv`
- `analysis_output/wheel_wander_report.xlsx` (full stats + per-vehicle-type)
- `analysis_output/figures/` (6 publication-quality PNG charts)

---

## Output Excel Structure

| Column          | Description                               |
| --------------- | ----------------------------------------- |
| Video           | Source video filename                     |
| Frame           | Frame number                              |
| Time(s)         | Timestamp in seconds                      |
| Vehicle ID      | Unique track ID                           |
| Type            | truck / bus / car / motorcycle / bicycle  |
| L-Wheel Off.(m) | Left wheel offset from right boundary     |
| R-Wheel Off.(m) | Right wheel offset from right boundary    |
| Center Off.(m)  | Vehicle center offset from right boundary |
| LW BEV-X(m)     | Left wheel X in bird's-eye-view metres    |
| RW BEV-X(m)     | Right wheel X in bird's-eye-view metres   |
| Confidence      | Detection confidence score                |

---

## Calibration Tips

- **Lane width**: Standard NH lane = 3.5 m (use as `--lane-width`)
- **Road width**: Measure from video or use 7.0 m (2-lane) / 10.5 m (3-lane)
- **rect-width / rect-length**: The real-world size of your clicked quadrilateral
- Once `calib.json` is saved, you never need to click again for that camera

---

## Model Options (--model)

| Model     | Size   | Speed    | Accuracy               |
| --------- | ------ | -------- | ---------------------- |
| `yolo11n` | Nano   | Fastest  | Lower                  |
| `yolo11s` | Small  | Fast     | Good                   |
| `yolo11m` | Medium | Balanced | Better                 |
| `yolo11l` | Large  | Slower   | High                   |
| `yolo11x` | XLarge | Slowest  | **Best** ← recommended |

---

## Generated Figures

1. `01_distribution_overview.png` — KDE distribution by vehicle type
2. `02_truck_wheel_wander.png` — Histogram + normal fit for trucks
3. `03_boxplot_by_type.png` — Box-plots per vehicle class
4. `04_lane_utilisation_heatmap.png` — 2D heat-map of wheel positions
5. `05_center_offset_timeseries.png` — Offset evolution over time
6. `06_vehicle_distribution.png` — Pie chart of vehicle mix

---

_IIT Kharagpur BTP | Autumn Semester_
