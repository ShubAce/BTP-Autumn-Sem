# How to Run This Project on Another PC

This guide provides step-by-step instructions to setup and run the **Wheel Wander & Lateral Offset Analysis** pipeline on a new computer from scratch.

---

## 📋 1. Prerequisites & System Requirements

- **Operating System**: Windows 10/11 (or Linux / macOS)
- **Python**: Version **3.9, 3.10, or 3.11** installed and added to your system `PATH`.
- **GPU (Optional, but Recommended)**: NVIDIA GPU with CUDA drivers for fast video processing. (CPU mode will work automatically if no GPU is detected).

---

## 📁 2. Step-by-Step Setup Instructions

### Step 1: Copy Code & Videos to the New PC
1. Copy the project folder (containing `pipeline.py`, `batch.py`, `analyze.py`, `requirements.txt`, etc.) to the new PC.
2. Put your input video files (e.g. `.mp4`, `.avi`, `.TS`) into a folder on your PC (e.g. `C:\Traffic_Videos` or inside a `videos/` subfolder).

---

### Step 2: Open Terminal & Create a Virtual Environment

1. Open **PowerShell** or **Command Prompt** and navigate to the project directory:
   ```powershell
   cd "C:\path\to\your\project_folder"
   ```

2. Create a Python Virtual Environment:
   ```powershell
   python -m venv venv
   ```

3. Activate the Virtual Environment:
   - **Windows PowerShell**:
     ```powershell
     .\venv\Scripts\Activate.ps1
     ```
     *(If you get a script execution policy error, run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process` first)*.
   - **Windows CMD**:
     ```cmd
     venv\Scripts\activate.bat
     ```
   - **Linux / macOS**:
     ```bash
     source venv/bin/activate
     ```

---

### Step 3: Install Required Dependencies

Upgrade `pip` and install all required Python packages:

```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

---

### Step 4: Model Weights (`yolo11x.pt`)

- **Online Mode (Internet available)**: The model weights (`yolo11x.pt`) will automatically download on the first run.
- **Offline Mode (No Internet)**: Copy the file `yolo11x.pt` directly into the main project root folder alongside `pipeline.py`.

---

## 🚀 3. How to Run the Project

### Option A: Process a Single Video

Run `pipeline.py` with your video file path:

```powershell
python pipeline.py --source "C:\Traffic_Videos\sample_video.mp4" --model yolo11x --output outputs/sample_annotated.mp4 --calib-file calib.json
```

#### 🎯 First-Time Interactive Calibration (8 Clicks)
When the video window opens, follow the on-screen prompt to click **8 reference points**:

1. **Phase 1 (4 Clicks)**: Click the **4 corners** of a known road section in order:
   - **Point 1**: Far Left
   - **Point 2**: Far Right
   - **Point 3**: Near Left
   - **Point 4**: Near Right
2. **Phase 2 (2 Clicks)**: Click **2 points** along the **Right Road Boundary** (where Offset = 0 m).
3. **Phase 3 (2 Clicks)**: Click **2 points** across the road to set the **Measurement Gate Line**.

> 💡 **Note**: Calibration is automatically saved to `calib.json` and reused for subsequent runs.

#### ⌨️ On-Screen Keyboard Controls:
| Key | Action |
|---|---|
| `SPACE` | Pause / Resume video playback |
| `E` | Export CSV & Excel results immediately |
| `S` | Save current frame screenshot |
| `R` | Redo camera calibration |
| `Q` / `ESC` | Quit (saves data automatically on exit) |

---

### Option B: Batch Process All Videos in a Folder

To process a directory full of traffic videos automatically:

```powershell
python batch.py --video-dir "C:\Traffic_Videos" --output-dir outputs --calib-file calib.json --model yolo11x
```

- Each video will be processed using the saved calibration or prompt for calibration setup if needed.
- Results will be saved under the `outputs/` folder.

---

### Option C: Generate Statistical Reports & Graphs

After processing videos and generating CSV files in the `outputs/` folder, run `analyze.py`:

```powershell
python analyze.py --csv outputs/*/*_offsets.csv --output-dir analysis_output
```

#### 📊 Outputs Generated in `analysis_output/`:
- `statistics.csv` — Summary metrics (mean, std, min, max, P5, P95, KS normality test).
- `wheel_wander_report.xlsx` — Formatted Excel report with raw data and stats tabs.
- `figures/` — 6 publication-quality PNG charts (distributions, heatmaps, box plots, time series).

---

## 🛠️ 4. Common Troubleshooting Tips

1. **Path with Spaces**: Always wrap file paths in double quotes, e.g. `"C:\My Videos\traffic.mp4"`.
2. **Import Errors**: Ensure your virtual environment is active (`(venv)` should appear in your terminal prompt).
3. **Faster / Lighter Model**: If GPU memory or CPU speed is limited, switch from `yolo11x` to a smaller model like `--model yolo11m` or `--model yolo11s`.
