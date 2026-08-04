# -*- coding: utf-8 -*-
"""
=============================================================================
  Wheel Wander & Lateral Offset Measurement Pipeline  --  v4.0
  BTP | IIT Kharagpur | Computer Vision for Traffic Engineering
=============================================================================
  Upgrades over prior versions:
    * YOLOv11x (largest COCO model) for detection
    * Dual-wheel localization  (bottom-left / bottom-right of vehicle bbox,
      refined by gradient-weighted centroid inside wheel-radius crop)
    * LEFT + RIGHT + CENTER offsets per vehicle (not just center)
    * Signed offsets  (+ = further from right boundary)
    * Calibration auto-save / auto-load  (JSON) so you only click once
    * CLAHE night-time pre-processor  (adaptive, brightness-triggered)
    * Annotated output video  with offset rulers and wheel markers
    * Rich Excel export  (measurements + per-sheet stats + auto charts)
    * Statistical analysis module  (run standalone after processing)
    * BoT-SORT tracker  (handles camera vibration via CMC)
    * Full argparse CLI
=============================================================================
  Usage:
    python pipeline.py --source "path/to/video.TS" --road-width 7.0 --lane-width 3.5
  Keys during playback:
    SPACE  ->  pause / resume
    S      ->  save current frame as PNG
    E      ->  export Excel + CSV immediately
    R      ->  redo calibration
    Q/ESC  ->  quit + auto-export
=============================================================================
"""
from __future__ import annotations
import argparse, csv, json, sys, time
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import cv2
import numpy as np
# -- optional heavy deps -- graceful fallback -------------------------------
try:
    from ultralytics import YOLO
    YOLO_OK = True
except ImportError:
    YOLO_OK = False
try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                  GradientFill)
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import DataPoint
    XLSX_OK = True
except ImportError:
    XLSX_OK = False
try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False
# -- COCO class map --------------------------------------------------------
VEHICLE_CLASSES: Dict[int, str] = {
    0: "rider",            # person/rider
    1: "bicycle",          # bicycle/cycle
    2: "car",              # car / sedan / hatchback / SUV / autorickshaw / pickup
    3: "motorcycle",       # motorcycle / scooter / two_wheeler / autorickshaw
    5: "bus",              # bus / minibus
    6: "multi_axle_truck", # train in COCO -> multi-axle / container truck in traffic
    7: "truck",            # truck / pickup / LCV
}

# Per-class confidence thresholds to ensure small/fast 2-wheelers (bikes, motorcycles)
# and riders are retained even when low-scoring
CLASS_CONF_THRESHOLDS: Dict[int, float] = {
    0: 0.15,  # rider/person
    1: 0.12,  # bicycle
    3: 0.12,  # motorcycle
    2: 0.20,  # car
    5: 0.20,  # bus
    6: 0.20,  # multi-axle truck
    7: 0.20,  # truck
}

TRUCK_CLASSES = {"bus", "truck", "multi_axle_truck", "pickup"}  # Heavy vehicles for wheel wander focus

def classify_indian_vehicle(cls_id: int, bw: int, bh: int, conf: float) -> str:
    """
    Refines COCO class predictions into standard Indian traffic categories:
    - bicycle
    - motorcycle
    - autorickshaw (3-wheeler)
    - car
    - pickup / van (LCV)
    - bus
    - truck
    - multi_axle_truck
    - rider
    """
    aspect_ratio = bw / float(bh) if bh > 0 else 1.0
    area = bw * bh

    if cls_id == 1:
        return "bicycle"
    elif cls_id == 3:
        # Auto-rickshaws detected as motorcycle (boxy, wider profile)
        if aspect_ratio > 0.85 and area > 15000:
            return "autorickshaw"
        return "motorcycle"
    elif cls_id == 0:
        return "rider"
    elif cls_id == 2:  # COCO car
        # Auto-rickshaw (tall boxy 3-wheeler)
        if 0.65 <= aspect_ratio <= 1.05 and area < 28000:
            return "autorickshaw"
        # Small pickup / LCV (e.g. Tata Ace, Eeco, Bolero Pickup)
        elif 0.85 <= aspect_ratio <= 1.35 and area > 32000:
            return "pickup"
        return "car"
    elif cls_id == 5:  # COCO bus
        return "bus"
    elif cls_id == 6:  # COCO train -> multi axle truck in traffic
        return "multi_axle_truck"
    elif cls_id == 7:  # COCO truck
        if area > 80000 or aspect_ratio > 1.8:
            return "multi_axle_truck"
        elif area < 30000:
            return "pickup"
        return "truck"
    return "vehicle"
# -- palette ---------------------------------------------------------------
C_REF       = (0,   255, 180)   # reference boundary line
C_BOX       = (255, 200,   0)   # vehicle box
C_WHEEL_L   = (0,   200, 255)   # left wheel
C_WHEEL_R   = (255,  80, 200)   # right wheel
C_CENTER    = (255, 255,   0)   # center track
C_LABEL     = (255, 255, 255)
C_CALIB     = (0,   160, 255)
C_GRID      = (60,  180,  60)
C_WIRE      = (0,   255,   0)
C_RULER     = (255, 165,   0)   # measurement ruler
WIN         = "Wheel Wander Analyzer -- v4.0"
# ---------------------------------------------------------------------------
class NightEnhancer:
    """Adaptive CLAHE -- applied only when mean brightness < threshold."""
    def __init__(self, brightness_threshold: int = 80,
                 clip: float = 3.0, tile: Tuple[int,int] = (8, 8)):
        self.thresh   = brightness_threshold
        self.clahe    = cv2.createCLAHE(clipLimit=clip, tileGridSize=tile)
    def enhance(self, frame: np.ndarray) -> np.ndarray:
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if gray.mean() < self.thresh:
            lab = cv2.cvtColor(frame, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            l = self.clahe.apply(l)
            return cv2.cvtColor(cv2.merge([l, a, b]), cv2.COLOR_LAB2BGR)
        return frame
# ---------------------------------------------------------------------------
class Calibrator:
    """
    4-point homography calibration with JSON save/load.
    Click order:
      P1 = far-left    P2 = far-right
      P3 = near-left   P4 = near-right
    (corresponds to corners of the known road rectangle)
    """
    LABELS = ["P1 far-left", "P2 far-right", "P3 near-left", "P4 near-right"]
    INSTRUCTIONS = [
        "SETUP 1/3 -- Click 4 road-plane corners (8 points total):",
        "  1) FAR LEFT   2) FAR RIGHT",
        "  3) NEAR LEFT  4) NEAR RIGHT",
        "  Then click 2 right-boundary points and 2 gate points.",
        "  Use lane markings / kerbs as reference points.",
        "  Tip: enter real-world dimensions via --rect-width / --rect-length",
    ]
    def __init__(self, road_width_m: float, road_length_m: float,
                 canvas_px: float = 1200.0):
        self.rw   = float(road_width_m)
        self.rl   = float(road_length_m)
        self.scale = canvas_px / max(self.rw, self.rl)
        W_px, L_px = self.rw * self.scale, self.rl * self.scale
        self._dst = np.float32([[0, 0], [W_px, 0], [0, L_px], [W_px, L_px]])
        self.img_pts:  List[Tuple[float,float]] = []
        self.H:        Optional[np.ndarray] = None
        self.H_inv:    Optional[np.ndarray] = None
        self.confirmed = False
        self._hover:   Optional[Tuple[int,int]] = None
    # -- interaction ----------------------------------------------------
    def mouse_cb(self, event, x, y, flags, param):
        if self.confirmed: return
        if event == cv2.EVENT_MOUSEMOVE:
            self._hover = (x, y)
        elif event == cv2.EVENT_LBUTTONDOWN:
            if len(self.img_pts) < 4:
                self.img_pts.append((float(x), float(y)))
                if len(self.img_pts) == 4:
                    self._compute()
                    self.confirmed = True
        elif event == cv2.EVENT_RBUTTONDOWN and self.img_pts:
            self.img_pts.pop()
    def _compute(self):
        src = np.float32(self.img_pts)
        self.H     = cv2.getPerspectiveTransform(src, self._dst)
        self.H_inv = cv2.getPerspectiveTransform(self._dst, src)
    def reset(self):
        self.img_pts = []; self.H = None; self.H_inv = None
        self.confirmed = False; self._hover = None
    # -- coordinate transforms ------------------------------------------
    def to_bev(self, pt: Tuple[float, float]) -> np.ndarray:
        """Image pixel -> BEV metres."""
        p = np.array([[[float(pt[0]), float(pt[1])]]], dtype=np.float32)
        return cv2.perspectiveTransform(p, self.H)[0][0] / self.scale
    def to_img(self, bev_m: np.ndarray) -> np.ndarray:
        """BEV metres -> image pixel."""
        px = np.array(bev_m, dtype=np.float32) * self.scale
        p  = np.array([[[px[0], px[1]]]], dtype=np.float32)
        return cv2.perspectiveTransform(p, self.H_inv)[0][0]
    # -- serialisation --------------------------------------------------
    def save(self, path: str):
        d = {"img_pts": self.img_pts, "rw": self.rw, "rl": self.rl}
        with open(path, "w") as f:
            json.dump(d, f, indent=2)
        print(f"[CALIB] Saved -> {path}")
    def load(self, path: str) -> bool:
        try:
            with open(path) as f:
                d = json.load(f)
            self.img_pts = [tuple(p) for p in d["img_pts"]]
            if len(self.img_pts) == 4:
                self._compute(); self.confirmed = True
                print(f"[CALIB] Loaded from {path}")
                return True
        except Exception:
            pass
        return False
    # -- draw -----------------------------------------------------------
    def draw_overlay(self, frame: np.ndarray):
        h, w = frame.shape[:2]
        n    = len(self.img_pts)
        ph   = 22 + len(self.INSTRUCTIONS) * 22
        cv2.rectangle(frame, (0,0), (w, ph), (15,15,15), -1)
        for i, line in enumerate(self.INSTRUCTIONS):
            col = (80,220,80) if i == 0 else (170,170,170)
            cv2.putText(frame, line, (12, 20+i*22),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.62, col, 2, cv2.LINE_AA)
        cv2.putText(frame, f"Setup 1/3  4 now, 8 total  Points: {n}/4  (RClick=undo)",
                (w-520, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.68, C_CALIB, 2, cv2.LINE_AA)
        for i, p in enumerate(self.img_pts):
            pi = (int(p[0]), int(p[1]))
            cv2.circle(frame, pi, 9, C_CALIB, -1)
            cv2.circle(frame, pi, 9, (255,255,255), 1)
            cv2.putText(frame, self.LABELS[i], (pi[0]+11, pi[1]-7),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.56, C_CALIB, 2, cv2.LINE_AA)
        pairs = [(0,1),(0,2),(2,3),(1,3)]
        for a,b in pairs:
            if a < n and b < n:
                pa = (int(self.img_pts[a][0]), int(self.img_pts[a][1]))
                pb = (int(self.img_pts[b][0]), int(self.img_pts[b][1]))
                cv2.line(frame, pa, pb, C_CALIB, 1, cv2.LINE_AA)
        if self._hover and n < 4:
            if n > 0:
                cv2.line(frame,
                         (int(self.img_pts[-1][0]), int(self.img_pts[-1][1])),
                         self._hover, (90,90,90), 1, cv2.LINE_AA)
            cv2.circle(frame, self._hover, 5, (180,180,180), -1)
            cv2.putText(frame, f"-> {self.LABELS[n]}",
                        (self._hover[0]+9, self._hover[1]-8),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.50, (180,180,180), 1, cv2.LINE_AA)
# ---------------------------------------------------------------------------
class RightBoundary:
    """
    The rightmost road edge line -- all offsets measured from here.
    Click 2 points along the road edge.
    """
    def __init__(self):
        self.img_pts:  List[Tuple[float,float]] = []
        self._origin:  Optional[np.ndarray] = None
        self._tangent: Optional[np.ndarray] = None
        self._normal:  Optional[np.ndarray] = None
        self.confirmed = False
        self._hover    = None
    def mouse_cb(self, event, x, y, flags, param):
        if self.confirmed: return
        if event == cv2.EVENT_MOUSEMOVE:
            self._hover = (x, y)
        elif event == cv2.EVENT_LBUTTONDOWN:
            if len(self.img_pts) < 2:
                self.img_pts.append((float(x), float(y)))
                if len(self.img_pts) == 2:
                    self.confirmed = True
    def project(self, calib: Calibrator):
        bev = [calib.to_bev(p) for p in self.img_pts]
        p1, p2 = np.array(bev[0], float), np.array(bev[1], float)
        d = p2 - p1; L = np.linalg.norm(d)
        if L < 1e-9: return
        self._tangent = d / L
        self._normal  = np.array([-self._tangent[1], self._tangent[0]])
        self._origin  = p1
    def signed_offset(self, bev_pt) -> float:
        """Signed distance from boundary (+ve = further from boundary = inside road)."""
        if self._origin is None: return 0.0
        v = np.array(bev_pt, float) - self._origin
        return float(np.dot(v, self._normal))
    def foot_bev(self, bev_pt) -> np.ndarray:
        v = np.array(bev_pt, float) - self._origin
        t = np.dot(v, self._tangent)
        return self._origin + t * self._tangent
    def reset(self): self.__init__()
    def draw(self, frame: np.ndarray, calib: Calibrator):
        if self._origin is None: return
        BIG = 2000.0
        for off, col, lbl in [
            (0.0,   C_REF,    "RIGHT BOUNDARY (Ref=0)"),
        ]:
            shift = self._normal * off
            b1 = self._origin + shift - self._tangent * BIG
            b2 = self._origin + shift + self._tangent * BIG
            q1 = calib.to_img(b1).astype(int)
            q2 = calib.to_img(b2).astype(int)
            cv2.line(frame, tuple(q1), tuple(q2), col, 2, cv2.LINE_AA)
            mid = calib.to_img(self._origin + shift).astype(int)
            cv2.putText(frame, lbl, (mid[0]+6, mid[1]-8),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.48, col, 1, cv2.LINE_AA)
        for p in self.img_pts:
            cv2.circle(frame, (int(p[0]), int(p[1])), 6, C_REF, -1)
    def draw_selection(self, frame: np.ndarray, calib: Calibrator):
        h, w = frame.shape[:2]
        n = len(self.img_pts)
        msg = ("Click POINT 1 on the RIGHT road boundary"
               if n == 0 else "Click POINT 2 on the RIGHT road boundary")
        cv2.rectangle(frame, (0, h-44), (w, h), (15,15,15), -1)
        cv2.putText(frame, msg, (12, h-16),
                  cv2.FONT_HERSHEY_SIMPLEX, 0.72, (120,255,160), 2, cv2.LINE_AA)
        for p in self.img_pts:
            cv2.circle(frame, (int(p[0]), int(p[1])), 7, C_REF, -1)
        if n == 1 and self._hover:
            cv2.line(frame,
                     (int(self.img_pts[0][0]), int(self.img_pts[0][1])),
                     self._hover, C_REF, 1, cv2.LINE_AA)
# ---------------------------------------------------------------------------
class Tripwire:
    """A measurement gate line across the road."""
    def __init__(self):
        self.img_pts: List[Tuple[float,float]] = []
        self._p1 = self._p2 = None
        self.confirmed = False
        self._hover = None
    def mouse_cb(self, event, x, y, flags, param):
        if self.confirmed: return
        if event == cv2.EVENT_MOUSEMOVE: self._hover = (x, y)
        elif event == cv2.EVENT_LBUTTONDOWN:
            if len(self.img_pts) < 2:
                self.img_pts.append((float(x), float(y)))
                if len(self.img_pts) == 2: self.confirmed = True
    def project(self, calib: Calibrator):
        bev = [calib.to_bev(p) for p in self.img_pts]
        self._p1 = np.array(bev[0], float)
        self._p2 = np.array(bev[1], float)
    def distance_to_gate(self, bev_pt) -> float:
        if self._p1 is None or self._p2 is None: return 999.0
        A, B, C = self._p1, self._p2, np.array(bev_pt, float)
        lv = B - A; ll = np.linalg.norm(lv)
        if ll < 1e-9: return 999.0
        t  = np.clip(np.dot(C - A, lv) / ll**2, 0, 1)
        return float(np.linalg.norm(C - (A + t*lv)))
    def crossed(self, prev_bev, curr_bev) -> bool:
        if self._p1 is None: return False
        def ccw(A, B, C):
            return (C[1]-A[1])*(B[0]-A[0]) > (B[1]-A[1])*(C[0]-A[0])
        A, B, C, D = self._p1, self._p2, np.array(prev_bev,float), np.array(curr_bev,float)
        if ccw(A,C,D) != ccw(B,C,D) and ccw(A,B,C) != ccw(A,B,D):
            return True
        # proximity fallback
        return self.distance_to_gate(curr_bev) < 0.75
    def reset(self): self.__init__()
    def draw(self, frame: np.ndarray, calib: Calibrator):
        if len(self.img_pts) < 2 or self._p1 is None: return
        q1 = calib.to_img(self._p1).astype(int)
        q2 = calib.to_img(self._p2).astype(int)
        cv2.line(frame, tuple(q1), tuple(q2), C_WIRE, 3, cv2.LINE_AA)
        mid = ((q1[0]+q2[0])//2, (q1[1]+q2[1])//2)
        cv2.putText(frame, "MEASUREMENT GATE", (mid[0]+10, mid[1]-6),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, C_WIRE, 2, cv2.LINE_AA)
    def draw_selection(self, frame: np.ndarray):
        h, w = frame.shape[:2]
        n   = len(self.img_pts)
        msg = ("Click POINT 1 -- Measurement Gate (left)"
               if n == 0 else "Click POINT 2 -- Measurement Gate (right)")
        cv2.rectangle(frame, (0, h-44), (w, h), (10,10,40), -1)
        cv2.putText(frame, msg, (12, h-16),
                  cv2.FONT_HERSHEY_SIMPLEX, 0.72, (160,160,255), 2, cv2.LINE_AA)
        for p in self.img_pts:
            cv2.circle(frame, (int(p[0]), int(p[1])), 7, C_WIRE, -1)
        if n == 1 and self._hover:
            cv2.line(frame,
                     (int(self.img_pts[0][0]), int(self.img_pts[0][1])),
                     self._hover, C_WIRE, 2, cv2.LINE_AA)
# ---------------------------------------------------------------------------
class WheelLocalizer:
    """
    Estimates left & right wheel contact points for a vehicle bounding box.
    Method (no extra model needed):
      1. Crop bottom-third of the bounding box (where wheels live).
      2. Convert to HSV; create a dark-region mask (tire = dark).
      3. Find the leftmost and rightmost dark blobs -> wheel centres.
      4. Fall back to geometric estimation if blobs not found.
    """
    def __init__(self, dark_thresh: int = 60, min_area: int = 150):
        self.dark_thresh = dark_thresh
        self.min_area    = min_area
    def locate(self, frame: np.ndarray,
               x1: int, y1: int, x2: int, y2: int, vtype: str = ""
               ) -> Tuple[Tuple[int,int], Tuple[int,int]]:
        """Return (left_wheel_px, right_wheel_px) in full-image coordinates."""
        bh = y2 - y1
        bw = x2 - x1
        if vtype in ("bicycle", "motorcycle", "rider", "person") or bw < 40:
            wy = max(y1 + 5, y2 - max(3, int(bh * 0.05)))
            return (x1 + int(bw * 0.3), wy), (x1 + int(bw * 0.7), wy)
        crop_y1 = y1 + int(bh * 0.68)   # bottom 32 % of box
        crop = frame[crop_y1:y2, x1:x2]
        if crop.size == 0:
            return self._geometric(x1, y1, x2, y2)
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        _, dark = cv2.threshold(gray, self.dark_thresh, 255, cv2.THRESH_BINARY_INV)
        # morphological clean
        k  = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5,5))
        dark = cv2.morphologyEx(dark, cv2.MORPH_CLOSE, k)
        dark = cv2.morphologyEx(dark, cv2.MORPH_OPEN,  k)
        contours, _ = cv2.findContours(dark, cv2.RETR_EXTERNAL,
                                        cv2.CHAIN_APPROX_SIMPLE)
        blobs = []
        for c in contours:
            area = cv2.contourArea(c)
            if area < self.min_area: continue
            M = cv2.moments(c)
            if M["m00"] == 0: continue
            cx = int(M["m10"] / M["m00"]) + x1
            cy = int(M["m01"] / M["m00"]) + crop_y1
            blobs.append((cx, cy, area))
        if len(blobs) >= 2:
            blobs.sort(key=lambda b: b[0])     # sort by x
            left  = (blobs[0][0],  blobs[0][1])
            right = (blobs[-1][0], blobs[-1][1])
            return left, right
        elif len(blobs) == 1:
            # Only one blob -- use geometric for other side
            cx, cy = blobs[0][0], blobs[0][1]
            mid_x  = (x1 + x2) // 2
            if cx <= mid_x:
                return (cx, cy), (x2 - (mid_x - cx), cy)
            else:
                return (x1 + (cx - mid_x), cy), (cx, cy)
        return self._geometric(x1, y1, x2, y2)
    @staticmethod
    def _geometric(x1, y1, x2, y2) -> Tuple[Tuple[int,int], Tuple[int,int]]:
        """Fall-back: 20 % and 80 % of width, 5 px above bottom."""
        bw = x2 - x1
        wy = y2 - 5
        return (x1 + bw // 5, wy), (x2 - bw // 5, wy)
# ---------------------------------------------------------------------------
class LaneGrid:
    """Overlay grid showing distance from right boundary."""
    @staticmethod
    def draw(frame, boundary: RightBoundary, calib: Calibrator,
             lane_w: float, road_w: float):
        if boundary._origin is None: return
        overlay = frame.copy()
        BIG = 2000.0
        n = int(round(road_w / lane_w)) + 1
        for i in range(n):
            d     = i * lane_w
            shift = boundary._normal * d
            b1    = boundary._origin + shift - boundary._tangent * BIG
            b2    = boundary._origin + shift + boundary._tangent * BIG
            q1    = calib.to_img(b1).astype(int)
            q2    = calib.to_img(b2).astype(int)
            thick = 2 if i == 0 else 1
            col   = C_REF if i == 0 else C_GRID
            cv2.line(overlay, tuple(q1), tuple(q2), col, thick, cv2.LINE_AA)
            lbl_img = calib.to_img(boundary._origin + shift).astype(int)
            lbl_img = np.clip(lbl_img, [5,5],
                              [frame.shape[1]-70, frame.shape[0]-10])
            cv2.putText(overlay, f"{d:.1f}m",
                        tuple(lbl_img), cv2.FONT_HERSHEY_SIMPLEX,
                        0.42, (110,210,110), 1, cv2.LINE_AA)
        cv2.addWeighted(overlay, 0.4, frame, 0.6, 0, frame)
# ---------------------------------------------------------------------------
class Record:
    __slots__ = ("video","frame","ts","vid_id","vtype",
                 "lw_x","rw_x","lw_off","rw_off","ctr_off","conf",
                 "lw_img","rw_img")
    def __init__(self, video, frame, ts, vid_id, vtype,
                 lw_x, rw_x, lw_off, rw_off, ctr_off, conf,
                 lw_img=None, rw_img=None):
        self.video   = video
        self.frame   = frame
        self.ts      = ts
        self.vid_id  = vid_id
        self.vtype   = vtype
        self.lw_x    = round(lw_x,  4)
        self.rw_x    = round(rw_x,  4)
        self.lw_off  = round(lw_off, 4)
        self.rw_off  = round(rw_off, 4)
        self.ctr_off = round(ctr_off, 4)
        self.conf    = round(conf,  4)
        self.lw_img  = lw_img   # pixel coords for drawing
        self.rw_img  = rw_img
# ---------------------------------------------------------------------------
class ResultsLog:
    def __init__(self):
        self.records: List[Record] = []
    def add(self, rec: Record):
        self.records.append(rec)
    # -- CSV ------------------------------------------------------------
    def export_csv(self, path: str):
        fields = ["Video","Frame","Timestamp(s)","VehicleID","VehicleType",
                  "LeftWheelOffset(m)","RightWheelOffset(m)","CenterOffset(m)",
                  "LeftWheelX_BEV(m)","RightWheelX_BEV(m)","Confidence"]
        with open(path, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            for r in self.records:
                w.writerow({
                    "Video":                r.video,
                    "Frame":                r.frame,
                    "Timestamp(s)":         round(r.ts, 3),
                    "VehicleID":            r.vid_id,
                    "VehicleType":          r.vtype,
                    "LeftWheelOffset(m)":   r.lw_off,
                    "RightWheelOffset(m)":  r.rw_off,
                    "CenterOffset(m)":      r.ctr_off,
                    "LeftWheelX_BEV(m)":    r.lw_x,
                    "RightWheelX_BEV(m)":   r.rw_x,
                    "Confidence":           r.conf,
                })
        print(f"[CSV] -> {path}")
    # -- Excel (rich) ---------------------------------------------------
    def export_xlsx(self, path: str):
        if not XLSX_OK:
            print("[WARN] openpyxl not found -- falling back to CSV")
            self.export_csv(path.replace(".xlsx", ".csv"))
            return
        wb = openpyxl.Workbook()
        # - Sheet 1: Raw Data -----------------------------------------
        ws = wb.active; ws.title = "Raw Measurements"
        thin   = Side(style="thin", color="CCCCCC")
        bdr    = Border(left=thin, right=thin, top=thin, bottom=thin)
        H_FILL = PatternFill("solid", start_color="1A2744")
        A_FILL = PatternFill("solid", start_color="EEF2FA")
        W_FILL = PatternFill("solid", start_color="FFFFFF")
        T_FILL = PatternFill("solid", start_color="FFF8E7")  # truck highlight
        headers = ["#","Video","Frame","Time(s)","Vehicle ID","Type",
                   "L-Wheel Off.(m)","R-Wheel Off.(m)","Center Off.(m)",
                   "LW BEV-X(m)","RW BEV-X(m)","Confidence"]
        col_w   = [4, 22, 7, 9, 10, 11, 14, 14, 13, 12, 12, 11]
        # title row
        ws.merge_cells(f"A1:{chr(64+len(headers))}1")
        tc = ws["A1"]
        tc.value = "Wheel Wander & Lateral Offset Measurements -- BTP IIT Kharagpur"
        tc.font  = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
        tc.fill  = PatternFill("solid", start_color="0D1B40")
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
        # header row
        ws.append(headers)
        for c_i, (h, cw) in enumerate(zip(headers, col_w), 1):
            cell = ws.cell(row=2, column=c_i)
            cell.font  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
            cell.fill  = H_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = bdr
            ws.column_dimensions[chr(64+c_i)].width = cw
        # data rows
        for i, r in enumerate(self.records):
            row_data = [
                i+1, r.video, r.frame, round(r.ts,2), r.vid_id, r.vtype,
                r.lw_off, r.rw_off, r.ctr_off,
                r.lw_x, r.rw_x, r.conf
            ]
            ws.append(row_data)
            is_truck = r.vtype in TRUCK_CLASSES
            fill = T_FILL if is_truck else (A_FILL if i%2==0 else W_FILL)
            for c_i in range(1, len(headers)+1):
                cell = ws.cell(row=3+i, column=c_i)
                cell.font   = Font(name="Calibri", size=10,
                                   bold=True if is_truck else False)
                cell.fill   = fill
                cell.border = bdr
                cell.alignment = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.number_format = "0.0000"
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{chr(64+len(headers))}2"
        # - Sheet 2: Statistics ----------------------------------------
        ws2 = wb.create_sheet("Statistics")
        if PANDAS_OK and self.records:
            import pandas as _pd
            import numpy as _np
            df = _pd.DataFrame([{
                "type":    r.vtype,
                "lw_off":  r.lw_off,
                "rw_off":  r.rw_off,
                "ctr_off": r.ctr_off,
            } for r in self.records])
            ws2["A1"] = "Statistical Summary"
            ws2["A1"].font  = Font(bold=True, size=12, color="0D1B40")
            ws2["A1"].alignment = Alignment(horizontal="center")
            ws2.merge_cells("A1:F1")
            row = 3
            for vt in ["All"] + sorted(df["type"].unique().tolist()):
                sub = df if vt == "All" else df[df["type"] == vt]
                if sub.empty: continue
                ws2.cell(row=row, column=1, value=f"[{vt.upper()}]").font = Font(bold=True, color="1A2744")
                row += 1
                for col_name, series_name in [
                    ("lw_off",  "Left-Wheel Offset (m)"),
                    ("rw_off",  "Right-Wheel Offset (m)"),
                    ("ctr_off", "Center Offset (m)"),
                ]:
                    s = sub[col_name]
                    stats = {
                        "Metric": series_name,
                        "Count":  len(s),
                        "Mean":   round(s.mean(), 4),
                        "Std":    round(s.std(), 4),
                        "Min":    round(s.min(), 4),
                        "Max":    round(s.max(), 4),
                    }
                    if row == 4:
                        ws2.append(list(stats.keys()))
                        for c_i in range(1,7):
                            cell = ws2.cell(row=row-1, column=c_i)
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = H_FILL; cell.border = bdr
                            cell.alignment = Alignment(horizontal="center")
                    ws2.append(list(stats.values()))
                    for c_i in range(1,7):
                        ws2.cell(row=row, column=c_i).border = bdr
                    row += 1
                row += 1   # blank between vehicle types
            ws2.column_dimensions["A"].width = 24
            for col in "BCDEF":
                ws2.column_dimensions[col].width = 12
        wb.save(path)
        print(f"[XLSX] -> {path}")
# ---------------------------------------------------------------------------
# -- Drawing helpers --------------------------------------------------------
def draw_offset_ruler(frame, foot_img, wheel_img, offset_m, color):
    """Dashed line + numeric ruler from wheel to boundary."""
    p1, p2 = np.array(foot_img, float), np.array(wheel_img, float)
    n_seg = 16
    for i in range(n_seg):
        a = (p1 + (p2-p1)*(i/n_seg)).astype(int)
        b = (p1 + (p2-p1)*((i+0.5)/n_seg)).astype(int)
        cv2.line(frame, tuple(a), tuple(b), color, 1, cv2.LINE_AA)
    mid = ((p1+p2)/2).astype(int)
    txt = f"{offset_m*100:.1f}cm"
    (tw, th), _ = cv2.getTextSize(txt, cv2.FONT_HERSHEY_SIMPLEX, 0.44, 1)
    cv2.rectangle(frame, (mid[0]-2, mid[1]-th-4),
                  (mid[0]+tw+4, mid[1]+2), (15,15,15), -1)
    cv2.putText(frame, txt, (mid[0], mid[1]),
                cv2.FONT_HERSHEY_SIMPLEX, 0.54, color, 2, cv2.LINE_AA)
def draw_vehicle(frame, x1, y1, x2, y2, track_id, vtype, conf,
                 lw, rw, lw_off, rw_off, ctr_off,
                 boundary: RightBoundary, calib: Calibrator,
                 triggered: bool, W: int, H: int):
    box_col = (0,220,80) if triggered else C_BOX
    thick   = 3 if vtype in TRUCK_CLASSES else 2
    cv2.rectangle(frame, (x1,y1), (x2,y2), box_col, thick)
    # Wheel markers
    cv2.circle(frame, lw, 7, C_WHEEL_L, -1)
    cv2.circle(frame, lw, 7, (255,255,255), 1)
    cv2.circle(frame, rw, 7, C_WHEEL_R, -1)
    cv2.circle(frame, rw, 7, (255,255,255), 1)
    # Centre track dot
    cx = (lw[0]+rw[0])//2; cy = (lw[1]+rw[1])//2
    cv2.circle(frame, (cx,cy), 4, C_CENTER, -1)
    # Offset rulers to reference line (only if calibrated)
    if boundary._origin is not None:
        lw_bev  = calib.to_bev(lw)
        rw_bev  = calib.to_bev(rw)
        lf_bev  = boundary.foot_bev(lw_bev)
        rf_bev  = boundary.foot_bev(rw_bev)
        lf_img  = calib.to_img(lf_bev).astype(int)
        rf_img  = calib.to_img(rf_bev).astype(int)
        draw_offset_ruler(frame, tuple(lf_img), lw, abs(lw_off), C_WHEEL_L)
        draw_offset_ruler(frame, tuple(rf_img), rw, abs(rw_off), C_WHEEL_R)
    # Label pill
    label = (f"ID:{track_id}  {vtype}  "
             f"L:{lw_off*100:.1f}cm  R:{rw_off*100:.1f}cm  "
             f"Ctr:{ctr_off*100:.1f}cm  ({conf:.0%})")
    if triggered: label += " [OK]"
    (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.46, 1)
    lx = int(np.clip(x1, 0, W-tw-10))
    ly = max(0, y1 - th - 10)
    cv2.rectangle(frame, (lx,ly), (lx+tw+8, ly+th+8), (10,10,10), -1)
    cv2.putText(frame, label, (lx+4, ly+th+2),
                cv2.FONT_HERSHEY_SIMPLEX, 0.56, C_LABEL, 2, cv2.LINE_AA)
def draw_hud(frame, log: ResultsLog, paused: bool,
             fps_disp: float, video_name: str):
    trucks = [r for r in log.records if r.vtype in TRUCK_CLASSES]
    lines = [
        f"Video: {video_name}   Events: {len(log.records)}   "
        f"Trucks/Buses: {len(trucks)}",
        (f"Trucks -- Mean Ctr Offset: "
         f"{sum(r.ctr_off for r in trucks)/max(1,len(trucks))*100:.1f} cm  "
         f"Wander -: "
         f"{(np.std([r.ctr_off for r in trucks])*100 if trucks else 0):.1f} cm"
         if trucks else "Trucks -- no data yet"),
        f"FPS: {fps_disp:.1f}   "
        + ("|| PAUSED" if paused else "> RUNNING")
        + "   [SPACE=pause  E=export  R=recalib  Q=quit]",
    ]
    for i, txt in enumerate(lines):
        col = (0, 160, 255) if "PAUSED" in txt else (210,210,210)
        y   = 20 + i * 22
        cv2.putText(frame, txt, (10, y), cv2.FONT_HERSHEY_SIMPLEX,
                    0.60, (0,0,0), 5, cv2.LINE_AA)
        cv2.putText(frame, txt, (10, y), cv2.FONT_HERSHEY_SIMPLEX,
                    0.60, col, 2, cv2.LINE_AA)
# ---------------------------------------------------------------------------
# -- Interactive setup phases -----------------------------------------------
def phase_calibration(win, frame, calib: Calibrator, cap):
    cv2.setMouseCallback(win, calib.mouse_cb)
    while not calib.confirmed:
        disp = frame.copy(); calib.draw_overlay(disp)
        cv2.imshow(win, disp)
        k = cv2.waitKey(15) & 0xFF
        if k in (ord('q'), 27):
            cap.release(); cv2.destroyAllWindows(); sys.exit(0)
    cv2.setMouseCallback(win, lambda *a: None)
def phase_boundary(win, frame, bnd: RightBoundary,
                   calib: Calibrator, cap):
    cv2.setMouseCallback(win, bnd.mouse_cb)
    while not bnd.confirmed:
        disp = frame.copy()
        bnd.draw_selection(disp, calib)
        cv2.imshow(win, disp)
        k = cv2.waitKey(15) & 0xFF
        if k in (ord('q'), 27):
            cap.release(); cv2.destroyAllWindows(); sys.exit(0)
    cv2.setMouseCallback(win, lambda *a: None)
    bnd.project(calib)
def phase_tripwire(win, frame, wire: Tripwire,
                   calib: Calibrator, cap):
    cv2.setMouseCallback(win, wire.mouse_cb)
    while not wire.confirmed:
        disp = frame.copy()
        wire.draw_selection(disp)
        cv2.imshow(win, disp)
        k = cv2.waitKey(15) & 0xFF
        if k in (ord('q'), 27):
            cap.release(); cv2.destroyAllWindows(); sys.exit(0)
    cv2.setMouseCallback(win, lambda *a: None)
    wire.project(calib)
# ---------------------------------------------------------------------------
def run(source: str, road_width_m: float, lane_width_m: float,
        rect_width_m: float, rect_length_m: float,
        model_name: str, output_path: str,
        conf_thresh: float, calib_file: str, output_fps: float,
        force_calib: bool, iou_thresh: float = 0.50, imgsz: int = 1280,
        include_riders: bool = True):
    if not YOLO_OK:
        sys.exit("[ERROR] ultralytics not installed -- run: pip install ultralytics")
    video_name = Path(source).stem
    print(f"\n{'='*60}")
    print(f"  Wheel Wander Pipeline v4.0  |  IIT Kharagpur BTP")
    print(f"  Video : {source}")
    print(f"  Model : {model_name}")
    print(f"{'='*60}\n")
    # -- load model -----------------------------------------------------
    model = YOLO(f"{model_name}.pt")
    # -- open video -----------------------------------------------------
    src  = int(source) if source.isdigit() else source
    cap  = cv2.VideoCapture(src)
    if not cap.isOpened():
        sys.exit(f"[ERROR] Cannot open: {source}")
    fps  = float(output_fps)
    VW   = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    VH   = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    print(f"[INFO] {VW}-{VH}  @{fps:.1f}fps  frames-{total_frames}")
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    writer = cv2.VideoWriter(output_path, fourcc, fps, (VW, VH))
    ret, first_frame = cap.read()
    if not ret: sys.exit("[ERROR] Cannot read first frame.")
    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    # -- init components ------------------------------------------------
    enhancer   = NightEnhancer()
    calib      = Calibrator(rect_width_m, rect_length_m)
    boundary   = RightBoundary()
    wire       = Tripwire()
    localizer  = WheelLocalizer()
    log        = ResultsLog()
    # -- window ---------------------------------------------------------
    cv2.namedWindow(WIN, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(WIN, min(VW, 1440), min(VH, 810))
    # -- calibration (load or interactive) ------------------------------
    if not calib.confirmed:
        phase_calibration(WIN, first_frame, calib, cap)
        if calib_file:
            calib.save(calib_file)
    phase_boundary(WIN, first_frame, boundary, calib, cap)
    phase_tripwire(WIN, first_frame, wire, calib, cap)
    # -- main loop ------------------------------------------------------
    paused       = False
    frame_no     = 0
    saved_n      = 0
    prev_t       = time.perf_counter()
    fps_disp     = 0.0
    history_bev  : Dict[int, np.ndarray] = {}
    triggered_ids: set = set()
    frame        = first_frame.copy()
    while True:
        if not paused:
            ret, frame = cap.read()
            if not ret: break
            frame_no += 1
            ts = cap.get(cv2.CAP_PROP_POS_MSEC) / 1000.0
            # FPS counter
            now     = time.perf_counter()
            fps_disp = 1.0 / max(now - prev_t, 1e-9)
            prev_t  = now
            # Night enhancement
            frame = enhancer.enhance(frame)
            # Overlays
            LaneGrid.draw(frame, boundary, calib, lane_width_m, road_width_m)
            boundary.draw(frame, calib)
            wire.draw(frame, calib)
            # Active vehicle classes to track
            active_classes = {k: v for k, v in VEHICLE_CLASSES.items() if (k != 0 or include_riders)}

            # Low confidence threshold for YOLO so ByteTrack receives candidate boxes for 2-wheelers
            track_conf = min(conf_thresh, 0.12)

            track_kwargs = {
                "persist": True,
                "conf": track_conf,
                "iou": iou_thresh,
                "tracker": "bytetrack.yaml",
                "classes": list(active_classes.keys()),
                "verbose": False
            }
            if imgsz > 0:
                track_kwargs["imgsz"] = imgsz

            # Detection + tracking
            results = model.track(frame, **track_kwargs)[0]
            current_ids: set = set()
            if results.boxes is not None and len(results.boxes) > 0:
                boxes  = results.boxes.xyxy.cpu().numpy()
                ids    = results.boxes.id.cpu().numpy().astype(int) if results.boxes.id is not None else np.arange(1, len(boxes) + 1)
                clss   = results.boxes.cls.cpu().numpy().astype(int)
                confs  = results.boxes.conf.cpu().numpy()
                for box, track_id, cls_id, conf in zip(boxes, ids, clss, confs):
                    if cls_id not in active_classes: continue
                    min_conf = min(conf_thresh, CLASS_CONF_THRESHOLDS.get(cls_id, conf_thresh))
                    if conf < min_conf: continue
                    x1, y1, x2, y2 = map(int, box)
                    vtype = classify_indian_vehicle(cls_id, x2 - x1, y2 - y1, conf)
                    current_ids.add(track_id)
                    # -- wheel locations ---------------------------
                    lw_img, rw_img = localizer.locate(frame, x1, y1, x2, y2, vtype)
                    # -- BEV coords for both wheels + center -------
                    lw_bev  = calib.to_bev(lw_img)
                    rw_bev  = calib.to_bev(rw_img)
                    cx_bev  = (lw_bev + rw_bev) / 2.0
                    lw_off  = boundary.signed_offset(lw_bev)
                    rw_off  = boundary.signed_offset(rw_bev)
                    ctr_off = boundary.signed_offset(cx_bev)
                    # -- tripwire crossing check -------------------
                    crossed = False
                    if track_id not in triggered_ids and track_id in history_bev:
                        crossed = wire.crossed(history_bev[track_id], cx_bev)
                    history_bev[track_id] = cx_bev
                    if crossed:
                        rec = Record(
                            video   = video_name,
                            frame   = frame_no,
                            ts      = ts,
                            vid_id  = track_id,
                            vtype   = vtype,
                            lw_x    = float(lw_bev[0]),
                            rw_x    = float(rw_bev[0]),
                            lw_off  = lw_off,
                            rw_off  = rw_off,
                            ctr_off = ctr_off,
                            conf    = float(conf),
                            lw_img  = lw_img,
                            rw_img  = rw_img,
                        )
                        log.add(rec)
                        triggered_ids.add(track_id)
                        # Flash circle
                        cv2.circle(frame, (int((x1+x2)/2), int((y1+y2)/2)),
                                   40, (0,255,0), 4)
                    triggered = track_id in triggered_ids
                    draw_vehicle(frame, x1, y1, x2, y2,
                                 track_id, vtype, float(conf),
                                 lw_img, rw_img,
                                 lw_off, rw_off, ctr_off,
                                 boundary, calib, triggered, VW, VH)
            # prune dead tracks
            history_bev  = {k: v for k,v in history_bev.items() if k in current_ids}
            triggered_ids = {k for k in triggered_ids if k in current_ids or k in history_bev}
            draw_hud(frame, log, paused, fps_disp, video_name)
            writer.write(frame)
        cv2.imshow(WIN, frame)
        k = cv2.waitKey(1) & 0xFF
        if k in (ord('q'), 27):
            break
        elif k == ord(' '):
            paused = not paused
        elif k == ord('s'):
            saved_n += 1
            fn = str(Path(output_path).parent / f"frame_{saved_n:04d}.png")
            cv2.imwrite(fn, frame)
            print(f"[SAVED] {fn}")
        elif k == ord('e'):
            stem = str(Path(output_path).parent / video_name)
            log.export_xlsx(stem + "_offsets.xlsx")
            log.export_csv(stem + "_offsets.csv")
        elif k == ord('r'):
            calib.reset(); boundary.reset(); wire.reset()
            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ret2, first_frame = cap.read()
            if not ret2: break
            cap.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
            phase_calibration(WIN, first_frame, calib, cap)
            if calib_file: calib.save(calib_file)
            phase_boundary(WIN, first_frame, boundary, calib, cap)
            phase_tripwire(WIN, first_frame, wire, calib, cap)
    cap.release(); writer.release(); cv2.destroyAllWindows()
    # Auto-export on quit
    out_dir = Path(output_path).parent
    stem    = str(out_dir / video_name)
    log.export_xlsx(stem + "_offsets.xlsx")
    log.export_csv(stem + "_offsets.csv")
    print(f"\n[DONE] Processed {frame_no} frames -- {len(log.records)} events logged.")
    print(f"[DONE] Output video -> {output_path}")
    print(f"[DONE] Excel        -> {stem}_offsets.xlsx")
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Wheel Wander & Lateral Offset Pipeline v4.0 -- BTP IIT Kharagpur",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    ap.add_argument("--source",       required=True,
                    help="Path to video file (or camera index)")
    ap.add_argument("--road-width",   type=float, default=7.0,
                    help="Total road width in metres")
    ap.add_argument("--lane-width",   type=float, default=3.5,
                    help="Single lane width in metres")
    ap.add_argument("--rect-width",   type=float, default=7.0,
                    help="Calibration rectangle width (road axis) in metres")
    ap.add_argument("--rect-length",  type=float, default=8.0,
                    help="Calibration rectangle length (road direction) in metres")
    ap.add_argument("--model",        default="yolo11x",
                    help="YOLO model name (e.g. yolo11x, yolov8l, yolov8x)")
    ap.add_argument("--output",       default="output.mp4",
                    help="Output annotated video path")
    ap.add_argument("--conf",         type=float, default=0.20,
                    help="Detection confidence threshold")
    ap.add_argument("--iou",          type=float, default=0.50,
                    help="NMS IoU threshold for overlapping detections")
    ap.add_argument("--imgsz",        type=int, default=1280,
                    help="Inference image resolution (higher resolution identifies small bikes & motorcycles)")
    ap.add_argument("--no-riders",    action="store_true",
                    help="Exclude class 0 (person/rider) from vehicle detection")
    ap.add_argument("--fps",          type=float, default=24.0,
                    help="Output video frame rate")
    ap.add_argument("--calib-file",   default="",
                    help="Path to JSON calibration file (auto-saved/loaded)")
    ap.add_argument("--force-calib", action="store_true",
                    help="Ignore existing calibration file and redo phase 1")
    args = ap.parse_args()
    run(
        source        = args.source,
        road_width_m  = args.road_width,
        lane_width_m  = args.lane_width,
        rect_width_m  = args.rect_width,
        rect_length_m = args.rect_length,
        model_name    = args.model,
        output_path   = args.output,
        conf_thresh   = args.conf,
        calib_file    = args.calib_file,
        output_fps    = args.fps,
        force_calib   = args.force_calib,
        iou_thresh    = args.iou,
        imgsz         = args.imgsz,
        include_riders= not args.no_riders,
    )

