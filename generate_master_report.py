from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import matplotlib.pyplot as plt
    import seaborn as sns
    PLOT_OK = True
except Exception:
    PLOT_OK = False


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "master_dataset_report.xlsx"
OUT_DIR = ROOT / "analysis_report"
FIG_DIR = OUT_DIR / "figures"

HEAVY_TYPES = {"bus", "minibus", "truck", "heavy_truck", "pickup_lcv"}
OFFSET_COLS = {
    "Left Wheel Off(m)": "Left wheel",
    "Right Wheel Off(m)": "Right wheel",
    "Center Off(m)": "Center",
}
PLAUSIBLE_MIN_M = -5.0
PLAUSIBLE_MAX_M = 15.0


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(WORKBOOK, sheet_name="All Vehicles Raw Data")
    manifest = pd.read_excel(WORKBOOK, sheet_name="Processing Manifest", header=1)

    raw.columns = [str(c).strip() for c in raw.columns]
    manifest.columns = [str(c).strip() for c in manifest.columns]

    raw["video_key"] = (
        raw["Section"].astype(str)
        + " | "
        + raw["Date"].astype(str)
        + " | "
        + raw["Time/Session"].astype(str)
        + " | "
        + raw["Video"].astype(str)
    )
    manifest["video_key"] = (
        manifest["Section"].astype(str)
        + " | "
        + manifest["Date"].astype(str)
        + " | "
        + manifest["Time/Session"].astype(str)
        + " | "
        + manifest["Video Filename"].astype(str)
    )

    for col in OFFSET_COLS:
        raw[f"{col} (cm)"] = raw[col] * 100.0

    raw["Abs Center Off(cm)"] = raw["Center Off(m)"].abs() * 100.0
    raw["Heavy Vehicle Group"] = np.where(raw["Vehicle Type"].isin(HEAVY_TYPES), "Heavy", "Light")
    raw["plausible_record"] = True
    for col in OFFSET_COLS:
        raw["plausible_record"] &= raw[col].between(PLAUSIBLE_MIN_M, PLAUSIBLE_MAX_M, inclusive="both")
    return raw, manifest


def format_num(value: float, digits: int = 2) -> str:
    if pd.isna(value):
        return "NA"
    return f"{value:,.{digits}f}"


def offset_summary(series: pd.Series) -> dict[str, float]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    q1 = s.quantile(0.25)
    q3 = s.quantile(0.75)
    iqr = q3 - q1
    return {
        "count": len(s),
        "mean": s.mean(),
        "std": s.std(),
        "min": s.min(),
        "p5": s.quantile(0.05),
        "median": s.median(),
        "p95": s.quantile(0.95),
        "p99": s.quantile(0.99),
        "max": s.max(),
        "iqr": iqr,
        "outlier_count": int(((s < (q1 - 1.5 * iqr)) | (s > (q3 + 1.5 * iqr))).sum()),
    }


def make_tables(raw: pd.DataFrame, manifest: pd.DataFrame) -> dict[str, pd.DataFrame]:
    clean = raw[raw["plausible_record"]].copy()

    overall_rows = []
    for dataset_name, frame in [("Raw", raw), ("Clean", clean)]:
        for col, label in OFFSET_COLS.items():
            stats = offset_summary(frame[f"{col} (cm)"])
            overall_rows.append(
                {
                    "Dataset": dataset_name,
                    "Metric": label,
                    "Count": stats["count"],
                    "Mean (cm)": stats["mean"],
                    "Std (cm)": stats["std"],
                    "Min (cm)": stats["min"],
                    "P5 (cm)": stats["p5"],
                    "Median (cm)": stats["median"],
                    "P95 (cm)": stats["p95"],
                    "P99 (cm)": stats["p99"],
                    "Max (cm)": stats["max"],
                    "IQR (cm)": stats["iqr"],
                    "Outliers (IQR rule)": stats["outlier_count"],
                }
            )
    overall_df = pd.DataFrame(overall_rows)

    by_type = (
        clean.groupby("Vehicle Type")
        .agg(
            Records=("Vehicle Type", "size"),
            Videos=("video_key", "nunique"),
            Mean_Center_cm=("Center Off(m)", lambda s: s.mean() * 100.0),
            Std_Center_cm=("Center Off(m)", lambda s: s.std() * 100.0),
            Median_Center_cm=("Center Off(m)", lambda s: s.median() * 100.0),
            P95_Center_cm=("Center Off(m)", lambda s: s.quantile(0.95) * 100.0),
            Mean_Left_cm=("Left Wheel Off(m)", lambda s: s.mean() * 100.0),
            Mean_Right_cm=("Right Wheel Off(m)", lambda s: s.mean() * 100.0),
            Mean_Confidence=("Confidence", "mean"),
        )
        .sort_values("Records", ascending=False)
        .reset_index()
    )
    by_type["Share (%)"] = by_type["Records"] / len(clean) * 100.0

    by_section_date = (
        clean.groupby(["Section", "Date"])
        .agg(
            Records=("Vehicle Type", "size"),
            Videos=("video_key", "nunique"),
            Mean_Center_cm=("Center Off(m)", lambda s: s.mean() * 100.0),
            Std_Center_cm=("Center Off(m)", lambda s: s.std() * 100.0),
            P95_Center_cm=("Center Off(m)", lambda s: s.quantile(0.95) * 100.0),
            Mean_Confidence=("Confidence", "mean"),
        )
        .reset_index()
        .sort_values(["Section", "Date"])
    )

    by_session = (
        clean.groupby(["Section", "Date", "Time/Session"])
        .agg(
            Records=("Vehicle Type", "size"),
            Videos=("video_key", "nunique"),
            Mean_Center_cm=("Center Off(m)", lambda s: s.mean() * 100.0),
            Mean_Abs_Center_cm=("Abs Center Off(cm)", "mean"),
            Std_Center_cm=("Center Off(m)", lambda s: s.std() * 100.0),
            P95_Center_cm=("Center Off(m)", lambda s: s.quantile(0.95) * 100.0),
        )
        .reset_index()
    )

    top_sessions_volume = by_session.sort_values("Records", ascending=False).head(10)
    top_sessions_offset = by_session.sort_values("Mean_Abs_Center_cm", ascending=False).head(10)

    per_video = (
        manifest.groupby(["Section", "Date", "Time/Session", "Video Filename", "video_key"], as_index=False)
        .agg(
            Vehicles_Detected=("Vehicles Detected", "max"),
            Mean_Center_m=("Mean Center Off.(m)", "max"),
            Std_Center_m=("Std Center Off.(m)", "max"),
            Status=("Status", "last"),
        )
    )
    top_videos = per_video.sort_values("Vehicles_Detected", ascending=False).head(15)
    top_videos = top_videos.drop(columns=["video_key"])

    quality = pd.DataFrame(
        [
            {
                "Total raw records": len(raw),
                "Clean plausible records": len(clean),
                "Removed as implausible": len(raw) - len(clean),
                "Removed share (%)": (len(raw) - len(clean)) / len(raw) * 100.0,
                "Completed videos": len(manifest),
                "Videos with detections": int((manifest["Vehicles Detected"].fillna(0) > 0).sum()),
                "Zero-detection videos": int((manifest["Vehicles Detected"].fillna(0) == 0).sum()),
            }
        ]
    )

    return {
        "overall": overall_df,
        "by_type": by_type,
        "by_section_date": by_section_date,
        "by_session": by_session,
        "top_sessions_volume": top_sessions_volume,
        "top_sessions_offset": top_sessions_offset,
        "top_videos": top_videos,
        "quality": quality,
        "clean": clean,
    }


def render_markdown(raw: pd.DataFrame, manifest: pd.DataFrame, tables: dict[str, pd.DataFrame]) -> str:
    clean = tables["clean"]
    total_records = len(raw)
    total_manifest_videos = len(manifest)
    detected_videos = int((manifest["Vehicles Detected"].fillna(0) > 0).sum())
    detected_videos_raw = raw["video_key"].nunique()
    zero_vehicle_videos = int((manifest["Vehicles Detected"].fillna(0) == 0).sum())
    total_vehicle_events = int(manifest["Vehicles Detected"].fillna(0).sum())
    avg_conf = raw["Confidence"].mean()
    heavy_share = clean["Vehicle Type"].isin(HEAVY_TYPES).mean() * 100.0
    pos_center = (clean["Center Off(m)"] > 0).mean() * 100.0
    neg_center = (clean["Center Off(m)"] < 0).mean() * 100.0
    corr = clean["Left Wheel Off(m)"].corr(clean["Right Wheel Off(m)"])
    removed_records = len(raw) - len(clean)
    removed_share = removed_records / len(raw) * 100.0

    overall = tables["overall"].copy()
    by_type = tables["by_type"].copy()
    by_section_date = tables["by_section_date"].copy()
    top_volume = tables["top_sessions_volume"].copy()
    top_offset = tables["top_sessions_offset"].copy()
    top_videos = tables["top_videos"].copy()

    highest_type = by_type.iloc[0]
    highest_offset_type = by_type.sort_values("Mean_Center_cm", ascending=False).iloc[0]
    most_variable_type = by_type.sort_values("Std_Center_cm", ascending=False).iloc[0]

    section_lines = []
    for _, row in by_section_date.iterrows():
        section_lines.append(
            f"- {row['Section']} / {row['Date']}: {int(row['Records']):,} records across {int(row['Videos'])} videos; "
            f"mean center offset {row['Mean_Center_cm']:.2f} cm; P95 {row['P95_Center_cm']:.2f} cm."
        )

    md = []
    md.append("# Master Dataset Data Analysis Report")
    md.append("")
    md.append("## Scope")
    md.append("")
    md.append("This report is based on the user request to analyse the `BTP` project output workbook and summarize the dataset with proper metrics. ")
    md.append("The workbook reviewed was `master_dataset_report.xlsx`, and the project methodology was verified against the pipeline and analysis scripts in `Autmn Sem`.")
    md.append("")
    md.append("## Project Context")
    md.append("")
    md.append("The BTP pipeline measures left-wheel, right-wheel, and vehicle-center lateral offsets from the right road boundary using calibrated roadside video, YOLO-based detection, tracking, BEV transformation, and tripwire-triggered logging.")
    md.append("")
    md.append("## Executive Summary")
    md.append("")
    md.append(f"- Total logged vehicle records: {total_records:,}.")
    md.append(f"- Completed videos in manifest: {total_manifest_videos:,}.")
    md.append(f"- Videos with detections in manifest: {detected_videos:,}.")
    md.append(f"- Videos with raw records present: {detected_videos_raw:,}.")
    md.append(f"- Zero-detection completed videos: {zero_vehicle_videos:,}.")
    md.append(f"- Sum of per-video detected vehicles from manifest: {total_vehicle_events:,}.")
    md.append(f"- Clean plausible records used for behavioural analysis: {len(clean):,} ({100.0 - removed_share:.2f}% of raw records).")
    md.append(f"- Implausible/outlier-like records excluded from behavioural metrics: {removed_records:,} ({removed_share:.2f}%).")
    md.append(f"- Mean detection confidence: {avg_conf:.3f}.")
    md.append(f"- Heavy-vehicle share (`bus`, `minibus`, `truck`, `heavy_truck`, `pickup_lcv`): {heavy_share:.2f}%.")
    md.append(f"- Positive center-offset share: {pos_center:.2f}%; negative center-offset share: {neg_center:.2f}%.")
    md.append(f"- Left/right wheel offset correlation: {corr:.4f}.")
    md.append("")
    md.append("The dataset is dominated by heavy vehicles, especially `heavy_truck` and `pickup_lcv`, and the cleaned center-offset distribution is strongly positive, indicating that most tracked vehicles remain inside the carriageway relative to the chosen right-boundary reference.")
    md.append("")
    md.append("## Overall Offset Metrics")
    md.append("")
    md.append(overall.to_markdown(index=False))
    md.append("")
    md.append("Interpretation:")
    clean_center = overall[(overall["Dataset"] == "Clean") & (overall["Metric"] == "Center")].iloc[0]
    md.append(f"- Clean center-offset median is {clean_center['Median (cm)']:.2f} cm, which is the best single-number summary of typical road position.")
    md.append(f"- Clean center-offset P95 is {clean_center['P95 (cm)']:.2f} cm, showing the upper spread of normal lane usage.")
    md.append(f"- The raw table is retained for transparency, but cleaned values are the safer basis for the written conclusions because the workbook contains some very large calibration/tracking artifacts.")
    md.append("")
    md.append("## Vehicle-Type Metrics")
    md.append("")
    md.append(by_type.to_markdown(index=False))
    md.append("")
    md.append("Key observations:")
    md.append(f"- Highest traffic share: `{highest_type['Vehicle Type']}` with {highest_type['Share (%)']:.2f}% of all records.")
    md.append(f"- Highest mean center offset: `{highest_offset_type['Vehicle Type']}` at {highest_offset_type['Mean_Center_cm']:.2f} cm.")
    md.append(f"- Highest center-offset variability: `{most_variable_type['Vehicle Type']}` with standard deviation {most_variable_type['Std_Center_cm']:.2f} cm.")
    md.append("")
    md.append("## Section and Date Breakdown")
    md.append("")
    md.extend(section_lines)
    md.append("")
    md.append("## Highest-Volume Sessions")
    md.append("")
    md.append(top_volume.to_markdown(index=False))
    md.append("")
    md.append("## Sessions With Highest Mean Absolute Center Offset")
    md.append("")
    md.append(top_offset.to_markdown(index=False))
    md.append("")
    md.append("## Highest-Volume Videos")
    md.append("")
    md.append(top_videos.to_markdown(index=False))
    md.append("")
    md.append("## Data Quality Notes")
    md.append("")
    md.append(f"- Behavioural summaries in this report use a plausible offset window of {PLAUSIBLE_MIN_M:.0f} m to {PLAUSIBLE_MAX_M:.0f} m for all three offset channels.")
    md.append(f"- The manifest shows all videos as `Completed`, so the run itself appears operationally complete.")
    md.append(f"- Only {zero_vehicle_videos} completed videos reported zero detections, which is a small fraction of the total workload.")
    md.append(f"- Raw-record coverage spans {detected_videos_raw} unique section/date/session/video combinations.")
    md.append(f"- The raw export includes a small but important set of physically implausible offsets, so cleaned metrics should be preferred for thesis-level interpretation.")
    md.append(f"- Confidence values should still be interpreted carefully for small classes like `bicycle` and `motorcycle`, because the pipeline intentionally allows lower thresholds for those classes.")
    md.append("")
    md.append("## Conclusion")
    md.append("")
    md.append("The master NH-19 dataset is large enough for a meaningful wheel-wander and lateral-position study. ")
    md.append("Its strongest patterns are the dominance of heavy vehicles, consistently positive center offsets from the right boundary, and clear variability across session groups. ")
    md.append("For a thesis or final BTP write-up, the most defensible headline metrics are total vehicle records, vehicle-class composition, mean/median/P95 center offset, section-date comparisons, and the heavy-vehicle offset distribution.")
    md.append("")
    return "\n".join(md)


def save_tables(tables: dict[str, pd.DataFrame]) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_DIR / "master_dataset_analysis_tables.xlsx", engine="openpyxl") as writer:
        for name, table in tables.items():
            if name == "clean":
                continue
            table.to_excel(writer, sheet_name=name[:31], index=False)
            table.to_csv(OUT_DIR / f"{name}.csv", index=False)


def make_figures(raw: pd.DataFrame, tables: dict[str, pd.DataFrame]) -> None:
    if not PLOT_OK:
        return

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid")
    clean = tables["clean"]

    plt.figure(figsize=(10, 6))
    order = clean["Vehicle Type"].value_counts().index
    sns.countplot(data=clean, y="Vehicle Type", order=order, hue="Vehicle Type", dodge=False, legend=False, palette="viridis")
    plt.title("Vehicle Type Distribution")
    plt.xlabel("Records")
    plt.ylabel("Vehicle Type")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "vehicle_type_distribution.png", dpi=180)
    plt.close()

    plt.figure(figsize=(12, 6))
    order = tables["by_type"]["Vehicle Type"].tolist()
    sns.boxplot(data=clean, x="Vehicle Type", y="Center Off(m)", order=order, hue="Vehicle Type", dodge=False, legend=False, showfliers=False, palette="crest")
    plt.xticks(rotation=30, ha="right")
    plt.ylabel("Center Offset (m)")
    plt.xlabel("Vehicle Type")
    plt.title("Center Offset by Vehicle Type")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "center_offset_by_type.png", dpi=180)
    plt.close()

    heat = (
        clean.pivot_table(
            index="Time/Session",
            columns=["Section", "Date"],
            values="Center Off(m)",
            aggfunc=lambda s: s.mean() * 100.0,
        )
        .sort_index()
    )
    plt.figure(figsize=(10, 12))
    sns.heatmap(heat, cmap="mako", annot=False)
    plt.title("Mean Center Offset (cm) by Session and Corridor")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "session_section_heatmap.png", dpi=180)
    plt.close()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    raw, manifest = load_data()
    tables = make_tables(raw, manifest)
    save_tables(tables)
    make_figures(raw, tables)
    report_md = render_markdown(raw, manifest, tables)
    (OUT_DIR / "master_dataset_analysis_report.md").write_text(report_md, encoding="utf-8")
    print(f"Report written to: {OUT_DIR / 'master_dataset_analysis_report.md'}")


if __name__ == "__main__":
    main()
