import json
from pathlib import Path

chk_path = Path("outputs/checkpoint.json")
if chk_path.exists():
    with open(chk_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    vids = data.get("videos", {})
    reset_count = 0
    for key, v in list(vids.items()):
        if v.get("section") == "LHS" and v.get("date") == "2nd May":
            if v.get("status") in ("Completed", "In Progress", "Failed"):
                v["status"] = "Pending"
                v["vehicles_detected"] = 0
                v["mean_ctr_off"] = 0.0
                v["std_ctr_off"] = 0.0
                v["error_reason"] = ""
                v["timestamp"] = ""
                v["output_vid"] = ""
                v["output_csv"] = ""
                v["output_xlsx"] = ""
                reset_count += 1
    data["videos"] = vids
    with open(chk_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    print(f"[RESET] Successfully reset {reset_count} videos under 'LHS / 2nd May' to Pending!")
else:
    print("[WARN] outputs/checkpoint.json not found.")
