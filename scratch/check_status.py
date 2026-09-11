import json
from pathlib import Path

chk_path = Path("outputs/checkpoint.json")
if chk_path.exists():
    data = json.load(open(chk_path, encoding="utf-8"))
    vids = data.get("videos", {})
    print("Total recorded in checkpoint:", len(vids))
    completed = [v for k, v in vids.items() if v.get("status") == "Completed"]
    print("Completed count:", len(completed))
    by_group = {}
    for v in vids.values():
        grp = f"{v.get('section')}/{v.get('date')}"
        st = v.get("status")
        if grp not in by_group:
            by_group[grp] = {}
        by_group[grp][st] = by_group[grp].get(st, 0) + 1

    print("\nStatus breakdown by (Section / Date):")
    for grp, st_map in sorted(by_group.items()):
        print(f"  {grp}: {st_map}")
else:
    print("No checkpoint.json found in outputs/")
