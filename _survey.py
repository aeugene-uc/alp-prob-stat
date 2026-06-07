import json, re
nb = json.load(open(r"c:\Rei\UC\Semester 4\Statistics\Week 16\alp-prob-stat\AI.ipynb", encoding="utf-8"))
cells = nb["cells"]
print("TOTAL CELLS:", len(cells))
print("\n=== HEADINGS (markdown #) ===")
for i, c in enumerate(cells):
    if c["cell_type"] == "markdown":
        src = "".join(c["source"]) if isinstance(c["source"], list) else c["source"]
        for line in src.splitlines():
            if line.strip().startswith("#"):
                print(f"[{i:>3}] {line.strip()}")

print("\n=== CODE CELLS WITH IMAGE OUTPUTS (plots) ===")
plot_count = 0
for i, c in enumerate(cells):
    if c["cell_type"] == "code":
        has_img = False
        for o in c.get("outputs", []):
            data = o.get("data", {})
            if any(k.startswith("image/") for k in data):
                has_img = True
        if has_img:
            plot_count += 1
            src = "".join(c["source"]) if isinstance(c["source"], list) else c["source"]
            first = src.strip().splitlines()[0] if src.strip() else ""
            # find nearest preceding heading
            title = ""
            for j in range(i, -1, -1):
                if cells[j]["cell_type"] == "markdown":
                    s2 = "".join(cells[j]["source"]) if isinstance(cells[j]["source"], list) else cells[j]["source"]
                    for line in s2.splitlines():
                        if line.strip().startswith("#"):
                            title = line.strip()
                    if title:
                        break
            print(f"[{i:>3}] plot | under: {title[:60]} | code: {first[:50]}")
print("\nTOTAL PLOT CELLS:", plot_count)
print("\n=== ALL CODE CELL FIRST LINES (overview) ===")
for i, c in enumerate(cells):
    if c["cell_type"] == "code":
        src = "".join(c["source"]) if isinstance(c["source"], list) else c["source"]
        lines = [l for l in src.splitlines() if l.strip()]
        first = lines[0] if lines else "(empty)"
        print(f"[{i:>3}] {first[:75]}")
