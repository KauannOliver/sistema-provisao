from pathlib import Path
import shutil

root = Path(__file__).resolve().parents[1]
for name in ("ProvisaoBD.xlsx", "impostos.json"):
    dest = root / "banco" / name
    source = root / "examples" / ("ProvisaoBD.template.xlsx" if name.endswith("xlsx") else "impostos.template.json")
    if dest.exists():
        print(f"Preserving existing local file: {name}")
    else:
        shutil.copyfile(source, dest)
        print(f"Created empty local input: {name}")
