from pathlib import Path
from openpyxl import Workbook

out = Path(__file__).parent / "generated"
out.mkdir(exist_ok=True)
for name, qty in (("before.xlsx", 2), ("after.xlsx", 5)):
    wb = Workbook(); ws = wb.active; ws.title = "Inventory"
    ws.append(["Item", "Qty"]); ws.append(["Filter", qty]); wb.save(out / name)
print(f"Created samples in {out}")
