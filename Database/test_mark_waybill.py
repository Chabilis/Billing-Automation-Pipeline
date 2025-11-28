import importlib.util
import traceback, json, pathlib

# Load read_waybills.py by path (no package install needed)
rb_path = pathlib.Path(__file__).parent / "read_waybills.py"
spec = importlib.util.spec_from_file_location("read_waybills", str(rb_path))
rb = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rb)

# Regenerate waybills.json/csv
rb.save_waybills_to_files()
wb = rb.get_next_unclaimed_waybill()
print("NEXT_WB:" + str(wb))
try:
    print("ABOUT_TO_MARK")
    res = rb.mark_waybill_used(wb) if wb else None
    print("DONE_MARK")
    print("MARKED:" + str(res))
except Exception:
    print("ERROR_MARKING")
    traceback.print_exc()

# Print first 3 entries from waybills.json
p = pathlib.Path(__file__).parent / "waybills.json"
with open(p, 'r') as f:
    arr = json.load(f)
print("FIRST_ENTRIES:", arr[:3])
