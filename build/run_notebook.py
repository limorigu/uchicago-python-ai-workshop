"""Execute a notebook's code cells sequentially (skipping Colab shell lines) to verify it runs."""
import json, sys, traceback, os, io, contextlib
from pathlib import Path

nb_path = Path(sys.argv[1]).resolve()
os.chdir(nb_path.parent)
cells = [c for c in json.loads(nb_path.read_text())["cells"] if c["cell_type"] == "code"]
ns = {"display": print}
import matplotlib; matplotlib.use("Agg")
for i, c in enumerate(cells):
    src = c["source"] if isinstance(c["source"], str) else "".join(c["source"])
    src = "\n".join((l[:len(l) - len(l.lstrip())] + "pass  # shell line skipped") if l.lstrip().startswith("!") else l
                    for l in src.splitlines())
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exec(compile(src, f"cell{i}", "exec"), ns)
        out = buf.getvalue().strip()
        print(f"--- cell {i} OK" + (f"\n{out[:700]}" if out else ""))
    except Exception:
        print(f"--- cell {i} FAILED\n{src}\n"); traceback.print_exc(); sys.exit(1)
print("ALL CELLS OK:", nb_path.name)
