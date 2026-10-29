"""
Run every Python test suite and report a combined result.

    python tests/run_all.py

(The UI suites run separately with `cd tests && npm install && npm test`.)
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SUITES = [
    "test_html_tables.py",
    "test_backend_fixes.py",
    "test_fidelity.py",
    "test_table_corner_cases.py",
    "test_corner_cases.py",
]

total_pass = total_fail = 0
broken = []
for suite in SUITES:
    env = dict(os.environ, DOCAGENT_DEBUG="0")
    proc = subprocess.run([sys.executable, os.path.join(HERE, suite)], cwd=ROOT,
                          capture_output=True, text=True, env=env)
    out = proc.stdout + proc.stderr
    m = re.search(r"=== (\d+) passed, (\d+) failed ===", out)
    if m:
        p, f = int(m.group(1)), int(m.group(2))
    else:
        p, f = 0, 1
    total_pass += p
    total_fail += f
    status = "ok" if proc.returncode == 0 and f == 0 else "FAILED"
    print(f"{suite:32s} {p:4d} passed {f:3d} failed  {status}")
    if status != "ok":
        broken.append(suite)
        for line in out.splitlines():
            if "[FAIL]" in line or "Traceback" in line or "Error" in line:
                print("    " + line.strip())

print(f"\nTOTAL: {total_pass} passed, {total_fail} failed")
sys.exit(1 if broken else 0)
