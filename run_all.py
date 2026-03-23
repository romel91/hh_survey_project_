"""
Run all three pipeline scripts in order.
Execute from the project root: python run_all.py
"""
import subprocess
import sys
import time

scripts = [
    'scripts/01_data_cleaning.py',
    'scripts/02_eda_analysis.py',
    'scripts/03_excel_builder.py',
]

for script in scripts:
    print(f"\n{'=' * 55}")
    print(f"  Running {script}")
    print(f"{'=' * 55}\n")
    start = time.time()
    result = subprocess.run([sys.executable, script])
    elapsed = time.time() - start
    if result.returncode != 0:
        print(f"\n✗  {script} failed (exit code {result.returncode}). Stopping.")
        sys.exit(result.returncode)
    print(f"  Completed in {elapsed:.1f}s")

print("\n✓  All scripts completed successfully.\n")
