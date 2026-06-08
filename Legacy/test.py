"""
Run this from your Bulk_upload folder:
    python test_validate.py

Place your test Excel files in test_files/ before running.
This script does NOT touch any load logic.
"""

import os
import sys

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "Scripts")
TEST_DIR    = os.path.join(BASE_DIR, "test_files")

sys.path.insert(0, SCRIPTS_DIR)

# ── Map: folder-name-like label -> (validate module name, test filename pattern)
VALIDATORS = {
    "Customer": ("validate_customer", "customer"),
    #"BR":       ("validate_br",       "br"),
}

def find_test_files(label):
    """Find xlsx files in test_files/ that loosely match the entity label."""
    if not os.path.exists(TEST_DIR):
        print(f"  ⚠️  test_files/ folder not found at {TEST_DIR}")
        return []
    matches = [
        os.path.join(TEST_DIR, f)
        for f in os.listdir(TEST_DIR)
        if f.endswith(".xlsx")
    ]
    return matches


def run_test(entity, module_name):
    print(f"\n{'='*50}")
    print(f"  Testing: {entity}")
    print(f"{'='*50}")

    try:
        import importlib
        mod = importlib.import_module(module_name)
    except ImportError as e:
        print(f"  ❌ Could not import {module_name}: {e}")
        return

    files = find_test_files(entity)
    if not files:
        print(f"  ⚠️  No .xlsx files found in test_files/")
        return

    for file_path in files:
        filename = os.path.basename(file_path)
        print(f"\n  📄 File: {filename}")
        try:
            has_errors, error_count = mod.validate(file_path)
            if has_errors:
                print(f"  ❌ Validation FAILED — {error_count} row(s) have errors")
                print(f"     Open the file to see red highlights + Error column")
            else:
                print(f"  ✅ Validation PASSED — all rows clean")
        except Exception as e:
            print(f"  ❌ Script error: {e}")


if __name__ == "__main__":
    print("\n🔍 QAD Validation Test Runner")
    print("   Files will be modified in-place (Error column + highlights)\n")

    for entity, (module_name, _) in VALIDATORS.items():
        run_test(entity, module_name)

    print(f"\n{'='*50}")
    print("  Done. Open your test files to review.")
    print(f"{'='*50}\n")