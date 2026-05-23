"""Run every build script in the directory, in deterministic order."""
from __future__ import annotations
import importlib.util
import sys
import traceback
from pathlib import Path

HERE = Path(__file__).parent

SCRIPTS = sorted(
    p for p in HERE.glob("build_*.py")
    if not p.name.startswith("_")
)


def run(script: Path) -> tuple[bool, str]:
    spec = importlib.util.spec_from_file_location(script.stem, script)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[script.stem] = mod
    try:
        spec.loader.exec_module(mod)
        mod.main()
        return True, ""
    except SystemExit:
        return True, ""
    except Exception:
        return False, traceback.format_exc()


def main():
    ok = 0
    fail = []
    for s in SCRIPTS:
        sys.path.insert(0, str(HERE))
        success, err = run(s)
        if success:
            ok += 1
        else:
            fail.append((s.name, err))
            print(f"FAILED: {s.name}\n{err}")
    print(f"\n=== {ok} of {len(SCRIPTS)} build scripts ran ===")
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    main()
