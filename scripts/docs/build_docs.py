"""Strict MkDocs build into a new output directory; never clear an existing one."""

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--site-dir", type=Path, default=Path("site"))
    args = parser.parse_args()
    target = args.site_dir if args.site_dir.is_absolute() else ROOT / args.site_dir
    # lexists behavior: also reject a dangling symlink rather than following it.
    if target.exists() or target.is_symlink():
        print(f"Refusing existing output path; choose a new directory: {target}", file=sys.stderr)
        return 2
    if not target.resolve().is_relative_to(ROOT):
        print("Output must stay inside the repository.", file=sys.stderr)
        return 2
    return subprocess.call(
        [sys.executable, "-m", "mkdocs", "build", "--strict", "--site-dir", str(target)],
        cwd=ROOT,
    )


if __name__ == "__main__":
    raise SystemExit(main())
