from __future__ import annotations

import fnmatch
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
ZIP_PATH = DIST / "docx_xml_tool_v2_4_2.zip"

EXCLUDED_DIRS = {
    "__pycache__",
    ".pytest_cache",
    "docx_xml_outputs",
    "edits",
    "dist",
    "docs",
}
EXCLUDED_PATTERNS = [
    "*.pyc",
    "*.pyo",
    "*.log",
    "*_docx_xml/*",
    "~$*.docx",
]


def should_exclude(path: Path) -> bool:
    relative = path.relative_to(ROOT)
    parts = set(relative.parts)
    if parts & EXCLUDED_DIRS:
        return True
    rel_text = relative.as_posix()
    return any(fnmatch.fnmatch(rel_text, pattern) for pattern in EXCLUDED_PATTERNS)


def main() -> int:
    DIST.mkdir(exist_ok=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(ROOT.rglob("*")):
            if path.is_dir() or should_exclude(path):
                continue
            archive.write(path, path.relative_to(ROOT.parent).as_posix())
    print(f"Release package written: {ZIP_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
