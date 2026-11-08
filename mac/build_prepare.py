#!/usr/bin/env python3
"""Refresh the explicit source-file permission list in the Mac build module."""

from pathlib import Path


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    code = root / "code"
    builder = root / "mac" / "build_SMFMac.bas"
    names = sorted(
        path.name
        for path in code.iterdir()
        if path.suffix.lower() in {".bas", ".cls"}
        and path.name.lower() != "exportmods.bas"
    )
    if not names or "Dictionary.cls" not in names:
        raise SystemExit("The Mac source modules are missing.")
    if any('"' in name or "|" in name or not name.isascii() for name in names):
        raise SystemExit("A source filename cannot be represented in the VBA permission list.")
    chunks = [names[index:index + 5] for index in range(0, len(names), 5)]
    lines = []
    for index, chunk in enumerate(chunks):
        value = "|".join(chunk)
        if index:
            value = "|" + value
        prefix = "    filenames = " if index == 0 else "    filenames = filenames & "
        lines.append(prefix + '"' + value + '"')
    text = builder.read_text()
    start = "    ' BEGIN GENERATED SOURCE FILES\n"
    end = "    ' END GENERATED SOURCE FILES"
    before, remainder = text.split(start, 1)
    _, after = remainder.split(end, 1)
    updated = before + start + "\n".join(lines) + "\n" + end + after
    if updated != text:
        builder.write_text(updated)
    print(f"The Mac builder requests access to {len(names)} project source files.")


if __name__ == "__main__":
    main()
