#!/usr/bin/env python3
"""Read saved VBA component records without opening Excel or executing macros.

This tool reads the CFB directory, FAT, miniFAT and the actual PROJECT stream
inside xl/vbaProject.bin. It does not scan raw sectors for text fragments.
PROJECT records establish the saved component classifications, not whether
the VBA code compiles or executes. Unsaved editor changes are not inspected.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import struct
import sys
import zipfile


class CompoundFileError(ValueError):
    pass


class CompoundFile:
    FREE = 0xFFFFFFFF
    END = 0xFFFFFFFE
    FIRST_RESERVED = 0xFFFFFFFA

    def __init__(self, data: bytes):
        self.data = data
        if len(data) < 512 or data[:8] != bytes.fromhex("d0cf11e0a1b11ae1"):
            raise CompoundFileError("The VBA project is not a CFB file")
        if self.u16(28) != 0xFFFE:
            raise CompoundFileError("Unsupported CFB byte order")
        self.major = self.u16(26)
        shift = self.u16(30)
        if (self.major, shift) not in {(3, 9), (4, 12)} or self.u16(32) != 6:
            raise CompoundFileError("Unsupported CFB sector format")
        self.sector_size = 1 << shift
        self.mini_size = 64
        if len(data) % self.sector_size:
            raise CompoundFileError("The CFB file has an incomplete sector")

        difat = list(struct.unpack_from("<109I", data, 76))
        sid = self.u32(68)
        seen = set()
        for _ in range(self.u32(72)):
            if sid in seen:
                raise CompoundFileError("The DIFAT chain contains a cycle")
            seen.add(sid)
            words = self.words(self.sector(sid))
            difat.extend(words[:-1])
            sid = words[-1]
        fat_count = self.u32(44)
        fat_sectors = [sid for sid in difat if sid != self.FREE][:fat_count]
        if len(fat_sectors) != fat_count or len(set(fat_sectors)) != fat_count:
            raise CompoundFileError("The FAT sector list is incomplete or duplicated")
        self.fat = [value for sid in fat_sectors for value in self.words(self.sector(sid))]

        self.entries = []
        directory = self.regular_stream(self.u32(48))
        for offset in range(0, len(directory), 128):
            entry = directory[offset:offset + 128]
            if len(entry) != 128:
                raise CompoundFileError("The CFB directory has an incomplete entry")
            entry_type = entry[66]
            if not entry_type:
                continue
            name_length = struct.unpack_from("<H", entry, 64)[0]
            if name_length < 2 or name_length > 64 or name_length % 2:
                raise CompoundFileError("Invalid CFB directory name length")
            size = struct.unpack_from("<Q", entry, 120)[0]
            if self.major == 3:
                size &= 0xFFFFFFFF
            self.entries.append({
                "name": entry[:name_length - 2].decode("utf-16le"),
                "type": entry_type,
                "start": struct.unpack_from("<I", entry, 116)[0],
                "size": size,
            })
        roots = [entry for entry in self.entries if entry["type"] == 5]
        if len(roots) != 1:
            raise CompoundFileError("Expected one CFB root entry")
        self.root = roots[0]

    def u16(self, offset: int) -> int:
        return struct.unpack_from("<H", self.data, offset)[0]

    def u32(self, offset: int) -> int:
        return struct.unpack_from("<I", self.data, offset)[0]

    @staticmethod
    def words(data: bytes) -> tuple[int, ...]:
        return struct.unpack("<" + "I" * (len(data) // 4), data)

    def sector(self, sid: int) -> bytes:
        start = (sid + 1) * self.sector_size
        if sid >= self.FIRST_RESERVED or start + self.sector_size > len(self.data):
            raise CompoundFileError(f"Invalid CFB sector {sid}")
        return self.data[start:start + self.sector_size]

    def chain(self, start: int, allocation: list[int]):
        sid = start
        seen = set()
        while sid != self.END:
            if sid in seen or sid >= len(allocation) or sid >= self.FIRST_RESERVED:
                raise CompoundFileError(f"Invalid or cyclic CFB allocation chain at {sid}")
            seen.add(sid)
            yield sid
            sid = allocation[sid]

    def regular_stream(self, start: int, size: int | None = None) -> bytes:
        data = b"".join(self.sector(sid) for sid in self.chain(start, self.fat))
        if size is not None and len(data) < size:
            raise CompoundFileError("A regular stream is shorter than its directory size")
        return data if size is None else data[:size]

    def read_stream(self, name: str) -> bytes:
        matches = [entry for entry in self.entries
                   if entry["type"] == 2 and entry["name"].casefold() == name.casefold()]
        if len(matches) != 1:
            raise CompoundFileError(f"Expected exactly one {name} stream")
        entry = matches[0]
        if entry["size"] >= self.u32(56):
            return self.regular_stream(entry["start"], entry["size"])
        if entry["size"] == 0:
            return b""
        mini_fat_bytes = self.regular_stream(self.u32(60), self.u32(64) * self.sector_size)
        mini_fat = list(self.words(mini_fat_bytes))
        mini_stream = self.regular_stream(self.root["start"], self.root["size"])
        blocks = []
        for sid in self.chain(entry["start"], mini_fat):
            offset = sid * self.mini_size
            if offset + self.mini_size > len(mini_stream):
                raise CompoundFileError("A mini-sector lies outside the root mini-stream")
            blocks.append(mini_stream[offset:offset + self.mini_size])
        data = b"".join(blocks)
        if len(data) < entry["size"]:
            raise CompoundFileError("The mini-stream is shorter than its directory size")
        return data[:entry["size"]]


def project_records(project: bytes, encoding: str = "cp1252") -> dict:
    text = project.decode(encoding)
    records = {"Module": [], "Class": [], "Document": [], "BaseClass": []}
    names = []
    for line in text.splitlines():
        name = re.fullmatch(r'Name="([^"]+)"', line)
        if name:
            names.append(name[1])
        for kind in records:
            if line.startswith(kind + "="):
                records[kind].append(line[len(kind) + 1:])
    if len(names) != 1:
        raise CompoundFileError("Expected one Name record in the PROJECT stream")
    return {
        "projectName": names[0],
        "standardModules": records["Module"],
        "classModules": records["Class"],
        "documentModules": records["Document"],
        "baseClasses": records["BaseClass"],
    }


def inspect(path: Path, encoding: str = "cp1252") -> dict:
    payload = path.read_bytes()
    with zipfile.ZipFile(path) as archive:
        bad_member = archive.testzip()
        if bad_member:
            raise ValueError(f"ZIP integrity check failed for {bad_member}")
        cfb = CompoundFile(archive.read("xl/vbaProject.bin"))
    project = cfb.read_stream("PROJECT")
    result = project_records(project, encoding)
    result.update({
        "path": str(path.resolve()),
        "workbookBytes": len(payload),
        "workbookSHA256": hashlib.sha256(payload).hexdigest(),
        "projectStreamBytes": len(project),
        "projectEncoding": encoding,
        "cfbSectorBytes": cfb.sector_size,
        "counts": {"standardModules": len(result["standardModules"]),
                   "classModules": len(result["classModules"]),
                   "documentModules": len(result["documentModules"]),
                   "baseClasses": len(result["baseClasses"])},
        "scope": "Saved PROJECT stream records only; no VBA compilation or execution.",
    })
    return result


def class_errors(result: dict, expected: list[str]) -> list[str]:
    errors = []
    for name in expected:
        if name in result["classModules"]:
            continue
        classification = "standard module" if name in result["standardModules"] else "absent"
        errors.append(f"Expected class {name}; saved PROJECT record is {classification}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="Saved .xlsm workbook to inspect")
    parser.add_argument("--expect-class", action="append", default=[], metavar="NAME",
                        help="Fail unless PROJECT classifies NAME as a class module; repeatable")
    parser.add_argument("--encoding", default="cp1252", help="PROJECT text encoding; SMF checkpoint names are ASCII")
    args = parser.parse_args()
    try:
        result = inspect(args.workbook, args.encoding)
        errors = class_errors(result, args.expect_class)
        if errors:
            result["validationErrors"] = errors
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 1 if errors else 0
    except (OSError, ValueError, KeyError, UnicodeError, LookupError, zipfile.BadZipFile) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
