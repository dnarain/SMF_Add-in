#!/usr/bin/env python3
"""Read a saved suite workbook. No Excel automation, network, or implicit run state.

Usage: read_results.py --suite DIR --workbook FILE [--output FILE]
Cell values are saved observations, not newly calculated results.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

RAW_STATUSES = frozenset({"PASS", "CHECK_FAILED", "RUNTIME_ERROR", "CALL_ERROR",
                          "CHECK_ERROR", "HARNESS_ERROR", "SETUP_ERROR",
                          "ASSERTION_ERROR", "CLEANUP_ERROR", "PENDING", "RUNNING"})
FIELDS = dict(zip("ABCDEFGHIJKLMNOPQRS", [
    "id", "function", "group", "description", "expected", "status", "seconds",
    "actual", "error", "http_status", "http_url", "http_error", "finished",
    "http_reset_marker", "http_case_id", "http_run_id", "http_observation",
    "fetch_call_delta", "fetchurl_bridge_delta"]))


class WorkbookReadError(ValueError):
    """The saved workbook does not contain readable suite evidence."""


SECRET_KEYS = {"crumb", "crumbvalue", "cookie", "cookies", "cookievalue", "cookieheader", "setcookie", "sessioncookie", "sessioncookies", "apikey"}


def _redact_api_keys(text):
    """Redact exact key assignments without rewriting surrounding encoded text."""
    decoded, starts = [], []
    pos = 0
    while pos < len(text):
        starts.append(pos)
        if (text[pos] == '%' and pos + 2 < len(text)
                and all(c in '0123456789abcdefABCDEF' for c in text[pos + 1:pos + 3])
                and int(text[pos + 1:pos + 3], 16) < 128):
            decoded.append(chr(int(text[pos + 1:pos + 3], 16)))
            pos += 3
        else:
            decoded.append(text[pos])
            pos += 1
    starts.append(len(text))
    view = ''.join(decoded)
    lowered = view.lower()
    boundaries = '?&;#{}[,(' + ' \t\r\n' + "\"'"
    delimiters = '&;#' + ' \t\r\n' + "\"'<>},])"
    pieces, raw_at, pos = [], 0, 0
    while pos < len(view):
        if pos and view[pos - 1] not in boundaries:
            pos += 1
            continue
        name = next((key for key in ('api_key', 'apikey', 'api-key')
                     if lowered.startswith(key, pos)), None)
        if name is None:
            pos += 1
            continue
        end = pos + len(name)
        if end < len(view) and (view[end].isascii() and (view[end].isalnum() or view[end] in '_-')):
            pos += 1
            continue
        operator = end
        if operator < len(view) and view[operator] in "\"'":
            operator += 1
        while operator < len(view) and view[operator] in ' \t\r\n':
            operator += 1
        if operator >= len(view) or view[operator] not in ':=':
            pos = end
            continue
        encoded = starts[operator + 1] - starts[operator] > 1
        if pos and starts[pos] - starts[pos - 1] > 1:
            encoded = True
        value_at = operator + 1
        while value_at < len(view) and view[value_at] in ' \t\r\n':
            value_at += 1
        quote = ''
        if value_at < len(view) and view[value_at] in "\"'":
            quote = view[value_at]
            encoded = encoded or starts[value_at + 1] - starts[value_at] > 1
            value_at += 1
        end = value_at
        while end < len(view):
            raw_delimiter = starts[end + 1] - starts[end] == 1 or encoded
            if quote:
                if view[end] == quote and raw_delimiter:
                    back = end - 1
                    while back >= value_at and view[back] == '\\':
                        back -= 1
                    if (end - back - 1) % 2 == 0:
                        break
            elif view[end] in delimiters and raw_delimiter:
                break
            end += 1
        if end > value_at:
            pieces.extend((text[raw_at:starts[value_at]], '[redacted]'))
            raw_at = starts[end]
        pos = max(end, operator + 1)
    pieces.append(text[raw_at:])
    return ''.join(pieces)


def redact_anonymous(value):
    """Remove Yahoo session values and API keys from shared nested evidence."""
    if isinstance(value, dict):
        result = {}
        crumb_response = "getcrumb" in str(value.get("http_url", value.get("requested_url", ""))).lower()
        for key, item in value.items():
            normalized = re.sub(r"[^a-z]", "", str(key).lower())
            if normalized in SECRET_KEYS and item not in (None, "") and not isinstance(item, bool):
                result[key] = "[redacted]"
            elif crumb_response and key in {"actual", "body", "response_body"}:
                result[key] = "[redacted crumb response]"
            else:
                result[key] = redact_anonymous(item)
        return result
    if isinstance(value, (list, tuple)):
        return [redact_anonymous(item) for item in value]
    if not isinstance(value, str):
        return value
    value = re.sub(r"(?im)(\b(?:set-cookie|cookie)\s*:\s*)[^\r\n]*", r"\1[redacted]", value)
    value = _redact_api_keys(value)

    value = re.sub(r'''(?i)(["']?crumb["']?\s*[:=]\s*["']?)[^"'\s;&<>]+''', r"\1[redacted]", value)
    value = re.sub(r'''(?i)\b((?:A1S?|A3|GUC|GUCS)=)[^\s;"'<>]+''', r"\1[redacted]", value)
    return value


def write_private_json(path, value):
    """Raw native rows remain private; shared reports use redacted copies."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
        os.fchmod(stream.fileno(), 0o600)
        json.dump(value, stream, indent=2)
        stream.write("\n")


def local_name(name):
    return name.rsplit("}", 1)[-1]


def child(element, name):
    return next((part for part in element if local_name(part.tag) == name), None)


def rich_text(element):
    """Read display text, excluding phonetic guides and formatting text."""
    if element is None:
        return ""
    parts = []
    for item in element:
        tag = local_name(item.tag)
        if tag == "t":
            parts.append(item.text or "")
        elif tag == "r":
            parts.extend(t.text or "" for t in item if local_name(t.tag) == "t")
    return "".join(parts)


def cell_value(cell, strings):
    kind = cell.get("t", "n")
    value = child(cell, "v")
    text = "" if value is None or value.text is None else value.text
    if kind == "inlineStr":
        return rich_text(child(cell, "is"))
    if kind == "s":
        try:
            index = int(text)
            if index < 0:
                raise IndexError
            return strings[index]
        except (ValueError, IndexError) as error:
            raise WorkbookReadError(f"Invalid shared-string index at {cell.get('r')}: {text!r}") from error
    return text


def read_sheets(archive):
    strings = []
    if "xl/sharedStrings.xml" in archive.namelist():
        strings = [rich_text(item) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))
                   if local_name(item.tag) == "si"]
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = {item.get("Id"): item for item in
                     ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
    sheets, types = {}, {}
    sheet_list = child(workbook, "sheets")
    if sheet_list is None:
        raise WorkbookReadError("The workbook has no sheet list")
    for sheet in sheet_list:
        relation_id = next((v for k, v in sheet.attrib.items() if local_name(k) == "id"), None)
        relation = relationships.get(relation_id)
        if relation is None or relation.get("TargetMode") == "External":
            raise WorkbookReadError(f"Missing/internal worksheet relationship for {sheet.get('name')}")
        target = relation.get("Target", "")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        if part.startswith("../") or part not in archive.namelist():
            raise WorkbookReadError(f"Worksheet part does not exist: {part}")
        cells, cell_types = {}, {}
        for element in ET.fromstring(archive.read(part)).iter():
            if local_name(element.tag) != "c":
                continue
            address = element.get("r", "")
            if not re.fullmatch(r"[A-Z]+[1-9][0-9]*", address):
                raise WorkbookReadError(f"Invalid cell address in {part}: {address!r}")
            cells[address] = cell_value(element, strings)
            cell_types[address] = element.get("t", "n")
        sheets[sheet.get("name")] = cells
        types[sheet.get("name")] = cell_types
    return sheets, types


def key_values(cells):
    result = {}
    for address, value in cells.items():
        if re.fullmatch(r"A[1-9][0-9]*", address) and value:
            result[value] = cells.get("B" + address[1:], "")
    return result


def read(path, suite=None):
    """Return raw rows and file fingerprints; never interpret provider availability."""
    path = Path(path).resolve()
    payload = path.read_bytes()
    try:
        with zipfile.ZipFile(BytesIO(payload)) as archive:
            sheets, types = read_sheets(archive)
            vba = archive.read("xl/vbaProject.bin") if "xl/vbaProject.bin" in archive.namelist() else None
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as error:
        raise WorkbookReadError(f"Cannot read {path}: {error}") from error
    if "FunctionResults" not in sheets:
        raise WorkbookReadError("The workbook has no FunctionResults sheet")
    cells = sheets["FunctionResults"]
    rows = sorted({int(address[1:]) for address in cells if re.fullmatch(r"A[1-9][0-9]*", address)
                   and int(address[1:]) > 1 and cells[address]})
    records = []
    for row in rows:
        item = {name: cells.get(column + str(row), "") for column, name in FIELDS.items()}
        item["row"] = row
        item["raw_status"] = item["status"]
        item["xml_cell_types"] = {name: types["FunctionResults"].get(column + str(row), "")
                                  for column, name in FIELDS.items()}
        records.append(item)
    ids = Counter(item["id"] for item in records)
    diagnostics = [f"Duplicate saved case ID: {case_id}" for case_id, count in ids.items() if count > 1]
    diagnostics += [f"Unknown saved status for {item['id']}: {item['status']!r}" for item in records
                    if item["status"] not in RAW_STATUSES]
    control = sheets.get("TestControl", {})
    return {"schema_version": 2, "suite": str(Path(suite).resolve()) if suite else None,
            "workbook": str(path), "workbook_bytes": len(payload),
            "workbook_sha256": hashlib.sha256(payload).hexdigest(),
            "workbook_vba_sha256": hashlib.sha256(vba).hexdigest() if vba is not None else None,
            "control": control, "run_metadata": key_values(control),
            "journal": sheets.get("TestJournal", {}), "sheet_names": list(sheets),
            "counts": dict(Counter(item["status"] for item in records)),
            "diagnostics": diagnostics, "results": records}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--suite", type=Path, required=True, help="Explicit generated suite directory")
    parser.add_argument("--workbook", type=Path, required=True, help="Explicit saved native XLSM")
    parser.add_argument("--output", type=Path, help="Raw private JSON; default: SUITE/native-results.private.json")
    args = parser.parse_args(argv)
    try:
        suite = args.suite.resolve(strict=True)
        if not suite.is_dir():
            raise WorkbookReadError("--suite is not a directory")
        result = read(args.workbook.resolve(strict=True), suite)
        output = args.output or suite / "native-results.private.json"
        write_private_json(output, result)
    except (OSError, WorkbookReadError) as error:
        parser.exit(2, f"ERROR: {error}\n")
    print(json.dumps(redact_anonymous({"output": str(output.resolve()), "counts": result["counts"],
                      "run_metadata": result["run_metadata"], "diagnostics": result["diagnostics"]}), indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
