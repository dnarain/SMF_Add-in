#!/usr/bin/env python3
"""Check selected SMF source properties. This does not compile or execute VBA."""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
import json
from pathlib import Path
import re
import subprocess
import sys


def without_comment(line: str) -> str:
    """Remove an apostrophe comment without changing quoted VBA text."""
    quoted = False
    index = 0
    while index < len(line):
        char = line[index]
        if char == '"':
            if quoted and index + 1 < len(line) and line[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
        elif char == "'" and not quoted:
            return line[:index]
        index += 1
    if re.match(r"^\s*Rem(?:\s|$)", line, re.I):
        return ""
    return line


def without_strings(line: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', line)


class Expression:
    """Evaluate the Boolean subset used by this repository's VBA directives.

    VBA True is -1. Not, And, Or, Xor, Eqv, and Imp operate on bits.
    Unknown conditional constants have the VBA default value, zero.
    Unsupported tokens fail the check instead of choosing a branch silently.
    A platform constant can contain 1 even though the Boolean True value is -1.
    """

    TOKEN = re.compile(r"\s*(<>|<=|>=|[()=<>+-]|\d+|[A-Za-z_]\w*)")
    PRECEDENCE = {"imp": 1, "eqv": 2, "xor": 3, "or": 4, "and": 5,
                  "=": 7, "<>": 7, "<": 7, ">": 7, "<=": 7, ">=": 7}

    def __init__(self, text: str, constants: dict[str, int]):
        self.tokens: list[str] = []
        self.index = 0
        self.constants = constants
        position = 0
        text = text.strip()
        while position < len(text):
            match = self.TOKEN.match(text, position)
            if not match:
                raise ValueError(f"Unsupported conditional expression: {text!r}")
            self.tokens.append(match.group(1).lower())
            position = match.end()

    def evaluate(self) -> int:
        if not self.tokens:
            raise ValueError("Empty conditional expression")
        result = self.parse(0)
        if self.index != len(self.tokens):
            raise ValueError(f"Unexpected conditional token: {self.tokens[self.index]}")
        return result

    def parse(self, minimum: int) -> int:
        if self.index >= len(self.tokens):
            raise ValueError("Incomplete conditional expression")
        token = self.tokens[self.index]
        self.index += 1
        if token == "(":
            left = self.parse(0)
            if self.index >= len(self.tokens) or self.tokens[self.index] != ")":
                raise ValueError("Missing ')' in conditional expression")
            self.index += 1
        elif token == "not":
            left = ~self.parse(6)
        elif token in {"+", "-"}:
            value = self.parse(8)
            left = value if token == "+" else -value
        elif token.isdigit():
            left = int(token)
        elif re.fullmatch(r"[a-z_]\w*", token):
            left = self.constants.get(token, 0)
        else:
            raise ValueError(f"Unexpected conditional token: {token}")

        while self.index < len(self.tokens):
            operator = self.tokens[self.index]
            precedence = self.PRECEDENCE.get(operator, -1)
            if precedence < minimum:
                break
            self.index += 1
            right = self.parse(precedence + 1)
            if operator == "and":
                left &= right
            elif operator == "or":
                left |= right
            elif operator == "xor":
                left ^= right
            elif operator == "eqv":
                left = ~(left ^ right)
            elif operator == "imp":
                left = ~left | right
            else:
                relation = {"=": left == right, "<>": left != right,
                            "<": left < right, ">": left > right,
                            "<=": left <= right, ">=": left >= right}[operator]
                left = -1 if relation else 0
        return left


@dataclass
class Branch:
    parent: bool
    taken: bool
    active: bool
    had_else: bool = False


def active_lines(source: str, predefined: dict[str, int]) -> list[tuple[int, str]]:
    constants = {"true": -1, "false": 0, **predefined}
    stack: list[Branch] = []
    result: list[tuple[int, str]] = []
    active = True
    for number, original in enumerate(source.splitlines(), 1):
        line = without_comment(original).strip()
        if not line.startswith("#"):
            if active:
                result.append((number, original))
            continue
        match = re.fullmatch(r"#\s*Const\s+(\w+)\s*=\s*(.+)", line, re.I)
        if match:
            if active:
                constants[match[1].lower()] = Expression(match[2], constants).evaluate()
            continue
        match = re.fullmatch(r"#\s*If\s+(.+)\s+Then", line, re.I)
        if match:
            chosen = bool(Expression(match[1], constants).evaluate())
            stack.append(Branch(active, chosen, active and chosen))
        else:
            match = re.fullmatch(r"#\s*ElseIf\s+(.+)\s+Then", line, re.I)
            if match:
                if not stack or stack[-1].had_else:
                    raise ValueError(f"Line {number}: unmatched #ElseIf")
                chosen = bool(Expression(match[1], constants).evaluate())
                branch = stack[-1]
                branch.active = branch.parent and not branch.taken and chosen
                branch.taken |= chosen
            elif re.fullmatch(r"#\s*Else", line, re.I):
                if not stack or stack[-1].had_else:
                    raise ValueError(f"Line {number}: unmatched or duplicate #Else")
                branch = stack[-1]
                branch.active = branch.parent and not branch.taken
                branch.taken = True
                branch.had_else = True
            elif re.fullmatch(r"#\s*End\s*If", line, re.I):
                if not stack:
                    raise ValueError(f"Line {number}: unmatched #End If")
                stack.pop()
            else:
                raise ValueError(f"Line {number}: unsupported directive {line!r}")
        active = stack[-1].active if stack else True
    if stack:
        raise ValueError("Unclosed #If block")
    return result


def logical_lines(lines: list[tuple[int, str]]):
    pending = ""
    first = 0
    for number, source in lines:
        line = without_comment(source).strip()
        if not pending:
            first = number
        if re.search(r"\s_\s*$", line):
            pending += re.sub(r"\s_\s*$", " ", line)
            continue
        yield first, pending + line
        pending = ""
    if pending:
        raise ValueError(f"Line {first}: unfinished line continuation")


DECLARATION = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b(.*)$", re.I)


def normalized_signature(line: str) -> str:
    parts = re.split(r'("(?:[^"]|"")*")', line)
    signature = "".join(part if index % 2 else re.sub(r"\s+", "", part).lower()
                        for index, part in enumerate(parts))
    return re.sub(r"^public", "", signature)


def public_signatures(lines: list[tuple[int, str]]) -> dict[str, tuple[int, str]]:
    result = {}
    for number, line in logical_lines(lines):
        match = DECLARATION.match(line)
        if not match or (match[1] or "public").lower() != "public":
            continue
        signature = normalized_signature(line)
        name = match[3].lower()
        # The native downloader intentionally copies pURL instead of changing its caller.
        if name == "rchgeturldata1mac":
            signature = signature.replace("(byvalpurlasstring", "(purlasstring", 1)
        key = f"{match[2].lower()} {name}"
        if key in result:
            raise ValueError(f"Line {number}: duplicate public declaration {key}")
        result[key] = (number, signature)
    return result


TYPE_PATTERN = re.compile(
    r"\b(?:As\s+(?:New\s+)?|New\s+)(?:"
    r"(?:MSXML\d*|MSHTML|WinHttp|Scripting|MSScriptControl)\s*\.\s*\w+|"
    r"xmlhttp|HTMLDocument|ScriptControl|FileSystemObject|WinHttpRequest)\b", re.I)
COM_PATTERN = re.compile(r'\bCreateObject\s*\(\s*"([^"]+)"', re.I)
WINDOWS_PROGID = re.compile(
    r"^(?:MSScriptControl\.|Scripting\.|MSXML|WinHttp\.|InternetExplorer\.|"
    r"ADODB\.|WScript\.|htmlfile$)", re.I)
WINDOWS_LIBRARY = re.compile(
    r'\bLib\s+"(?:kernel32|user32|advapi32|wininet|ole32|ws2_32|gdi32|shell32|urlmon)(?:\.dll)?"', re.I)


# These identifier categories are reserved by MS-VBAL3.3.5.2. This catches
# invalid procedure parameters before VBE silently omits their declarations.
# https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7df907cb-ab6c-40d3-aa81-272742ce00c3
RESERVED_PARAMETER_NAMES = frozenset("""
Abs CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng CStr CVar
CVErr Date Debug DoEvents Fix Int Len LenB Me PSet Scale Sgn String Array
Circle Input InputB LBound UBound Boolean Byte Currency Double Integer Long
LongLong LongPtr Single Variant
""".lower().split())
PARAMETER_NAME = re.compile(
    r"[,(]\s*(?:Optional\s+)?(?:(?:ByVal|ByRef)\s+)?([A-Za-z_]\w*)"
    r"\s*(?:\(\s*\))?\s+As\b", re.I)


def reserved_parameters(lines: list[tuple[int, str]]) -> list[tuple[int, str]]:
    problems = []
    for number, line in logical_lines(lines):
        if not DECLARATION.match(line):
            continue
        for parameter in PARAMETER_NAME.finditer(without_strings(line)):
            if parameter[1].lower() in RESERVED_PARAMETER_NAMES:
                problems.append((number, f"Reserved VBA parameter name {parameter[1]}"))
    return problems


def forbidden_dependencies(lines: list[tuple[int, str]]) -> list[tuple[int, str]]:
    problems = []
    for number, line in logical_lines(lines):
        identifiers = without_strings(line)
        if TYPE_PATTERN.search(identifiers):
            problems.append((number, "Windows-only early-bound type"))
        if re.search(r"\bMacScript\s*\(", identifiers, re.I):
            problems.append((number, "deprecated MacScript call"))
        if WINDOWS_LIBRARY.search(line):
            problems.append((number, "Windows system library"))
        for match in COM_PATTERN.finditer(line):
            if WINDOWS_PROGID.match(match[1]):
                problems.append((number, f"Windows COM object {match[1]}"))
    return problems


def self_tests() -> int:
    definitions = {"mac": -1, "vba7": -1, "win64": 0}
    cases = [
        ("#If Mac Then\nmac\n#Else\nwindows\n#End If", ["mac"]),
        ("#If Not Mac Then\nwindows\n#ElseIf VBA7 Then\nmac7\n#Else\nother\n#EndIf", ["mac7"]),
        ("#If Mac Then\n#If Win64 Then\nwin\n#Else\nmac\n#End If\n#End If", ["mac"]),
        ("#Const UseCOM = True\n#If Mac Or Not UseCOM Then\nportable\n#End If", ["portable"]),
        ("#If UndefinedConstant Then\nbad\n#Else\ngood\n#End If", ["good"]),
        ("#If nOt (MAC = FALSE) And VBA7 Then ' explanation\nyes\n#End If", ["yes"]),
        ("#If False Then\n#Const Skip = True\n#Else\n#If Not Skip Then\ngood\n#End If\n#End If", ["good"]),
    ]
    for source, expected in cases:
        actual = [text for _, text in active_lines(source, definitions)]
        if actual != expected:
            raise AssertionError(f"Conditional fixture failed: {actual!r} != {expected!r}")
    for source in ("#If Mac Then", "#Else", "#End If", "#If Mac Then\n#Else\n#Else\n#End If"):
        try:
            active_lines(source, definitions)
        except ValueError:
            pass
        else:
            raise AssertionError(f"Invalid directive fixture accepted: {source}")
    code = [(1, 'x = "MacScript(ignored)" \' New HTMLDocument'),
            (2, 'Dim x As New HTMLDocument'),
            (3, 'Set y = CreateObject("WinHttp.WinHttpRequest.5.1")')]
    if [line for line, _ in forbidden_dependencies(code)] != [2, 3]:
        raise AssertionError("Dependency scanner string/comment fixture failed")
    if normalized_signature('Public Function F(Optional x As String = "A B")') != \
            'functionf(optionalxasstring="A B")':
        raise AssertionError("Signature normalization changed a quoted default")
    for mac_value in (1, -1):
        constants = {"mac": mac_value, "vba7": -1, "win64": 0}
        source = "#If Mac = 0 Then\nwindows\n#Else\nnative\n#End If"
        if [text for _, text in active_lines(source, constants)] != ["native"]:
            raise AssertionError(f"Explicit Mac comparison failed with Mac={mac_value}")
        source = "#If Not Mac Then\nwindows\n#Else\nnative\n#End If"
        expected = ["windows"] if mac_value == 1 else ["native"]
        if [text for _, text in active_lines(source, constants)] != expected:
            raise AssertionError(f"Numeric Not Mac regression fixture failed with Mac={mac_value}")
    for source, expected in (("Not 1", -2), ("Not True", 0), ("True", -1)):
        if Expression(source, {"true": -1, "false": 0}).evaluate() != expected:
            raise AssertionError(f"Numeric VBA expression fixture failed: {source}")
    if not unsafe_mac_guards("#If Not Mac Then") or unsafe_mac_guards("#If Mac = 0 Then"):
        raise AssertionError("Unsafe numeric Mac guard fixture failed")
    parameter_fixtures = [
        ("Private Function F(ByVal scale As Single) As String", 1),
        ("Private Function F(Optional ByRef SCALE As Single = 1) As String", 1),
        ("Private Function F(ByVal pArg As Long, _\n    ByVal Scale As Single) As String", 1),
        ("Private Function F(ByVal pictureScale As Single) As String", 0),
        ('Private Function F(Optional pArg As String = "ByVal Scale As Single")', 0),
    ]
    for source, expected in parameter_fixtures:
        actual = reserved_parameters(list(enumerate(source.splitlines(), 1)))
        if len(actual) != expected:
            raise AssertionError(f"Reserved parameter scanner fixture failed: {source!r}: {actual!r}")
    return len(cases) + 14 + len(parameter_fixtures)


def mac_configurations():
    # Cover numeric Mac=1 and Boolean-style Mac=-1. Neither is assumed equivalent
    # under bitwise Not. Win64 is tested in both zero and nonzero states.
    for mac_value in (1, -1):
        for win64 in (0, -1):
            key = f"mac={mac_value};win64={win64}"
            yield key, {"mac": mac_value, "vba7": -1, "win64": win64, "win32": 0}


def unsafe_mac_guards(source: str) -> list[int]:
    return [number for number, line in enumerate(source.splitlines(), 1)
            if re.search(r"^\s*#(?:If|ElseIf)\b.*\bNot\s+Mac\b", without_comment(line), re.I)]


def collect_git_baseline(root: Path, reference: str) -> dict:
    """Read this checkout's Git baseline. Never use a parent repository."""
    if not (root / ".git").exists():
        raise ValueError("The source folder has no .git entry; an explicit Git baseline is unavailable")
    top = subprocess.run(["git", "rev-parse", "--show-toplevel"], cwd=root,
                         capture_output=True, text=True, check=False)
    if top.returncode or Path(top.stdout.strip()).resolve() != root.resolve():
        raise ValueError("The Git baseline must belong to this source folder")
    commit = subprocess.run(["git", "rev-parse", f"{reference}^{{commit}}"], cwd=root,
                            capture_output=True, text=True, check=False)
    if commit.returncode:
        raise ValueError(f"Cannot read baseline commit {reference}")
    tracked = subprocess.run(["git", "ls-tree", "-r", "--name-only", reference, "--", "code"],
                             cwd=root, capture_output=True, text=True, check=True)
    modules = {}
    for relative in tracked.stdout.splitlines():
        path = Path(relative)
        if path.suffix.lower() != ".bas" or path.name == "exportMods.bas":
            continue
        source = subprocess.run(["git", "show", f"{reference}:{relative}"], cwd=root,
                                capture_output=True, check=True).stdout.decode("utf-8", "replace")
        modules[relative] = source
    data = {"schema_version": 2,
            "upstream_repository": "https://github.com/ClimberMel/SMF_Add-in",
            "upstream_commit": commit.stdout.strip(), "configurations": {}}
    for key, constants in mac_configurations():
        data["configurations"][key] = {
            "constants": constants,
            "module_signatures": {
                relative: {name: signature for name, (_, signature) in
                           public_signatures(active_lines(source, constants)).items()}
                for relative, source in modules.items()}}
    return data


def signature_baseline(root: Path, reference: str | None) -> dict:
    if reference is not None:
        return collect_git_baseline(root, reference)
    path = root / "mac" / "tests" / "upstream-signatures.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema_version") != 2 or not re.fullmatch(r"[0-9a-f]{40}", data.get("upstream_commit", "")):
        raise ValueError("Invalid bundled signature baseline metadata")
    for key, expected in mac_configurations():
        configuration = data["configurations"][key]
        if configuration["constants"] != expected or not configuration["module_signatures"]:
            raise ValueError("Invalid bundled signature baseline configuration")
    return data


def check(root: Path, baseline: str | None) -> int:
    failures = []
    files = sorted(path for path in (root / "code").glob("*")
                   if path.suffix.lower() in {".bas", ".cls"} and path.name != "exportMods.bas")
    try:
        original = signature_baseline(root, baseline)
    except (ValueError, OSError, KeyError, subprocess.CalledProcessError) as error:
        print(f"FAIL: {error}")
        return 1
    baseline_modules = next(iter(original["configurations"].values()))["module_signatures"]
    for relative in baseline_modules:
        path = root / relative
        if not path.exists():
            failures.append(f"{relative}: baseline VBA module was removed")
    signatures_checked = 0
    for path in files:
        for number in unsafe_mac_guards(path.read_text(encoding="utf-8", errors="replace")):
            failures.append(f"{path.relative_to(root)}:{number}: Not Mac negates a numeric constant; use Mac = 0")
    for key, constants in mac_configurations():
        before_modules = original["configurations"][key]["module_signatures"]
        public_names: dict[str, list[str]] = defaultdict(list)
        module_names: dict[str, list[str]] = defaultdict(list)
        for path in files:
            relative = path.relative_to(root).as_posix()
            source = path.read_text(encoding="utf-8", errors="replace")
            configuration = f"Mac={constants['mac']} VBA7={constants['vba7']} Win64={constants['win64']}"
            try:
                lines = active_lines(source, constants)
                for number, problem in reserved_parameters(lines):
                    failures.append(f"{relative}:{number} [{configuration}]: {problem}")
                for number, problem in forbidden_dependencies(lines):
                    failures.append(f"{relative}:{number} [{configuration}]: {problem}")
                module = re.search(r'^Attribute VB_Name = "([^"]+)"', source, re.M)
                if module:
                    module_names[module[1].lower()].append(relative)
                if path.suffix.lower() != ".bas":
                    continue
                current = public_signatures(lines)
                for name in current:
                    if name.startswith(("function ", "sub ")):
                        public_names[name.rsplit(" ", 1)[-1]].append(relative)
                for name, signature in before_modules.get(relative, {}).items():
                    signatures_checked += 1
                    if name not in current:
                        failures.append(f"{relative} [{configuration}]: missing public {name}")
                    elif current[name][1] != signature:
                        failures.append(f"{relative}:{current[name][0]} [{configuration}]: changed public signature {name}")
            except (ValueError, AssertionError) as error:
                failures.append(f"{relative} [{configuration}]: {error}")
        for identifiers in (public_names, module_names):
            for name, locations in identifiers.items():
                if len(locations) > 1:
                    failures.append(f"Duplicate identifier {name}: {', '.join(locations)}")
    if failures:
        for failure in failures:
            print(f"FAIL: {failure}")
        return 1
    print(f"PASS: {len(files)} modules; 4 native Mac conditional configurations (Mac=1/-1); "
          f"{signatures_checked} existing public signatures checked.")
    print(f"Signature baseline: upstream {original['upstream_commit']} "
          f"({'local Git' if baseline is not None else 'bundled JSON; no Git required'}).")
    print("No selected Windows dependencies, MacScript calls, or duplicate public procedures found in active Mac code.")
    print("STATIC CHECK ONLY: VBA compilation, Excel calculation, and network access were not executed.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--baseline", help="Use a local Git ref instead of the bundled upstream signature baseline")
    parser.add_argument("--self-test-only", action="store_true")
    args = parser.parse_args()
    count = self_tests()
    print(f"PASS: {count} conditional-parser, dependency-scanner and parameter-scanner fixtures.")
    return 0 if args.self_test_only else check(args.root.resolve(), args.baseline)


if __name__ == "__main__":
    sys.exit(main())
