#!/usr/bin/env python3
"""Check the Mac table-recipe dispatch source and inventory, without running VBA.

The Python recognizers below check their own fixtures and classify source text.
They do not establish VBA compilation, numeric coercion, extraction, or runtime
behavior. SMFMacTableRecipeSelfTest and SMFMacTableGuardSelfTest cover those
operations when the native Excel harness runs them.
"""

from __future__ import annotations

from pathlib import Path
import re

from port_static_checks import active_lines, logical_lines, DECLARATION, without_strings


ROOT = Path(__file__).resolve().parents[2]
NUMBER = re.compile(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?\Z")
STRING = re.compile(r'"(?:""|[^"])*"\Z')


def split_arguments(text: str) -> list[str] | None:
    """Split text after an opening parenthesis; require a complete outer call."""
    result: list[str] = []
    start = 0
    index = 0
    depth = 0
    quoted = False
    while index < len(text):
        character = text[index]
        if character == '"':
            if quoted and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
        elif not quoted:
            if character == "(":
                depth += 1
            elif character == ")":
                if depth:
                    depth -= 1
                else:
                    if text[index + 1:].strip():
                        return None
                    result.append(text[start:index].strip())
                    return result
            elif character == "," and not depth:
                result.append(text[start:index].strip())
                start = index + 1
        index += 1
    return None


def literal_call(expression: str, name: str, maximum: int) -> list[str] | None:
    expression = expression.strip()
    if expression.startswith("="):
        expression = expression[1:].strip()
    opening = expression.find("(")
    if opening < 0 or expression[:opening].strip().casefold() != name.casefold():
        return None
    arguments = split_arguments(expression[opening + 1:])
    if arguments is None or not 2 <= len(arguments) <= maximum:
        return None
    if not arguments[0] or not arguments[1]:
        return None
    for argument in arguments:
        if (argument and not STRING.fullmatch(argument) and not NUMBER.fullmatch(argument)
                and argument.casefold() not in {"true", "false"}):
            return None
    return arguments


def contains_table_call(expression: str) -> bool:
    masked = re.sub(r'"(?:""|[^"])*"', lambda match: " " * len(match[0]), expression)
    return bool(re.search(r"(?<![A-Za-z0-9_])RCHGetTableCell[ \t\r\n]*\(", masked, re.I))


def source_blocks(filename: str, constants: dict[str, int]) -> dict[str, list[str]]:
    source = (ROOT / "code" / filename).read_text()
    result: dict[str, list[str]] = {}
    current: list[str] | None = None
    for _, line in logical_lines(active_lines(source, constants)):
        declaration = DECLARATION.match(line)
        if declaration:
            current = [line]
            result[declaration[3].casefold()] = current
        elif current is not None:
            current.append(line)
            if re.fullmatch(r"End (?:Function|Sub|Property)", line, re.I):
                current = None
    return result


def recognizer_fixtures() -> int:
    accepted = [
        '=RCHGetTableCell("url",1)',
        ' rChGeTtAbLeCeLl ("url",-1,">Label") ',
        '=RCHGetTableCell("url",+1,,,,,,,,)',
        '=RCHGetTableCell("url",1E999,">Label",,,,,,,"overflow")',
        '=RCHGetTableCell("url",1,"comma, closing) and ""quoted""")',
        '=RCHGetTableCell("url",1,">Missing",,,,,,,False)',
    ]
    rejected = [
        '=RCHGetTableCell("url",1)*1000',
        '=PROPER(RCHGetTableCell("url",1))',
        '=RCHGetTableCell(LEFT("url",1),1)',
        '=RCHGetTableCell("url",1+0)',
        '=RCHGetTableCell("url",A1)',
        '=RCHGetTableCell("url",1',
        '=RCHGetTableCell("url",1,"unterminated)',
        '=RCHGetTableCell("url",1,,,,,,,,,,)',
        '=RCHGetTableCell(,1)',
        '=RCHGetTableCell("url",)',
        '=OtherFunction("url",1)',
        '==RCHGetTableCell("url",1)',
    ]
    for expression in accepted:
        assert literal_call(expression, "RCHGetTableCell", 11) is not None, expression
    for expression in rejected:
        assert literal_call(expression, "RCHGetTableCell", 11) is None, expression
    positive_guards = [
        '=RCHGetTableCell("url",1)',
        '=PROPER(RCHGetTableCell("url",1))',
        '=rchgettablecell ("url",1)*1000',
        '=modGetTableCell.RCHGetTableCell("url",1)',
        '=SUM(1,\tRCHGetTableCell\t("url",1))',
    ]
    negative_guards = [
        '="RCHGetTableCell(""url"",1)"',
        '=OtherRCHGetTableCell("url",1)',
        '=RCHGetTableCellExtra("url",1)',
        '=RCHGetTableCell',
        '=RCHGetTableCell + 1',
        '=SUM(1,2)',
        '="first"" RCHGetTableCell(""url"",1)"',
        '=123RCHGetTableCell("url",1)',
    ]
    for expression in positive_guards:
        assert contains_table_call(expression), expression
    for expression in negative_guards:
        assert not contains_table_call(expression), expression
    return len(accepted) + len(rejected) + len(positive_guards) + len(negative_guards)


def check_source(mac_value: int) -> None:
    constants = {"mac": mac_value, "vba7": -1, "win64": 0}
    mac = source_blocks("modMacCode.bas", constants)
    table = "\n".join(mac["smfmactrytablecellrecipe"])
    assert "Public Function smfMacTryTableCellRecipe(ByVal pExpression As String, ByRef pValue As Variant) As Boolean" in table
    assert "tokens(1 To 11) As String, values(1 To 11) As Variant" in table
    assert "smfMacRecipeArguments(Mid$(expression, openAt + 1), tokens, count)" in table
    assert "count < 2 Or Len(tokens(1)) = 0 Or Len(tokens(2)) = 0" in table
    defaults = {3: '"<BODY"', 4: '" "', 5: '" "', 6: '" "', 7: '0',
                8: '"</BODY"', 9: '0', 10: '"Error"', 11: '0'}
    for index, value in defaults.items():
        assert f"values({index}) = {value}" in table, (index, value)
    handled = table.index("smfMacTryTableCellRecipe = True")
    assert table.index("If Not smfMacRecipeIsNumber(tokens(i)) Then") < handled
    assert table.index("smfMacRecipeLiteral(tokens(i), values(i))") < handled
    assert table.index("Val(") > handled
    error_conversion = table.index("values(10) = Val(tokens(10))")
    assert handled < error_conversion < table.index("values(i) = Val(tokens(i))")
    for variable, index in (("cells", 2), ("rows", 7), ("look", 9), ("requestType", 11)):
        assert table.index(f"{variable} = CInt(values({index}))") > error_conversion
    direct = ("pValue = RCHGetTableCell(CStr(values(1)), cells, CStr(values(3)), "
              "CStr(values(4)), CStr(values(5)), CStr(values(6)), rows, "
              "CStr(values(8)), look, values(10), requestType)")
    assert re.sub(r"\s+", "", direct) in re.sub(r"\s+", "", table)
    assert "If smfMacTryTableCellRecipe Then pValue = values(10)" in table
    assert len(re.findall(r"\bpValue\s*=", table, re.I)) == 2
    assert not re.search(r"\bEvaluate\s*\(|\bApplication\.Run\b", without_strings(table), re.I)

    extractor = source_blocks("modGetElementNumber.bas", constants)
    rch = "\n".join(extractor["rchgetelementnumber"])
    first, second = rch.split('Case aParm(0) = "Evaluate"', 1)[1].split('Case Left(aParm(2), 1) = "="', 1)
    second = second.split('Case Else: RCHGetElementNumber = ""', 1)[0]
    for branch in (first, second):
        tag_at = branch.index("If Not smfMacTryTagContentRecipe(")
        table_at = branch.index("If Not smfMacTryTableCellRecipe(")
        guard_at = branch.index("If MacRecipeCallsTableCell(")
        evaluate_at = branch.index("vTemp = Evaluate(")
        assert tag_at < table_at < guard_at < evaluate_at
        guard_end = branch.index("End If", guard_at)
        guarded = branch[guard_at:guard_end]
        assert "RCHGetElementNumber = pError" in guarded
        assert "Exit Function" in guarded
    guard = "\n".join(extractor["macrecipecallstablecell"])
    assert guard.startswith("Private Function ")
    for fragment in ('Chr$(34)', 'quoted = Not quoted', '[A-Za-z0-9_]',
                     'StrComp(token, "RCHGetTableCell", vbTextCompare)', 'vbTab', 'vbCr', 'vbLf'):
        assert fragment in guard, fragment
    assert not re.search(r"\bEvaluate\s*\(|\bApplication\.Run\b", without_strings(guard), re.I)

    native_tests = source_blocks("modMacIntegrationTests.bas", constants)
    table_test = "\n".join(native_tests["smfmactablerecipeselftest"])
    assert sum(line.startswith("smfMacAssertTableRecipe ") for line in native_tests["smfmactablerecipeselftest"]) == 30
    rejected = table_test.split("rejected = Array(", 1)[1].split("For Each item In rejected", 1)[0].strip()
    assert len(split_arguments(rejected) or []) == 10
    assert "If checks <> 40 Then" in table_test
    for restoration in ("aData(i, 1) = oldKeys(i)", "aData(i, 2) = oldBodies(i)",
                        "sData(i) = oldTexts(i)", "sWebCache = oldCacheMode", "sLog = oldLog", "vError = oldError"):
        assert restoration in table_test
    guard_test = "\n".join(extractor["smfmactableguardselftest"])
    assert len(re.findall(r"actual = RCHGetElementNumber\(", guard_test)) == 4
    arrays = re.findall(r"For Each expression In Array\((.*?)\)\nIf ", guard_test, re.S)
    assert len(arrays) == 2
    assert [len(split_arguments(array + ")") or []) for array in arrays] == [5, 8]
    for restoration in ("aParms(19997) = oldParm", "vError = oldGuardError", "iInit = oldInit"):
        assert restoration in guard_test
    assert "Guard fixtures made an HTTP request" in guard_test

    windows = source_blocks("modGetElementNumber.bas", {"mac": 0, "vba7": -1, "win64": -1})
    win_rch = "\n".join(windows["rchgetelementnumber"])
    assert win_rch.count("smfEvaluateTwice(") == 2
    assert "smfMacTryTableCellRecipe(" not in win_rch
    assert "MacRecipeCallsTableCell(" not in win_rch
    assert "macrecipecallstablecell" not in windows


def inventory() -> tuple[int, int, int, int]:
    records: dict[int, str] = {}
    for suffix in range(23):
        path = ROOT / "docs" / "Elements" / f"smf-elements-{suffix}.txt"
        if not path.exists():
            continue
        for line in path.read_text().splitlines():
            fields = line.split(";")
            if len(fields) >= 4 and fields[0].isdigit() and 1 <= int(fields[0]) <= 20000:
                records[int(fields[0])] = fields[3]
    recipes = {key: value for key, value in records.items() if value.startswith("=")}
    table = {key for key, value in recipes.items() if literal_call(value, "RCHGetTableCell", 11) is not None}
    tag = {key for key, value in recipes.items() if literal_call(value, "smfGetTagContent", 12) is not None}
    guarded = {key for key, value in recipes.items() if key not in table and contains_table_call(value)}
    assert len(recipes) == 1815
    assert len(table) == 138 and len(tag) == 291
    assert 2003 in table
    assert len(guarded) == 35
    expected_guarded = {783, 784, 788, 791, 794, *range(800, 807), *range(811, 814),
                        822, 823, 825, *range(827, 837), 851, 853, 859, 865, 871, 2082, 13889}
    assert guarded == expected_guarded
    return len(recipes), len(tag), len(table), len(guarded)


def main() -> None:
    fixtures = recognizer_fixtures()
    for mac_value in (1, -1):
        check_source(mac_value)
    recipes, tag, table, guarded = inventory()
    print(f"PASS: {fixtures} Python recognizer fixtures; Mac=1/-1 source checks; Windows evaluator dispatch retained.")
    print(f"Inventory: {tag} literal Tag recipes + {table} literal Table recipes / {recipes}; {guarded} unsupported Table expressions guarded.")
    print("Native test source: 40 table-recipe checks and 17 table-guard checks, with state restoration.")
    print("Static checks do not compile or execute VBA and do not establish native runtime success.")


if __name__ == "__main__":
    main()
