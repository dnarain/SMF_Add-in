#!/usr/bin/env python3
"""Check the HTML entity data against Python's independent HTML reference.

Run smfMacHTMLSelfTest in Excel for the VBA parser behavior tests.
"""

import html.entities
from pathlib import Path
import re
import unittest


SOURCE = Path(__file__).resolve().parents[2] / "code/modMacHTML.bas"


def entity_table():
    source = SOURCE.read_text()
    columns = []
    for name in ("names", "codes"):
        expression = re.search(
            rf'{name} = Split\((.*?)\s*, "\|"\)', source, re.S
        )
        if expression is None:
            raise AssertionError(f"Cannot find the {name} entity column")
        columns.append("".join(re.findall(r'"([^"]*)"', expression[1])).split("|"))
    if len(columns[0]) != len(columns[1]):
        raise AssertionError("The entity names and code points have different lengths")
    return list(zip(columns[0], map(int, columns[1])))


class HTMLReferenceTests(unittest.TestCase):
    def test_entity_names_are_unique(self):
        table = entity_table()
        self.assertEqual(len(table), len({name for name, _ in table}))

    def test_all_html4_entities_are_present(self):
        names = {name for name, _ in entity_table()}
        self.assertFalse(set(html.entities.name2codepoint) - names)

    def test_named_entities_match_html_reference(self):
        for name, codepoint in entity_table():
            with self.subTest(entity=name):
                # The helper documents HTML4 names plus extra HTML5 names.
                # HTML5 changed the values of HTML4's lang/rang entities.
                if name in html.entities.name2codepoint:
                    expected = chr(html.entities.name2codepoint[name])
                else:
                    expected = html.entities.html5[name + ";"]
                self.assertEqual(expected, chr(codepoint))


if __name__ == "__main__":
    unittest.main(verbosity=2)
