"""Targeted syntax regression for statement keywords; not a VBA compiler."""
from pathlib import Path
import re
import unittest
ROOT=Path(__file__).resolve().parents[1]
STATEMENT_NAMES={'scale','step','line','width'}

def invalid_declarations(source):
    found=[]
    for line in source.splitlines():
        match=re.match(r'\s*Dim\s+(.+)',line,re.I)
        if not match:continue
        for declaration in match.group(1).split(','):
            token=re.match(r'\s*(\w+)',declaration)
            if token and token.group(1).casefold() in STATEMENT_NAMES:found.append(token.group(1))
    return found

class VbaIdentifierRegression(unittest.TestCase):
    def test_detects_statement_identifiers_in_dim_lists(self):
        self.assertEqual(invalid_declarations('Dim scale As Double, v As Double\nDim old As Double, step As Double\nDim line As Shape, width As Double'),['scale','step','line','width'])
    def test_network_modules_avoid_statement_identifiers(self):
        for p in (ROOT/'vba/network').glob('*.bas'):
            with self.subTest(module=p.name):self.assertEqual(invalid_declarations(p.read_text()),[])
    def test_office_line_and_dimension_members_remain(self):
        src=(ROOT/'vba/network/NetzDiagramm.bas').read_text()
        self.assertIn('verbindung.Line.EndArrowheadStyle',src)
        self.assertIn('node.Left + node.Width',src)
        self.assertIn('node.Top + node.Height',src)
