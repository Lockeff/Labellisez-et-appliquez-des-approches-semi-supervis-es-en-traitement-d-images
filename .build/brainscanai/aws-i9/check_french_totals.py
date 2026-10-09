"""Verify the actual native table amounts, including French decimal commas.

The bundled arithmetic parser only accepts English decimal points and cannot
read these French currency cells. Keep the user's formatting and check the
explicit sums on the real candidate package before finalization instead.
"""
import hashlib
import json
import re
import zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

build = Path('.build/brainscanai/aws-i9')
costs = json.loads((build / 'costs-fixed-fx.json').read_text(encoding='utf-8'))['display']
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
contracts = [
    (11, 3, ['machine_30h', 'disk_month', 'prototype_images', 'prototype_total']),
    (12, 1, ['compute_4m', 'storage_year', 'disk_month', 'requests', 'one_shot']),
    (12, 2, ['compute_4m', 'storage_rolling', 'disk_year', 'requests', 'rolling_year1']),
]
report = []
with zipfile.ZipFile(build / 'candidate.pptx') as package:
    for slide, column, keys in contracts:
        root = ET.fromstring(package.read(f'ppt/slides/slide{slide}.xml'))
        tables = root.findall('.//a:tbl', ns)
        assert len(tables) == 1
        rows = tables[0].findall('a:tr', ns)
        assert len(rows) == len(keys) + 1
        amounts = []
        for row, key in zip(rows[1:], keys):
            cell = row.findall('a:tc', ns)[column]
            text = ' '.join(t.text or '' for t in cell.findall('.//a:t', ns))
            matches = re.findall(r'(?<![\d,])(\d+,\d+)\s*\u20ac', text)
            assert matches, (slide, column, text)
            # Cells may also contain a monthly explanatory amount after the
            # primary cost; the first currency amount is the line's cost.
            assert matches[0] == costs[key], (slide, key, text, costs[key])
            amounts.append(Decimal(matches[0].replace(',', '.')))
        assert sum(amounts[:-1]) == amounts[-1], (slide, column, amounts)
        report.append({'slide': slide, 'table': 1, 'value_column': column,
                       'amounts': [str(a) for a in amounts], 'passed': True})
(build / 'french-table-totals.json').write_text(
    json.dumps({'passed': True, 'candidate_sha256': hashlib.sha256((build / 'candidate.pptx').read_bytes()).hexdigest(), 'contracts_checked': report}, indent=2), encoding='utf-8')
print('All three required sums match the French amounts in the actual PPTX tables.')
