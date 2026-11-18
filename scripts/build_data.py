"""Parse the EXPENSES pivot (bold rows = category totals, indented rows = breakup)
into data/expenses.json.   Usage: python scripts/build_data.py EXPENSES.xlsx"""
import sys, json, openpyxl, re, pathlib
src = sys.argv[1] if len(sys.argv) > 1 else 'EXPENSES.xlsx'
ws = openpyxl.load_workbook(src, data_only=True).active
hdr = [c.value for c in ws[1]]
months = [str(h).strip().title()[:3] for h in hdr[1:-1]]          # Jan..Nov
n = len(months)
INFLOW = {"owner's injection", "loan", "opening deposit"}           # money coming in
groups, cur, check = [], None, []
for r in ws.iter_rows(min_row=2):
    label = r[0].value
    if label is None: continue
    label = str(label).strip()
    vals = [float(c.value or 0) for c in r[1:1+n]]
    tot  = float(r[1+n].value or 0)
    if label.lower() == 'grand total':
        grand = vals; continue
    if r[0].font.b:                                                  # category row
        cur = {'name': label, 'kind': 'in' if label.lower() in INFLOW else 'out',
               'pivot': vals, 'items': []}
        groups.append(cur)
    else:                                                            # breakup row
        cur['items'].append({'name': label, 'v': [round(x, 2) for x in vals]})
# merge duplicate item names inside a group (e.g. 'Contra' / 'Contra ')
for g in groups:
    m = {}
    for it in g['items']:
        k = it['name']
        m[k] = [a+b for a, b in zip(m.get(k, [0]*n), it['v'])]
    g['items'] = [{'name': k, 'v': [round(x, 2) for x in v]} for k, v in m.items()]
# reconcile: category row == sum of its breakup rows ; all categories == grand total
bad = 0
for g in groups:
    s = [sum(i['v'][j] for i in g['items']) for j in range(n)]
    d = max(abs(a-b) for a, b in zip(s, g['pivot']))
    if d > 0.05: bad += 1; print('MISMATCH', g['name'], round(d, 2))
    del g['pivot']
gt = [sum(i['v'][j] for g in groups for i in g['items']) for j in range(n)]
print('categories', len(groups), 'items', sum(len(g['items']) for g in groups),
      'max diff vs Grand Total', round(max(abs(a-b) for a, b in zip(gt, grand)), 2), 'mismatches', bad)
out = pathlib.Path(__file__).resolve().parent.parent / 'data' / 'expenses.json'
out.write_text(json.dumps({'months': months, 'groups': groups}, separators=(',', ':')))
print('wrote', out)
# also refresh the data embedded in index.html
import re
idx = out.parent.parent / 'index.html'
if idx.exists():
    html = idx.read_text()
    html = re.sub(r'(<script type="application/json" id="exp">).*?(</script>)',
                  lambda m: m.group(1) + out.read_text() + m.group(2), html, flags=re.S)
    idx.write_text(html); print('updated', idx)
