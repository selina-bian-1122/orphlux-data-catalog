"""OrphLux Data Catalog build: xlsx -> gzip JSON rows split into 10 parts.
Usage: python3 -I build_catalog.py <input.xlsx> <outdir>
Writes <outdir>/p0..p9.bin, p0..p9.b64, manifest.json (sha1/size per part + data hash).
"""
import openpyxl, json, re, sys, os, gzip, base64, hashlib
src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
wb = openpyxl.load_workbook(src)
URL = re.compile(r'https?://[^\s|<>"]+')

def filled(ws):
    grid = {(c.row, c.column): c for row in ws.iter_rows() for c in row}
    vals = {k: v.value for k, v in grid.items()}
    for m in ws.merged_cells.ranges:
        top = vals.get((m.min_row, m.min_col))
        for r in range(m.min_row, m.max_row + 1):
            for col in range(m.min_col, m.max_col + 1):
                vals[(r, col)] = top
    return grid, vals

def clean(s):
    if s is None: return ''
    if isinstance(s, float): return f'{int(s):,}' if s.is_integer() else str(s)
    if isinstance(s, int): return f'{s:,}'
    s = str(s).replace('\r', '').replace('　', ' ')
    s = re.sub(r'[ \t]+\n', '\n', s); s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def wnum(s):
    m = re.fullmatch(r'(\d+(?:\.\d+)?)\s*[wW](\+?)', s.strip())
    if not m: return s
    n = float(m.group(1)) * 10000
    return (f'{n/1e6:g}M' if n >= 1e6 else f'{n/1e3:g}K') + m.group(2)

def split_links(text, cell):
    urls = URL.findall(text)
    rest = re.sub(r'\n{3,}', '\n\n', URL.sub('', text)).strip(' \n|')
    if cell is not None and cell.hyperlink and cell.hyperlink.target and cell.hyperlink.target not in urls:
        urls.append(cell.hyperlink.target)
    return rest, [u.rstrip('.,;)') for u in urls]

def link_label(u, default):
    if 'docs.google.com/document' in u: return 'Overview document'
    if 'docs.google.com/spreadsheets' in u: return 'Specification sheet'
    return default

items = []
def add(**kw): items.append(kw)

def sheet(name_part):
    for n in wb.sheetnames:
        if name_part.lower() in n.lower().replace('‑', '-'): return wb[n]
    raise SystemExit(f'Sheet not found: {name_part}. Sheets: {wb.sheetnames}')

# Projects
ws = sheet('Projects'); grid, v = filled(ws)
for r in range(2, ws.max_row + 1):
    cat, name, ov, samp, vol, status = (v.get((r, c)) for c in range(1, 7))
    if not (ov or samp): continue
    ovt, ovlinks = split_links(clean(ov), grid.get((r, 3)))
    st, slinks = split_links(clean(samp), grid.get((r, 4)))
    links = [{'label': link_label(u, 'Project overview'), 'url': u} for u in ovlinks]
    links += [{'label': link_label(u, 'Sample data'), 'url': u, 'sample': True} for u in slinks]
    nm, variant = clean(name), ''
    merged_name = any(m.min_row <= r <= m.max_row and m.min_col == 2 and m.max_row > m.min_row for m in ws.merged_cells.ranges)
    if merged_name and clean(cat) == 'Multimodal LLM Annotation' and items and items[-1]['name'] == nm:
        items[-1]['description'] += '\n\n' + ovt
        items[-1]['links'] += [l for l in links if l['url'] not in [x['url'] for x in items[-1]['links']]]
        continue
    if merged_name and ovt and clean(cat) != 'Multimodal LLM Annotation':
        first, _, restt = ovt.partition('\n')
        if len(first) < 80: variant, ovt = first.strip(), restt.strip()
    vv = clean(vol)
    if re.fullmatch(r'[\d,]+', vv): vv += ' hours'
    add(collection='projects', category=clean(cat), name=nm, variant=variant, description=ovt, strengths='',
        sampleNote=st, volume=vv, format='', status=clean(status), links=links)

# High-quality
ws = sheet('quality Dataset'); grid, v = filled(ws)
for r in range(2, ws.max_row + 1):
    cat, name, ov, strength, samp, vol = (v.get((r, c)) for c in range(1, 7))
    if not (name or samp): continue
    st, slinks = split_links(clean(samp), grid.get((r, 5)))
    ovt, ovl = split_links(clean(ov), grid.get((r, 3)))
    links = [{'label': link_label(u, 'Reference'), 'url': u} for u in ovl]
    links += [{'label': 'Sample data', 'url': u, 'sample': True} for u in slinks]
    if items and items[-1]['collection'] == 'hq' and clean(name) == items[-1]['name'] and not grid[(r, 2)].value:
        items[-1]['links'] += links; continue
    add(collection='hq', category=clean(cat), name=clean(name), variant='', description=ovt,
        strengths=clean(strength), sampleNote=st, volume=clean(vol), format='', status='', links=links)

# Off-the-shelf
for part, col in [('Text Category', 'text'), ('Graphic and Text', 'image'), ('Video Category', 'video'), ('Voice Category', 'audio')]:
    ws = sheet(part); grid, v = filled(ws)
    merged_rows = {r for m in ws.merged_cells.ranges if m.min_col == 1 and m.max_row > m.min_row for r in range(m.min_row, m.max_row + 1)}
    for r in range(2, ws.max_row + 1):
        name, content, vol, fmt, samp = (v.get((r, c)) for c in range(1, 6))
        if not (name or content): continue
        st, slinks = split_links(clean(samp), grid.get((r, 5)))
        links = [{'label': 'Sample data', 'url': u, 'sample': True} for u in slinks]
        nm, desc, variant, group = clean(name), clean(content), '', ''
        if r in merged_rows:
            group = nm; first = desc.split('\n')[0]
            if len(first) <= 90 and ':' not in first[:40]:
                variant = first; desc = desc[len(first):].strip()
            else:
                title = re.split(r'[:：,，.]', first)[0].strip()
                variant = title if len(title) <= 70 else ''
        if group and not variant and len(desc) <= 120: variant, desc = desc, ''
        if desc.startswith('PubMed is'): variant = 'PubMed'
        vv = wnum(clean(vol))
        m = re.fullmatch(r'Question (\d+)', vv)
        if m: vv = f'{int(m.group(1)):,} questions'
        vv = re.sub(r'^(\d+) Questions$', lambda m: f'{int(m.group(1)):,} questions', vv)
        vv = re.sub(r'^(\d+):00 hours$', r'\1 hours', vv)
        add(collection=col, category=group, name=nm, variant=variant, description=desc, strengths='',
            sampleNote=st, volume=vv, format=clean(fmt), status='', links=links)

for it in items:
    s = [l for l in it['links'] if l.get('sample')]
    if len(s) > 1:
        for n, l in enumerate(s, 1): l['label'] = f'Sample data {n}'

cols = ['projects', 'hq', 'text', 'image', 'video', 'audio']
def tidy(t):
    t = re.sub(r'[ \t]+', ' ', t); t = re.sub(r'\n ', '\n', t); return re.sub(r'\n{2,}', '\n\n', t)
rows = []
for i in items:
    links = []
    for l in i['links']:
        u = re.sub(r'\?usp=(drive_link|sharing)$', '', l['url'])
        u = re.sub(r'\?dmr=1&ec=wgc-drive-%5Bmodule%5D-goto$', '', u)
        links.append([u, 1 if l.get('sample') else 0, l['label']])
    rows.append([cols.index(i['collection']), i['category'], i['name'], i['variant'], tidy(i['description']), tidy(i['strengths']),
                 i['volume'], i['format'], i['status'], i['sampleNote'], links])
raw = json.dumps(rows, ensure_ascii=False, separators=(',', ':')).encode()
gz = gzip.compress(raw, 9, mtime=0)
N = 10; size = -(-len(gz) // N)
parts = []
for k in range(N):
    p = gz[k*size:(k+1)*size]
    open(f'{out}/p{k}.bin', 'wb').write(p)
    open(f'{out}/p{k}.b64', 'w').write(base64.b64encode(p).decode())
    parts.append({'file': f'd/p{k}.bin', 'sha': hashlib.sha1(p).hexdigest(), 'size': len(p)})
from collections import Counter
manifest = {'data_hash': hashlib.sha1(raw).hexdigest(), 'counts': dict(Counter(i['collection'] for i in items)), 'total': len(items), 'parts': parts}
json.dump(manifest, open(f'{out}/manifest.json', 'w'), indent=1)
print(json.dumps(manifest, indent=1))
