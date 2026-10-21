"""OrphLux Data Catalog live endpoint.

GET /api/catalog -> {"v":1,"total":N,"counts":{...},"rows":[...]}
Downloads the public Dataset List xlsx from Google Drive, parses it with the
same rules as the original build script, and lets Vercel's CDN cache the
response for an hour (stale-while-revalidate for a day). If the download or
parse fails, or the result looks wrong, it returns 503 and the page falls back
to its bundled snapshot.
"""
from http.server import BaseHTTPRequestHandler
from collections import Counter
import io, json, re, urllib.request
import openpyxl

FILE_ID = '1cHy8M1_U9icoOho4fCpz79fWqDS_madQ'
SOURCES = [
    f'https://drive.usercontent.google.com/download?id={FILE_ID}&export=download&confirm=t',
    f'https://drive.google.com/uc?export=download&id={FILE_ID}',
]
COLLECTIONS = ['projects', 'hq', 'text', 'image', 'video', 'audio']
MIN_TOTAL = 200


def fetch_xlsx():
    last = None
    for url in SOURCES:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (orphlux-data-catalog)'})
            with urllib.request.urlopen(req, timeout=20) as r:
                data = r.read()
            if data[:2] == b'PK':
                return data
            last = f'not an xlsx from {url.split("?")[0]} (check sharing / download settings)'
        except Exception as e:
            last = f'{type(e).__name__}: {e}'
    raise RuntimeError(last or 'download failed')


def build_rows(wb):
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
        raise ValueError(f'Sheet not found: {name_part}. Sheets: {wb.sheetnames}')

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
    return rows, items


def catalog():
    wb = openpyxl.load_workbook(io.BytesIO(fetch_xlsx()))
    rows, items = build_rows(wb)
    counts = dict(Counter(i['collection'] for i in items))
    if len(rows) < MIN_TOTAL or any(counts.get(c, 0) == 0 for c in COLLECTIONS):
        raise RuntimeError(f'sanity check failed: total={len(rows)} counts={counts}')
    return {'v': 1, 'total': len(rows), 'counts': counts, 'rows': rows}


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            body, status = catalog(), 200
            cache = 'public, max-age=0, s-maxage=3600, stale-while-revalidate=86400'
        except Exception as e:
            body, status = {'error': str(e)[:300]}, 503
            cache = 'no-store'
        out = json.dumps(body, ensure_ascii=False, separators=(',', ':')).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', cache)
        self.send_header('Content-Length', str(len(out)))
        self.end_headers()
        self.wfile.write(out)
