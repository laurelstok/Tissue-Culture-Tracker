#!/usr/bin/env python3
"""
TC Tracker local server — run with: python3 server.py
Then open: http://localhost:8081
"""
import os, cgi, json, shutil, io, re, subprocess, tempfile
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse
try:
    import urllib.request as _urllib_req
    _HAS_URLLIB = True
except ImportError:
    _HAS_URLLIB = False

# ── Confluency API settings ───────────────────────────────────────────────────
# Set CONFLUENCY_API_URL to the address of your confluency_api.py server
# e.g. http://192.168.1.50:8082
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tc_settings.json')

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE) as f:
                return json.load(f)
        except: pass
    return {'confluency_api_url': ''}

def save_settings(data):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(data, f, indent=2)

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images')
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── helpers ──────────────────────────────────────────────────────────────────

def parse_sci(s):
    if not s: return None
    s = str(s).strip().replace(',', '')
    try: return float(s)
    except: pass
    sup = {'\u2070':0,'\u00b9':1,'\u00b2':2,'\u00b3':3,'\u2074':4,
           '\u2075':5,'\u2076':6,'\u2077':7,'\u2078':8,'\u2079':9,'\u207b':'-'}
    m = re.match(r'^([\d.]+)?\u00d710(.+)$', s)
    if m:
        coeff = float(m.group(1)) if m.group(1) else 1.0
        exp_str = ''.join(str(sup.get(c, c)) for c in m.group(2))
        try: return coeff * (10 ** int(exp_str))
        except: return None
    return None


def get_lineage_project(rec, passages):
    seen, cur = set(), rec
    while cur and cur.get('id') not in seen:
        seen.add(cur.get('id'))
        if cur.get('project'): return cur['project']
        pid = cur.get('parent')
        if not pid: break
        cur = next((p for p in passages if p.get('id') == pid), None)
    return ''


def calc_fold_change(rec, passages):
    total = parse_sci(rec.get('totalViableCells'))
    if not total: return None
    seeded = parse_sci(rec.get('seedingTotal'))
    if not seeded and rec.get('vessels'):
        s = sum((parse_sci(v.get('seedingTotal')) or 0) * (v.get('qty') or 1) for v in rec['vessels'])
        if s: seeded = s
    if not seeded:
        par = next((p for p in passages if p.get('id') == rec.get('parent')), None)
        if par:
            seeded = parse_sci(par.get('seedingTotal'))
            if not seeded and par.get('vessels'):
                s = sum((parse_sci(v.get('seedingTotal')) or 0) * (v.get('qty') or 1) for v in par['vessels'])
                if s: seeded = s
            if not seeded and par.get('plateData'):
                wa = {'6-well plate': 9.5, '12-well plate': 3.8, '24-well plate': 1.9,
                      '48-well plate': 0.95, '96-well plate': 0.32, '384-well plate': 0.056}
                area = wa.get(par.get('wells', ''), 9.5)
                tw = 0
                for plate in par['plateData'].values():
                    for well in plate.values():
                        if well.get('count'): tw += parse_sci(well['count']) or 0
                        elif well.get('seeding') and well.get('occupied'):
                            tw += (parse_sci(well['seeding']) or 0) * area
                if tw: seeded = tw
    return (total / seeded) if seeded and seeded > 0 else None


# ── Excel export ──────────────────────────────────────────────────────────────

def build_excel(data):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    passages = data.get('passages', [])
    projects = data.get('projects', [])
    HEADER_BG = 'FF1A1915'
    ABG = {'thaw': 'FFE6F1FB', 'inherit': 'FFEEEDFE', 'passage': 'FFE1F5EE',
           'freeze': 'FFFAEEDA', 'experiment': 'FFFAECE7'}
    thin = Side(style='thin', color='FFD0D0D0')
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hcell(ws, r, c, v, w=None):
        cell = ws.cell(r, c, v)
        cell.font = Font(bold=True, color='FFFFFFFF', name='Arial', size=10)
        cell.fill = PatternFill('solid', start_color=HEADER_BG)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bdr
        if w: ws.column_dimensions[get_column_letter(c)].width = w
        return cell

    def dcell(ws, r, c, v, bg=None, align='left', fmt=None):
        cell = ws.cell(r, c, v)
        cell.font = Font(name='Arial', size=10)
        if bg: cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
        cell.border = bdr
        if fmt: cell.number_format = fmt
        return cell

    def vstr(rec):
        if rec.get('vessels'):
            return ', '.join('({}){}' .format(v.get('qty', 1), v.get('type', '')) for v in rec['vessels'])
        return rec.get('wells', '')

    ws1 = wb.active
    ws1.title = 'Records'
    ws1.freeze_panes = 'A2'
    ws1.row_dimensions[1].height = 36
    c1 = [('Record ID',14),('Parent ID',14),('Line',14),('Project',16),
          ('Action',10),('P#',6),('Date',12),('Days',6),
          ('Condition',20),('Substrate',14),('Vessel(s)',22),
          ('Split ratio',10),('Confluency %',10),
          ('Seeding density',14),('Total seeded',14),
          ('Viable cells/mL',14),('% Viability',10),('Volume mL',10),
          ('Total viable cells',16),('Fold change',12),
          ('Vials',8),('Cells/vial',14),('Cryoprotectant',16),
          ('Storage',18),('Vial viability %',12),
          ('Exp ID',12),('Assay',16),('Treatment',20),('Timepoints',14),
          ('Operator',14),('Notes',30),('Feeds',8),('Images',8)]
    for ci, (label, w) in enumerate(c1, 1): hcell(ws1, 1, ci, label, w)
    for ri, rec in enumerate(passages, 2):
        bg = ABG.get(rec.get('action', ''))
        proj = get_lineage_project(rec, passages)
        fc = calc_fold_change(rec, passages)
        vals = [rec.get('id',''), rec.get('parent','') or '', rec.get('line',''), proj,
                rec.get('action',''), rec.get('passageNum',''), rec.get('date',''),
                rec.get('days','') or '', rec.get('condition',''), rec.get('substrate',''),
                vstr(rec), rec.get('splitRatio',''), rec.get('confluency','') or '',
                rec.get('seedingDensity',''), rec.get('seedingTotal',''),
                rec.get('viableCellsPerMl',''), rec.get('viabilityPct','') or '',
                rec.get('volumeMl','') or '', rec.get('totalViableCells',''),
                round(fc, 2) if fc else '',
                rec.get('vials','') or '', rec.get('cellsPerVial',''),
                rec.get('cryo',''), rec.get('storage',''), rec.get('viability','') or '',
                rec.get('expId',''), rec.get('assay',''), rec.get('treatment',''),
                rec.get('timepoints',''), rec.get('user',''), rec.get('note',''),
                len(rec.get('feeds', [])), len(rec.get('images', []))]
        for ci, v in enumerate(vals, 1):
            align = 'center' if ci in (6, 8, 12, 13, 20, 32, 33) else 'left'
            fmt = '0.00"x"' if ci == 20 and v else None
            dcell(ws1, ri, ci, v, bg=bg, align=align, fmt=fmt)
    ws1.auto_filter.ref = ws1.dimensions

    ws2 = wb.create_sheet('Feed log')
    ws2.freeze_panes = 'A2'
    c2 = [('Record ID',14),('Line',14),('Project',14),('Action',10),
          ('Feed date',12),('Media type',22),('Volume mL',10),
          ('Operator',14),('Notes',30),('Cat #',14),('Lot #',14),('Expiration',12)]
    for ci, (l, w) in enumerate(c2, 1): hcell(ws2, 1, ci, l, w)
    row = 2
    for rec in passages:
        proj = get_lineage_project(rec, passages)
        bg = ABG.get(rec.get('action', ''))
        for feed in rec.get('feeds', []):
            vals = [rec.get('id',''), rec.get('line',''), proj, rec.get('action',''),
                    feed.get('date',''), feed.get('media',''), feed.get('volumeMl','') or '',
                    feed.get('user',''), feed.get('note',''),
                    feed.get('mediaCat',''), feed.get('mediaLot',''), feed.get('mediaExp','')]
            for ci, v in enumerate(vals, 1): dcell(ws2, row, ci, v, bg=bg)
            row += 1
    if row == 2: ws2.cell(2, 1, 'No feed events recorded yet.')

    ws3 = wb.create_sheet('Plate wells')
    ws3.freeze_panes = 'A2'
    c3 = [('Record ID',14),('Line',14),('Project',14),('Action',10),('Date',12),
          ('Plate key',14),('Well',8),('Seeding density',14),('Cell count',14),
          ('Condition',20),('Treatment',20),('Notes',30),('Contaminated',12)]
    for ci, (l, w) in enumerate(c3, 1): hcell(ws3, 1, ci, l, w)
    row = 2
    for rec in passages:
        proj = get_lineage_project(rec, passages)
        bg = ABG.get(rec.get('action', ''))
        for pk, plate in rec.get('plateData', {}).items():
            for wid, well in plate.items():
                if not well.get('occupied') and not well.get('contaminated'): continue
                vals = [rec.get('id',''), rec.get('line',''), proj, rec.get('action',''),
                        rec.get('date',''), pk, wid,
                        well.get('seeding',''), well.get('count',''),
                        well.get('cond',''), well.get('treat',''), well.get('note',''),
                        'Yes' if well.get('contaminated') else '']
                for ci, v in enumerate(vals, 1): dcell(ws3, row, ci, v, bg=bg)
                row += 1
    if row == 2: ws3.cell(2, 1, 'No plate well data recorded yet.')

    ws4 = wb.create_sheet('Summary')
    ws4.freeze_panes = 'A2'
    c4 = [('Cell line',16),('Project(s)',22),('Records',10),('Max passage',12),
          ('Avg days/passage',16),('Records with fold change',20),('Avg fold change',16),
          ('Freeze stocks',12),('Experiments',12)]
    for ci, (l, w) in enumerate(c4, 1): hcell(ws4, 1, ci, l, w)
    lines_list = list(dict.fromkeys(r.get('line', '') for r in passages))
    for ri, line in enumerate(lines_list, 2):
        lp = [r for r in passages if r.get('line') == line]
        projs = list(dict.fromkeys(filter(None, (get_lineage_project(r, passages) for r in lp))))
        days = [r.get('days', 0) for r in lp if r.get('days', 0) > 0]
        avg_d = round(sum(days) / len(days), 1) if days else ''
        max_p = max((r.get('passageNum', 0) for r in lp), default=0)
        fcs = [f for f in (calc_fold_change(r, passages) for r in lp) if f is not None]
        avg_fc = round(sum(fcs) / len(fcs), 2) if fcs else ''
        vals = [line, ', '.join(projs), len(lp), max_p, avg_d, len(fcs), avg_fc,
                sum(1 for r in lp if r.get('action') == 'freeze'),
                sum(1 for r in lp if r.get('action') == 'experiment')]
        for ci, v in enumerate(vals, 1):
            dcell(ws4, ri, ci, v, align='center' if ci > 2 else 'left',
                  fmt='0.00"x"' if ci == 7 and v else None)

    if projects:
        ws5 = wb.create_sheet('Projects')
        hcell(ws5, 1, 1, 'Project name', 20)
        for ri, p in enumerate(projects, 2): dcell(ws5, ri, 1, p)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Excel preview (returns headers + sample rows for column mapping UI) ───────

def preview_xlsx(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = wb.sheetnames
    result = {}
    for sheet_name in sheets:
        ws = wb[sheet_name]
        headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
        headers = [h for h in headers if h]
        rows = []
        for row in ws.iter_rows(min_row=2, max_row=4):
            row_data = [str(c.value).strip() if c.value is not None else '' for c in row[:len(headers)]]
            rows.append(row_data)
        result[sheet_name] = {'headers': headers, 'sample_rows': rows}
    return {'sheets': list(sheets), 'preview': result}


# ── Excel import with column mapping ─────────────────────────────────────────

def parse_xlsx_with_mapping(file_bytes, mapping):
    from openpyxl import load_workbook
    import random, string
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet_name = mapping.get('sheet', wb.sheetnames[0])
    ws = wb[sheet_name]
    col_map = mapping.get('columns', {})
    action_map = mapping.get('action_map', {})
    default_action = mapping.get('default_action', 'passage')

    headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
    header_idx = {h: i for i, h in enumerate(headers)}

    def get_val(row, field):
        col_name = col_map.get(field, '')
        if not col_name or col_name not in header_idx: return ''
        v = row[header_idx[col_name]].value
        if v is None: return ''
        # handle datetime objects from openpyxl
        if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
        # handle Excel serial date numbers
        if isinstance(v, (int, float)) and field == 'date':
            try:
                from datetime import datetime, timedelta
                # Excel epoch is 1899-12-30
                d = datetime(1899, 12, 30) + timedelta(days=float(v))
                return d.strftime('%Y-%m-%d')
            except: pass
        sv = str(v).strip()
        # strip time component from date strings: '2025-02-18 00:00:00' -> '2025-02-18'
        if len(sv) > 10 and (sv[10] == ' ' or sv[10] == 'T') and ':' in sv[10:]:
            sv = sv[:10]
        return sv

    passages = []
    for row in ws.iter_rows(min_row=2):
        if all(c.value is None for c in row): continue
        rec_id = get_val(row, 'id')
        if not rec_id:
            rec_id = 'IMP-P-' + ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
        raw_action = get_val(row, 'action')
        action = action_map.get(raw_action, raw_action.lower() if raw_action else default_action)
        if action not in ('thaw', 'inherit', 'passage', 'freeze', 'experiment'):
            action = default_action
        try: p_num = int(float(get_val(row, 'passageNum')))
        except: p_num = 0
        try: days = int(float(get_val(row, 'days')))
        except: days = 0
        rec = {
            'id': rec_id, 'parent': get_val(row, 'parent') or None,
            'line': get_val(row, 'line'), 'project': get_val(row, 'project'),
            'action': action, 'passageNum': p_num, 'date': get_val(row, 'date'),
            'days': days, 'condition': get_val(row, 'condition'),
            'substrate': get_val(row, 'substrate'), 'wells': get_val(row, 'wells'),
            'splitRatio': get_val(row, 'splitRatio'), 'confluency': get_val(row, 'confluency'),
            'seedingDensity': get_val(row, 'seedingDensity'), 'seedingTotal': get_val(row, 'seedingTotal'),
            'viableCellsPerMl': get_val(row, 'viableCellsPerMl'), 'viabilityPct': get_val(row, 'viabilityPct'),
            'volumeMl': get_val(row, 'volumeMl'), 'totalViableCells': get_val(row, 'totalViableCells'),
            'user': get_val(row, 'user'), 'note': get_val(row, 'note'),
            'vials': get_val(row, 'vials'), 'cellsPerVial': get_val(row, 'cellsPerVial'),
            'cryo': get_val(row, 'cryo'), 'storage': get_val(row, 'storage'),
            'expId': get_val(row, 'expId'), 'assay': get_val(row, 'assay'),
            'treatment': get_val(row, 'treatment'), 'timepoints': get_val(row, 'timepoints'),
            'feeds': [], 'images': [], 'plateData': {},
        }
        # build vessels array from type + qty
        vessel_type = rec.get('wells','')
        vessel_qty_str = get_val(row, 'vesselQty')
        if vessel_type:
            try: vessel_qty = int(float(vessel_qty_str)) if vessel_qty_str else 1
            except: vessel_qty = 1
            vessel_class = 'plate' if any(p in vessel_type.lower() for p in ['well','plate','wp']) else 'flask'
            rec['vessels'] = [{'type': vessel_type, 'vesselClass': vessel_class, 'qty': vessel_qty,
                               'seeding': rec.get('seedingDensity') or None,
                               'seedingTotal': rec.get('seedingTotal') or None,
                               'seedingUnit': 'cells/cm2'}]
            rec['wells'] = '({}) {}'.format(vessel_qty, vessel_type)
        passages.append(rec)
    # clean up parents: thaw/inherit records pointing to non-existent IDs
    # are legitimately parentless (pointing to a physical vial, not a record)
    passage_ids = {r['id'] for r in passages}
    for rec in passages:
        if rec.get('parent') and rec['parent'] not in passage_ids:
            if rec.get('action') in ('thaw', 'inherit'):
                # store vial reference in note/vial field, clear parent
                if not rec.get('vial'): rec['vial'] = rec['parent']
                rec['parent'] = None
            # for passages with multiple parents (semicolon separated), take first
            elif rec.get('parent') and (';' in rec['parent'] or '+' in rec['parent']):
                parts = [p.strip() for p in rec['parent'].replace('+',';').split(';') if p.strip()]
                # use first part that exists, else clear
                rec['parent'] = next((p for p in parts if p in passage_ids), None)
    return {'passages': passages, 'projects': []}


# ── Excel import (our own format, no mapping needed) ─────────────────────────

def parse_xlsx(file_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    passages = {}

    if 'Records' in wb.sheetnames:
        ws = wb['Records']
        headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
        def col(row, name):
            try:
                idx = headers.index(name)
                v = row[idx].value
                if v is None: return ''
                if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
                sv = str(v).strip()
                if name in ('Date',) and len(sv) > 10 and sv[10] in (' ','T') and ':' in sv:
                    sv = sv[:10]
                return sv
            except (ValueError, IndexError): return ''
        for row in ws.iter_rows(min_row=2):
            rec_id = col(row, 'Record ID')
            if not rec_id or rec_id == 'None': continue
            action = col(row, 'Action')
            try: p_num = int(float(col(row, 'P#')))
            except: p_num = 0
            try: days = int(float(col(row, 'Days')))
            except: days = 0
            rec = {'id': rec_id, 'parent': col(row, 'Parent ID') or None,
                   'line': col(row, 'Line'), 'project': col(row, 'Project'),
                   'action': action, 'passageNum': p_num, 'date': col(row, 'Date'),
                   'days': days, 'condition': col(row, 'Condition'),
                   'substrate': col(row, 'Substrate'), 'wells': col(row, 'Vessel(s)'),
                   'splitRatio': col(row, 'Split ratio'), 'confluency': col(row, 'Confluency %'),
                   'seedingDensity': col(row, 'Seeding density'), 'seedingTotal': col(row, 'Total seeded'),
                   'viableCellsPerMl': col(row, 'Viable cells/mL'), 'viabilityPct': col(row, '% Viability'),
                   'volumeMl': col(row, 'Volume mL'), 'totalViableCells': col(row, 'Total viable cells'),
                   'user': col(row, 'Operator'), 'note': col(row, 'Notes'),
                   'feeds': [], 'images': [], 'plateData': {}}
            if action == 'freeze':
                rec.update({'vials': col(row, 'Vials'), 'cellsPerVial': col(row, 'Cells/vial'),
                            'cryo': col(row, 'Cryoprotectant'), 'storage': col(row, 'Storage'),
                            'viability': col(row, 'Vial viability %')})
            if action == 'experiment':
                rec.update({'expId': col(row, 'Exp ID'), 'assay': col(row, 'Assay'),
                            'treatment': col(row, 'Treatment'), 'timepoints': col(row, 'Timepoints')})
            passages[rec_id] = rec

    if 'Feed log' in wb.sheetnames:
        ws = wb['Feed log']
        headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
        def fcol(row, name):
            try:
                idx = headers.index(name); v = row[idx].value
                if v is None: return ''
                if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
                sv = str(v).strip()
                if name in ('Feed date','Date') and len(sv) > 10 and sv[10] in (' ','T') and ':' in sv:
                    sv = sv[:10]
                return sv
            except (ValueError, IndexError): return ''
        for row in ws.iter_rows(min_row=2):
            rid = fcol(row, 'Record ID')
            if rid and rid in passages:
                feed = {'date': fcol(row, 'Feed date'), 'media': fcol(row, 'Media type'),
                        'volumeMl': fcol(row, 'Volume mL'), 'user': fcol(row, 'Operator'),
                        'note': fcol(row, 'Notes'), 'mediaCat': fcol(row, 'Cat #'),
                        'mediaLot': fcol(row, 'Lot #'), 'mediaExp': fcol(row, 'Expiration')}
                if any(feed.values()): passages[rid]['feeds'].append(feed)

    if 'Plate wells' in wb.sheetnames:
        ws = wb['Plate wells']
        headers = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]
        def wcol(row, name):
            try:
                idx = headers.index(name); v = row[idx].value
                if v is None: return ''
                if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
                sv = str(v).strip()
                if name == 'Date' and len(sv) > 10 and sv[10] in (' ','T') and ':' in sv:
                    sv = sv[:10]
                return sv
            except (ValueError, IndexError): return ''
        for row in ws.iter_rows(min_row=2):
            rid = wcol(row, 'Record ID'); pk = wcol(row, 'Plate key'); wid = wcol(row, 'Well')
            if rid and rid in passages and pk and wid:
                if pk not in passages[rid]['plateData']: passages[rid]['plateData'][pk] = {}
                passages[rid]['plateData'][pk][wid] = {
                    'seeding': wcol(row, 'Seeding density'), 'count': wcol(row, 'Cell count'),
                    'cond': wcol(row, 'Condition'), 'treat': wcol(row, 'Treatment'),
                    'note': wcol(row, 'Notes'), 'occupied': True,
                    'contaminated': wcol(row, 'Contaminated').lower() == 'yes'}

    projects = []
    if 'Projects' in wb.sheetnames:
        ws = wb['Projects']
        for row in ws.iter_rows(min_row=2):
            v = row[0].value
            if v: projects.append(str(v).strip())

    return {'passages': list(passages.values()), 'projects': projects}


# ── HTTP handler ──────────────────────────────────────────────────────────────

class TCHandler(SimpleHTTPRequestHandler):

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/upload':
            self.handle_upload()
        elif parsed.path == '/export-excel':
            self.handle_excel_export()
        elif parsed.path == '/preview-excel':
            self.handle_excel_preview()
        elif parsed.path == '/generate-report':
            self.handle_generate_report()
        elif parsed.path == '/proxy-confluency':
            self.handle_proxy_confluency()
        elif parsed.path == '/get-settings':
            self.handle_get_settings()
        elif parsed.path == '/save-settings':
            self.handle_save_settings()
        elif parsed.path == '/import-excel':
            self.handle_excel_import()
        else:
            self.send_error(404, 'Not found')

    def handle_excel_export(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body)
            xlsx = build_excel(data)
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="tc-records.xlsx"')
            self.send_header('Content-Length', str(len(xlsx)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(xlsx)
            print('  [excel] exported {} records'.format(len(data.get('passages', []))))
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_json(500, {'error': str(e)})

    def handle_excel_preview(self):
        try:
            ct = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in ct:
                self.send_json(400, {'error': 'Expected multipart/form-data'}); return
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': ct})
            if 'file' not in form:
                self.send_json(400, {'error': 'No file field'}); return
            fi = form['file']
            file_bytes = fi.file.read()
            tmp_path = os.path.join(UPLOAD_DIR, '__import_tmp__.xlsx')
            with open(tmp_path, 'wb') as f: f.write(file_bytes)
            result = preview_xlsx(file_bytes)
            print('  [preview-excel] sheets:', result['sheets'])
            self.send_json(200, result)
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_json(500, {'error': str(e)})

    def handle_excel_import(self):
        try:
            ct = self.headers.get('Content-Type', '')
            if 'application/json' in ct:
                length = int(self.headers.get('Content-Length', 0))
                body = json.loads(self.rfile.read(length))
                tmp_path = os.path.join(UPLOAD_DIR, '__import_tmp__.xlsx')
                if not os.path.exists(tmp_path):
                    self.send_json(400, {'error': 'No file uploaded yet. Please upload the Excel file first.'}); return
                with open(tmp_path, 'rb') as f: file_bytes = f.read()
                result = parse_xlsx_with_mapping(file_bytes, body)
            elif 'multipart/form-data' in ct:
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers,
                    environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': ct})
                if 'file' not in form:
                    self.send_json(400, {'error': 'No file field'}); return
                fi = form['file']
                file_bytes = fi.file.read()
                result = parse_xlsx(file_bytes)
            else:
                self.send_json(400, {'error': 'Expected multipart/form-data or JSON'}); return
            print('  [import-excel] {} records'.format(len(result['passages'])))
            self.send_json(200, result)
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_json(500, {'error': str(e)})

    def handle_upload(self):
        try:
            ct = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in ct:
                self.send_json(400, {'error': 'Expected multipart/form-data'}); return
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': ct})
            if 'file' not in form:
                self.send_json(400, {'error': 'No file field'}); return
            fi = form['file']
            fn = os.path.basename(fi.filename)
            fn = ''.join(c for c in fn if c.isalnum() or c in '._- ').strip()
            if not fn: self.send_json(400, {'error': 'Invalid filename'}); return
            if 'filename' in form:
                cn = ''.join(c for c in os.path.basename(form['filename'].value) if c.isalnum() or c in '._- ')
                if cn: fn = cn
            dest = os.path.join(UPLOAD_DIR, fn)
            with open(dest, 'wb') as f: shutil.copyfileobj(fi.file, f)
            size = os.path.getsize(dest)
            print('  [upload] {} ({:,} bytes)'.format(fn, size))
            self.send_json(200, {'filename': fn, 'size': size})
        except Exception as e:
            print('  [upload error]', e)
            self.send_json(500, {'error': str(e)})

    def handle_generate_report(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length).decode()
            # Write temp json file
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                f.write(body)
                tmp_json = f.name
            # Output pptx path
            out_pptx = tmp_json.replace('.json', '.pptx')
            # Find generate_report.js next to server.py
            script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'generate_report.js')
            if not os.path.exists(script):
                self.send_json(404, {'error': 'generate_report.js not found next to server.py'})
                return
            result = subprocess.run(
                ['node', script, body, out_pptx],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode != 0 or not os.path.exists(out_pptx):
                self.send_json(500, {'error': result.stderr or 'Report generation failed'})
                return
            with open(out_pptx, 'rb') as f:
                pptx_data = f.read()
            os.unlink(tmp_json)
            os.unlink(out_pptx)
            fn = (json.loads(body).get('rec', {}).get('expId') or 'experiment_report') + '.pptx'
            fn = fn.replace('/', '_').replace(' ', '_')
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition', 'attachment; filename="{}"'.format(fn))
            self.send_header('Content-Length', str(len(pptx_data)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(pptx_data)
            print('  [report] generated {}'.format(fn))
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_json(500, {'error': str(e)})

    def handle_get_settings(self):
        self.send_json(200, load_settings())

    def handle_save_settings(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            settings = load_settings()
            settings.update(body)
            save_settings(settings)
            self.send_json(200, {'ok': True, 'settings': settings})
        except Exception as e:
            self.send_json(500, {'error': str(e)})

    def handle_proxy_confluency(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            settings = load_settings()
            api_url = settings.get('confluency_api_url', '').rstrip('/')
            if not api_url:
                self.send_json(400, {'error': 'Confluency API URL not configured. Set it in TC Tracker settings.'})
                return
            req = _urllib_req.Request(
                api_url + '/analyze',
                data=body,
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            with _urllib_req.urlopen(req, timeout=120) as resp:
                result = json.loads(resp.read())
            print('  [confluency proxy] {} images analyzed'.format(result.get('summary', {}).get('analyzed', '?')))
            self.send_json(200, result)
        except Exception as e:
            self.send_json(500, {'error': str(e)})

    def send_json(self, code, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, fmt, *args):
        if args and str(args[1]) not in ('200', '304'):
            super().log_message(fmt, *args)


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port = 8081
    server = HTTPServer(('localhost', port), TCHandler)
    print('TC Tracker running at http://localhost:{}'.format(port))
    print('Images folder: {}'.format(UPLOAD_DIR))
    print('Press Ctrl+C to stop\n')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nStopped.')