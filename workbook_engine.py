from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from random import Random
from statistics import mean, correlation, linear_regression
import argparse, json, posixpath, re, tempfile

SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SHEET = 'Simulated Raw'
CELLS = ('C', 'D', 'E')
RAW_PART = 'xl/worksheets/simulator_raw.xml'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
args = argparse.ArgumentParser()
args.add_argument('--source', type=Path)
args.add_argument('--basic', action='store_true')
args.add_argument('--output', type=Path)
args.add_argument('--seed', type=int, default=261125)
args.add_argument('--counts', type=int, nargs='+', default=(24,12,4))
args.add_argument('--bounds', type=float, nargs='+', default=(.5,5,5,15,15,30))
args.add_argument('--slopes', type=float, nargs=2, default=(1.03,1.031))
args.add_argument('--intercepts', type=float, nargs=2, default=(.03,.03))
args.add_argument('--between-method-sd', type=float, default=20)
args.add_argument('--between-frend-sd', type=float, default=.4)
args.add_argument('--json', action='store_true')
options = args.parse_args()
if len(options.bounds) != len(options.counts) * 2:
    args.error('Provide a lower and upper bound for each concentration bin.')
if options.basic and options.source:
    args.error('Choose the built-in layout or an Excel template.')
rng = Random(options.seed)
levels = ((0.5, 15), (2, 8), (5, 5), (30, 3))
def sd(x):
    if x < levels[0][0]:
        return levels[0][0] * levels[0][1] / 100
    if x > levels[-1][0]:
        return x * levels[-1][1] / 100
    for (a, acv), (b, bcv) in zip(levels, levels[1:]):
        if a <= x <= b:
            return (a * acv + (x-a)/(b-a)*(b*bcv-a*acv))/100
    raise ValueError(f'No precision point for {x}')
def replicates(x):
    shared = rng.gauss(0, x * options.between_method_sd / 100)
    f10 = options.slopes[0] * x + options.intercepts[0] + shared
    f20 = options.slopes[1] * x + options.intercepts[1] + shared + rng.gauss(0, x * options.between_frend_sd / 100)
    for base, sigma in ((x, x*.035), (f10, sd(x)), (f20, sd(x))):
        yield [round(rng.gauss(base,sigma), 4) for _ in range(2)]

def patch_cells(xml, updates):
    pattern = r'(<c\b[^>]*?\br="([A-Z]+\d+)"[^>]*?)(?:/>|>.*?</c>)'
    changed = set()
    def replace(match):
        ref = match[2]
        if ref not in updates:
            return match[0]
        changed.add(ref)
        content = updates[ref]
        attrs = re.sub(r' t="[^"]*"', '', match[1])
        return attrs + (' t="inlineStr"' if content.startswith('<is>') else '') + '>' + content + '</c>'
    xml = re.sub(pattern, replace, xml, flags=re.S)
    assert changed == set(updates), sorted(set(updates)-changed)
    return xml
def string(col, row, value):
    return f'<c r="{col}{row}" t="inlineStr"><is><t>{value}</t></is></c>'
def number(column, row, value, formula=None):
    f = f'<f>{formula}</f>' if formula is not None else ''
    return f'<c r="{column}{row}">{f}<v>{value}</v></c>'

def hide_intro(xml):
    pattern = r'(<row\b[^>]*?\br=")([2-9]|[1-8][0-9]|9[0-4])("[^>]*?)(?:/>|>.*?</row>)'
    xml = re.sub(pattern, lambda match: f'<row r="{match[2]}" hidden="1"/>', xml, flags=re.S)
    def remove_intro_merges(match):
        entries = re.findall(r'<mergeCell\b[^>]+/>', match[0])
        kept = [entry for entry in entries
                if not any(2 <= int(row) <= 94 for row in re.findall(r'[A-Z]+(\d+)', entry))]
        return f'<mergeCells count="{len(kept)}">{"".join(kept)}</mergeCells>' if kept else ''
    return re.sub(r'<mergeCells\b[^>]*>.*?</mergeCells>', remove_intro_merges, xml, flags=re.S)

def filter_relationships(xml, keep):
    return re.sub(r'<Relationship\b[^>]+/>',
                  lambda match: match[0] if keep(ET.fromstring(match[0])) else '', xml)

def dependencies(parts):
    needed = {'_rels/.rels'}
    pending = [('', '_rels/.rels')]
    while pending:
        folder, rel_file = pending.pop()
        for relationship in ET.fromstring(parts(rel_file)).findall(f'{{{PKG_REL}}}Relationship'):
            if relationship.get('TargetMode') == 'External':
                continue
            target = relationship.get('Target')
            path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(folder, target))
            if path in needed:
                continue
            needed.add(path)
            target_folder, filename = posixpath.split(path)
            companion = posixpath.join(target_folder, '_rels', filename + '.rels')
            if parts(companion, missing_ok=True) is not None:
                needed.add(companion)
                pending.append((target_folder, companion))
    return needed

def compact_strings(analysis, original_strings):
    indices = {}
    def replace_cell(match):
        old = int(match[2])
        return f'{match[1]}{indices.setdefault(old, len(indices))}{match[3]}'
    cell_pattern = r'(<c\b[^>]*\bt="s"[^>]*><v>)(\d+)(</v>)'
    analysis, references = re.subn(cell_pattern, replace_cell, analysis)
    entries = re.findall(r'<si(?:\s[^>]*)?>.*?</si>', original_strings, re.S)
    header = original_strings[:original_strings.index('<si>')]
    header = re.sub(r'\bcount="\d+"', f'count="{references}"', header)
    header = re.sub(r'\buniqueCount="\d+"', f'uniqueCount="{len(indices)}"', header)
    strings = header + ''.join(entries[index] for index in indices) + '</sst>'
    return analysis, strings

def write_basic_workbook(out, raw_xml, means):
    headers = ('Synthetic sample','Comparator 1 mean','F10 mean','FREND 2.0 mean')
    rows = [''.join(string(col,1,label) for col,label in zip('ABCD',headers))]
    for i, sample in enumerate(zip(*means,strict=True), 2):
        averages = ''.join(number(column,i,value,f"'{SHEET}'!{raw_column}{i+1}")
                           for column,raw_column,value in zip('BCD','JDG',sample,strict=True))
        rows.append(string('A',i,f'SIM-{i-1:03}') + averages)
    headings = ('X device','Y device','N','OLS Slope','OLS Intercept','Pearson r')
    rows[0] += ''.join(string(col,1,label) for col,label in zip('GHIJKL',headings))
    end = len(rows)
    for row, (x, y) in enumerate(((0,1),(0,2),(1,2)), 2):
        x_col, y_col = 'BCD'[x], 'BCD'[y]
        xs, ys = f'{x_col}2:{x_col}{end}', f'{y_col}2:{y_col}{end}'
        fit = linear_regression(means[x],means[y])
        rows[row-1] += string('G',row,headers[x+1].removesuffix(' mean'))
        rows[row-1] += string('H',row,headers[y+1].removesuffix(' mean'))
        rows[row-1] += number('I',row,end-1,f'COUNT({xs})')
        rows[row-1] += number('J',row,fit.slope,f'SLOPE({ys},{xs})')
        rows[row-1] += number('K',row,fit.intercept,f'INTERCEPT({ys},{xs})')
        rows[row-1] += number('L',row,correlation(means[x],means[y]),f'CORREL({xs},{ys})')
    data = ''.join(f'<row r="{i}">{cells}</row>' for i,cells in enumerate(rows,1))
    analysis = (f'<worksheet xmlns="{SS}"><cols><col min="1" max="4" width="24" customWidth="1"/>'
                f'<col min="7" max="12" width="17" customWidth="1"/></cols><sheetData>{data}</sheetData></worksheet>')
    entries = {'xl/workbook.xml': 'sheet.main', 'xl/worksheets/analysis.xml': 'worksheet', RAW_PART: 'worksheet'}
    overrides = ''.join(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.{part}+xml"/>'
                        for path,part in entries.items())
    types = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>' + overrides + '</Types>')
    links = f'<Relationships xmlns="{PKG_REL}">'
    workbook = (f'<workbook xmlns="{SS}" xmlns:r="{REL}"><sheets>'
                f'<sheet name="Method comparison" sheetId="1" r:id="rId1"/>'
                f'<sheet name="{SHEET}" sheetId="2" r:id="rId2"/></sheets>'
                '<calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>')
    with ZipFile(out,'w',ZIP_DEFLATED) as package:
        package.writestr('[Content_Types].xml',types)
        package.writestr('_rels/.rels',links + f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        package.writestr('xl/_rels/workbook.xml.rels',links + f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/analysis.xml"/>'
                         f'<Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/simulator_raw.xml"/></Relationships>')
        package.writestr('xl/workbook.xml',workbook)
        package.writestr('xl/worksheets/analysis.xml',analysis)
        package.writestr(RAW_PART,raw_xml)

root = Path(tempfile.gettempdir())
out = options.output or Path(tempfile.mkdtemp(prefix='frend-pilot-'))/'synthetic-linkage-check.xlsx'
groups = (('F10', 1, 'BCD'), ('FREND 2.0', 2, 'EFG'), ('Comparator 1', 0, 'HIJ'))
rows = [
    string('A',1,'Synthetic sample') + ''.join(string(columns[0],1,name) for name,_,columns in groups),
    ''.join(string(column,2,label) for _,_,columns in groups
            for column,label in zip(columns,('Rep 1','Rep 2','Mean'))),
]
bins = tuple(zip(options.bounds[::2], options.bounds[1::2], options.counts, strict=True))
settings = [
    ('Simulation seed', options.seed), ('Comparator 1 X / F10 Y: slope', options.slopes[0]),
    ('Comparator 1 X / F10 Y: intercept', options.intercepts[0]),
    ('Comparator 1 X / FREND 2.0 Y: slope', options.slopes[1]),
    ('Comparator 1 X / FREND 2.0 Y: intercept', options.intercepts[1]),
    ('Comparator–FREND specimen SD (%)', options.between_method_sd),
    ('F10–FREND 2.0 specimen SD (%)', options.between_frend_sd)]
for i, (lo,hi,count) in enumerate(bins, 1):
    settings.extend(((f'Bin {i} lower',lo),(f'Bin {i} upper',hi),(f'Bin {i} count',count)))
values = []
linked = {}
for lo,hi,count in bins:
    for _ in range(count):
        sample = len(values) + 1
        triplet = tuple(replicates(rng.uniform(lo,hi)))
        values.append(triplet)
        row = sample + 2
        cells = [string('A',row,f'SIM-{sample:03}')]
        for _, device, (col1,col2,col_mean) in groups:
            pair = triplet[device]
            average = mean(pair)
            cells.extend((number(col1,row,pair[0]), number(col2,row,pair[1]),
                          number(col_mean,row,average,f'AVERAGE({col1}{row}:{col2}{row})')))
            linked[f'{CELLS[device]}{98+sample}']=(col_mean,row,average,col1,col2)
        rows.append(''.join(cells))
for row, (label, value) in enumerate(settings, 1):
    if row > len(rows):
        rows.append('')
    rows[row-1] += string('L',row,label) + number('M',row,value)
sheet_data = ''.join(f'<row r="{i}">{cells}</row>' for i,cells in enumerate(rows,1))
merges = '<mergeCells count="4"><mergeCell ref="A1:A2"/><mergeCell ref="B1:D1"/><mergeCell ref="E1:G1"/><mergeCell ref="H1:J1"/></mergeCells>'
raw_xml = (f'<worksheet xmlns="{SS}"><cols><col min="1" max="10" width="16" customWidth="1"/>'
           f'<col min="12" max="12" width="47" customWidth="1"/></cols><sheetData>{sheet_data}</sheetData>{merges}</worksheet>')

means=[[mean(pairs[device]) for pairs in values] for device in range(3)]
def print_summary():
    pairs = [(means[1],means[2]), (means[0],means[1]), (means[0],means[2])]
    print(json.dumps({'replicates':values,'mean':means,'correlations':
        [correlation(x,y) for x,y in pairs],
        'fits':[linear_regression(x,y)._asdict() for x,y in pairs]}))

if options.basic:
    write_basic_workbook(out, raw_xml, means)
    if options.json:
        print_summary()
    raise SystemExit(0)

source = options.source or max(root.glob('frend-template-*.xlsx'), key=lambda file:file.stat().st_mtime)
with ZipFile(source) as original:
    workbook = original.read('xl/workbook.xml').decode()
    rel_path = 'xl/_rels/workbook.xml.rels'
    links = original.read(rel_path).decode()
    rels = {x.attrib['Id']:x.attrib['Target'] for x in ET.fromstring(links)}
    book = ET.fromstring(workbook)
    analysis_sheet = book.find(f'.//{{{SS}}}sheet[@name="Method comparison"]')
    analysis_id = analysis_sheet.get(f'{{{REL}}}id')
    target = rels[analysis_id]
    path = target.lstrip('/') if target.startswith('/') else f'xl/{target}'
    analysis = original.read(path).decode()
    updates = {}
    for row in range(99,299):
        for col in CELLS:
            ref=f'{col}{row}'
            row_info = linked.get(ref)
            payload = '' if not row_info else f"<f>'{SHEET}'!{row_info[0]}{row_info[1]}</f><v>{row_info[2]}</v>"
            updates[ref] = payload
    for ref,name in {'C98':'Comparator 1','D98':'F10','E98':'FREND 2.0',
                     'Q99':'Comparator 1','R99':'F10','Q100':'Comparator 1','R100':'FREND 2.0',
                     'Q101':'F10','R101':'FREND 2.0'}.items():
        updates[ref] = f'<is><t>{name}</t></is>'
    analysis = patch_cells(analysis, updates)
    analysis = hide_intro(analysis)
    original_sheets = book.find(f'{{{SS}}}sheets')
    original_index = list(original_sheets).index(analysis_sheet)
    keep_sheet = next(tag[0] for tag in re.finditer(r'<sheet\b[^>]*/>',workbook)
                      if f'r:id="{analysis_id}"' in tag[0])
    new_sheets = f'<sheets>{keep_sheet}<sheet name="{SHEET}" sheetId="900" r:id="rId900"/></sheets>'
    workbook = re.sub(r'<sheets>.*?</sheets>', lambda _:new_sheets,workbook,flags=re.S)
    def selected_names(match):
        names = re.findall(r'<definedName\b[^>]*>.*?</definedName>', match[0], re.S)
        selected = [name.replace(f'localSheetId="{original_index}"', 'localSheetId="0"')
                    for name in names if f'localSheetId="{original_index}"' in name]
        return '<definedNames>' + ''.join(selected) + '</definedNames>' if selected else ''
    workbook = re.sub(r'<definedNames>.*?</definedNames>', selected_names, workbook, flags=re.S)
    workbook = re.sub(r'<externalReferences>.*?</externalReferences>', '', workbook, flags=re.S)
    workbook = re.sub(r'(?<=\s)(activeTab|firstSheet)="\d+"', r'\1="0"', workbook)
    calc = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    workbook = re.sub(r'<calcPr\b[^>]*/>',calc,workbook,count=1)
    links = filter_relationships(links, lambda link:
        link.get('Id') == analysis_id or link.get('Type').split('/')[-1] in
        {'styles', 'theme', 'sharedStrings', 'sheetMetadata'})
    links = links.replace('</Relationships>',f'<Relationship Id="rId900" Type="{REL}/worksheet" Target="worksheets/simulator_raw.xml"/></Relationships>')
    root_links = filter_relationships(original.read('_rels/.rels').decode(),
        lambda link: link.get('Type').split('/')[-1] in {'officeDocument','core-properties'})
    analysis, strings = compact_strings(analysis, original.read('xl/sharedStrings.xml').decode())
    edits = {'xl/workbook.xml':workbook,rel_path:links,'_rels/.rels':root_links,
             path:analysis,'xl/sharedStrings.xml':strings,RAW_PART:raw_xml}
    def part(name, missing_ok=False):
        if name in edits:
            return edits[name]
        try:
            return original.read(name)
        except KeyError:
            if not missing_ok:
                raise
    needed = dependencies(part)
    types = original.read('[Content_Types].xml').decode()
    types = re.sub(r'<Override\b[^>]+/>', lambda match:
        match[0] if ET.fromstring(match[0]).get('PartName').lstrip('/') in needed else '', types)
    types = types.replace('</Types>','<Override PartName="/xl/worksheets/simulator_raw.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    with ZipFile(out,'w',ZIP_DEFLATED) as output:
        output.writestr('[Content_Types].xml', types)
        for member in sorted(needed):
            output.writestr(member, part(member))
    with ZipFile(out) as output:
        assert output.testzip() is None
        for member in edits:
            ET.fromstring(output.read(member))
        exported_sheets = ET.fromstring(output.read('xl/workbook.xml')).find(f'{{{SS}}}sheets')
        assert [sheet.get('name') for sheet in exported_sheets] == ['Method comparison', SHEET]
        def sheet_cells(xml):
            return {cell.get('r'):cell for cell in ET.fromstring(xml).iter(f'{{{SS}}}c')}
        original_cells = sheet_cells(original.read(path))
        edited_cells = sheet_cells(output.read(path))
        raw_cells = sheet_cells(output.read('xl/worksheets/simulator_raw.xml'))
        def field(cell, name):
            return cell.findtext(f'{{{SS}}}{name}')
        for address, (column,row,expected,col1,col2) in linked.items():
            assert field(edited_cells[address], 'f') == f"'{SHEET}'!{column}{row}"
            assert field(raw_cells[f'{column}{row}'], 'f') == f'AVERAGE({col1}{row}:{col2}{row})'
            raw_pair = [float(field(raw_cells[f'{col}{row}'], 'v')) for col in (col1,col2)]
            assert mean(raw_pair) == expected
        for address in ('AB96','AC96','AB99','AC99','N111','Q111','AG300'):
            assert field(edited_cells[address], 'f')
            assert field(edited_cells[address], 'f') == field(original_cells[address], 'f')
        chart_parts=[x for x in needed if re.fullmatch(r'xl/charts/chart\d+\.xml',x)]
        assert chart_parts and all(output.read(x)==original.read(x) for x in chart_parts)

if options.json:
    print_summary()
else:
    print(f'Verified {len(values)} paired specimens, {len(values)*3*2} raw numbers, {len(linked)} mean links, {len(chart_parts)} unchanged chart parts.')
    print(f'Synthetic Pearson r, F10 to F20: {correlation(means[1],means[2]):.6f}; Comparator to F10: {correlation(means[0],means[1]):.6f}')
    print('OOXML verified. Native Excel recalculation was not executed.')

