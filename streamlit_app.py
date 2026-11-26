from html import escape
from io import BytesIO
from math import isfinite
from pathlib import Path
import json
import logging
import os
import subprocess
import sys
from tempfile import TemporaryDirectory
from zipfile import BadZipFile, ZipFile

import altair as alt
import pandas as pd
import streamlit as st

HERE = Path(__file__).resolve().parent
COLUMNS = ('농도 하한', '농도 상한', '검체 수')
DEFAULT_BINS = [[.5, 5., 24], [5., 15., 12], [15., 30., 4]]
DEVICES = ('F10', 'FREND 2.0', 'Comparator 1')
DEVICE_INDEX = (1, 2, 0)
COMPARE = ((0, 1, 1, 'Comparator 1 (X) → F10 (Y)'),
           (1, 2, 0, 'F10 (X) → FREND 2.0 (Y)'),
           (0, 2, 2, 'Comparator 1 (X) → FREND 2.0 (Y)'))
MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
LOG = logging.getLogger(__name__)


def read_bins(table):
    bins = []
    for i, (lo, hi, count) in enumerate(table.itertuples(index=False, name=None), 1):
        if any(pd.isna(value) for value in (lo, hi, count)):
            raise ValueError(f'{i}번 구간의 빈 칸을 입력해 주세요.')
        if not (isfinite(lo) and isfinite(hi) and 0 <= lo < hi <= 1e12):
            raise ValueError(f'{i}번 구간에서 하한이 상한보다 작아야 합니다.')
        if not (0 <= count <= 200 and int(count) == count):
            raise ValueError(f'{i}번 구간의 검체 수는 0~200 사이의 정수여야 합니다.')
        if bins and bins[-1][1] != lo:
            raise ValueError(f'{i-1}번 구간의 상한과 {i}번 구간의 하한을 같게 입력해 주세요.')
        bins.append([float(lo), float(hi), int(count)])
    return bins


def resize_bins(rows, size):
    if len(rows) == size:
        return rows
    if size < len(rows):
        tail = rows[size-1:]
        return rows[:size-1] + [[tail[0][0], tail[-1][1], sum(row[2] for row in tail)]]
    lo, hi, count = rows[-1]
    pieces = size-len(rows)+1
    edges = [lo + (hi-lo)*i/pieces for i in range(pieces)] + [hi]
    if any(a >= b for a, b in zip(edges, edges[1:])):
        raise ValueError('마지막 농도 구간이 너무 좁아 이 개수로 나눌 수 없습니다.')
    q, remainder = divmod(count, pieces)
    return rows[:-1] + [[a, b, q + int(i < remainder)] for i,(a,b) in enumerate(zip(edges,edges[1:]))]


def apply_bin_edits(bins, edits):
    updated = [row.copy() for row in bins]
    for raw_index, changes in edits.items():
        index = int(raw_index)
        for label, value in changes.items():
            column = COLUMNS.index(label)
            updated[index][column] = value
            if column == 0 and index:
                updated[index-1][1] = value
            if column == 1 and index < len(updated)-1:
                updated[index+1][0] = value
    return updated


def save_editor_changes(key):
    state = st.session_state
    edits = state[key]['edited_rows']
    if edits:
        state.bin_base = apply_bin_edits(state.bin_base, edits)
        state.bin_revision += 1


def template_bytes(upload):
    if upload:
        data = upload.getvalue()
    else:
        configured = os.environ.get('FREND_TEMPLATE_PATH')
        path = Path(configured) if configured else HERE / 'FREND_Template.xlsx'
        if not path.is_file():
            return None
        data = path.read_bytes()
    if len(data) > 20 * 1024**2:
        raise ValueError('20MB 이하인 엑셀 양식을 선택해 주세요.')
    try:
        with ZipFile(BytesIO(data)) as source:
            if sum(item.file_size for item in source.infolist()) > 100 * 1024**2:
                raise ValueError('압축 해제 용량이 100MB 이하인 양식을 선택해 주세요.')
            workbook = source.read('xl/workbook.xml')
            if b'name="Method comparison"' not in workbook:
                raise ValueError('양식에 Method comparison 시트가 필요합니다.')
    except (BadZipFile, KeyError) as exc:
        raise ValueError('Excel에서 열 수 있는 원본 .xlsx 양식을 선택해 주세요.') from exc
    return data


def generate(data, bins, seed, slopes, intercepts, comparator_sd, frend_sd):
    counts = [row[2] for row in bins]
    if not 4 <= sum(counts) <= 200:
        raise ValueError('전체 검체 수를 4~200개로 설정해 주세요.')
    with TemporaryDirectory(prefix='frend-streamlit-') as temp:
        source = Path(temp) / 'template.xlsx'
        output = Path(temp) / 'result.xlsx'
        command = [sys.executable, str(HERE / 'workbook_engine.py'), '--source', str(source),
                   '--output', str(output), '--seed', str(seed), '--counts', *map(str, counts),
                   '--bounds', *(str(edge) for lo,hi,_ in bins for edge in (lo,hi)),
                   '--slopes', *(f'{value:.15f}' for value in slopes),
                   '--intercepts', *(f'{value:.15f}' for value in intercepts),
                   '--between-method-sd', str(comparator_sd), '--between-frend-sd', str(frend_sd), '--json']
        source.write_bytes(data)
        try:
            proc = subprocess.run(command, capture_output=True, timeout=50, check=True)
            result = json.loads(proc.stdout)
            result['workbook'] = output.read_bytes()
        except (OSError, subprocess.SubprocessError, ValueError) as exc:
            LOG.warning('Workbook engine failed: %s', type(exc).__name__)
            raise ValueError('이 Excel 양식으로 결과를 생성할 수 없습니다. 기존 분석적 성능시험 양식을 확인해 주세요.') from exc
    result['filename'] = f'FREND_Method_Comparison_SIM_{seed}.xlsx'
    result['seed'] = seed
    return result


def scatter_chart(means, fit, x, y, domain):
    df = pd.DataFrame({'X': means[x], 'Y': means[y], '검체': [f'SIM-{i:03}' for i in range(1,len(means[x])+1)]})
    encoding = dict(x=alt.X('X:Q', scale=alt.Scale(domain=domain, nice=False)),
                    y=alt.Y('Y:Q', scale=alt.Scale(domain=domain, nice=False)))
    points = alt.Chart(df).mark_circle(color='#1766b4', opacity=.7, size=30).encode(
        **encoding, tooltip=['검체:N','X:Q','Y:Q'])
    equality = alt.Chart(pd.DataFrame({'X': domain, 'Y': domain})).mark_line(
        color='#78899c', strokeDash=[5,4], clip=True).encode(**encoding)
    regression = alt.Chart(pd.DataFrame({'X': domain, 'Y': [fit['slope']*v+fit['intercept'] for v in domain]})).mark_line(
        color='#b45129', clip=True).encode(**encoding)
    return (equality + regression + points).properties(height=200)


def replicates_html(run):
    groups = ''.join(f'<th scope="colgroup" colspan="3">{escape(device)}</th>' for device in DEVICES)
    subhead = '<th scope="col">Rep 1</th><th scope="col">Rep 2</th><th scope="col">Mean</th>' * len(DEVICES)
    rows = []
    for i, sample in enumerate(run['replicates']):
        data = ''.join(f'<td>{sample[dev][0]:.4f}</td><td>{sample[dev][1]:.4f}</td><td class="mean">{run["mean"][dev][i]:.5f}</td>'
                       for dev in DEVICE_INDEX)
        rows.append(f'<tr><th scope="row">SIM-{i+1:03}</th>{data}</tr>')
    return '''<style>
.frend-reps{max-height:430px;overflow:auto;border:1px solid #d9e4ef;border-radius:.45rem}
.frend-reps table{border-collapse:collapse;width:100%;min-width:620px;font-size:.86rem;font-variant-numeric:tabular-nums}
.frend-reps th,.frend-reps td{padding:.46rem .55rem;border:1px solid #d9e4ef;text-align:right}
.frend-reps thead{position:sticky;top:0;background:#eaf2fc}.frend-reps thead th{text-align:center}
.frend-reps th[scope=row]{text-align:left;white-space:nowrap}.frend-reps .mean{color:#155492;font-weight:650}
</style>''' + f'<div class="frend-reps"><table><thead><tr><th rowspan="2" scope="col">검체</th>{groups}</tr><tr>{subhead}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


st.set_page_config(page_title='FREND Performance Simulator', layout='wide')
st.title('FREND Performance Simulator')
st.caption('Method Comparison · 합성 데이터')
left, right = st.columns([.9,1.45], gap='large')
with left:
    st.subheader('생성 조건')
    upload = st.file_uploader('분석적 성능시험 엑셀 양식', type=['xlsx'],
                             help='승인된 빈 양식을 사용하세요. 이 파일은 현재 세션의 엑셀 생성에 사용됩니다.')
    state = st.session_state
    if 'bin_base' not in state:
        state.bin_base = DEFAULT_BINS
        state.last_valid_bins = DEFAULT_BINS
        state.bin_revision = 0

    size = st.number_input('농도 구간 개수', min_value=1, max_value=200, value=3, step=1, key='bin_count')
    if size != len(state.bin_base):
        try:
            state.bin_base = resize_bins(state.last_valid_bins, size)
            state.last_valid_bins = state.bin_base
            state.bin_revision += 1
        except ValueError as error:
            st.warning(str(error))

    editor_key = f'bins_{state.bin_revision}'
    table = st.data_editor(pd.DataFrame(state.bin_base, columns=COLUMNS), hide_index=True,
                           key=editor_key, num_rows='fixed', width='stretch',
                           on_change=save_editor_changes, args=(editor_key,),
                           column_config={
                               '농도 하한': st.column_config.NumberColumn(min_value=0., required=True),
                               '농도 상한': st.column_config.NumberColumn(min_value=0., required=True),
                               '검체 수': st.column_config.NumberColumn(min_value=0, max_value=200, step=1, required=True)})
    st.caption(f'총 검체 수 {int(table[COLUMNS[2]].sum())} · 인접 구간의 농도 경계는 같이 바뀝니다.')
    bins = None
    try:
        bins = read_bins(table)
        state.last_valid_bins = bins
    except ValueError as error:
        st.warning(str(error))

    st.markdown('**Comparator 1 (X) → FREND (Y)**')
    a,b = st.columns(2)
    slope1 = a.number_input('F10 slope', min_value=.01, max_value=100., value=1.03, step=.01)
    int1 = b.number_input('F10 intercept', min_value=-1e6, max_value=1e6, value=.03, step=.01)
    slope2 = a.number_input('FREND 2.0 slope', min_value=.01, max_value=100., value=1.031, step=.001, format='%.4f')
    int2 = b.number_input('FREND 2.0 intercept', min_value=-1e6, max_value=1e6, value=.03, step=.01)
    c_sd = st.number_input('Comparator–FREND 검체별 상대 SD (%)', 0., 100., 20., step=.1)
    f_sd = st.number_input('F10–FREND 2.0 검체별 상대 SD (%)', 0., 100., .4, step=.1)
    seed = st.number_input('난수 seed', min_value=0, max_value=4294967295, value=261125, step=1,
                           help='설정과 seed가 같으면 같은 데이터를 다시 생성합니다.')
    if st.button('데이터 생성', type='primary', key='generate', disabled=bins is None or len(table) != size):
        try:
            data = template_bytes(upload)
            if data is None:
                raise ValueError('분석적 성능시험 엑셀 양식을 먼저 업로드해 주세요.')
            with st.spinner('합성 데이터와 엑셀을 생성하는 중입니다.'):
                state.result = generate(data, bins, seed, [slope1,slope2], [int1,int2], c_sd,f_sd)
        except (ValueError,OSError) as error:
            st.error(str(error))

with right:
    st.subheader('생성 결과')
    result = st.session_state.get('result')
    if result is None:
        st.info('엑셀 양식을 선택한 뒤 데이터를 생성하면 결과가 표시됩니다.')
    else:
        st.caption(f"생성된 검체 {len(result['replicates'])}개 · seed {result['seed']}")
        graphs, reps = st.tabs(['그래프', '반복값'])
        with graphs:
            st.caption('회색 점선 y=x · 주황색 실선 OLS')
            means = result['mean']
            domain = [min(0., *(v for device in means for v in device)),
                      1.04*max(1., *(v for device in means for v in device))]
            slots = st.columns(2)
            for order,(x,y,i,title) in enumerate(COMPARE):
                with slots[order%2]:
                    fit = result['fits'][i]
                    st.markdown(f'**{title}**')
                    st.caption(f"Pearson r = {result['correlations'][i]:.4f} · OLS Y = {fit['slope']:.4f}X {fit['intercept']:+.4f}")
                    st.altair_chart(scatter_chart(means,fit,x,y,domain), width='stretch', key=f'graph_{order}')
        with reps:
            st.html(replicates_html(result))
        st.download_button('엑셀 다운로드', result['workbook'], file_name=result['filename'],
                           mime=MIME, on_click='ignore', key='xlsx')

