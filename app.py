import os
import shutil
import subprocess
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

from roster import (ID_HEADERS, NAME_HEADERS, RESULT_HEADERS, Roster, clean,
                    columns, compare_rows, header_index, photo_rows,
                    read_tables, result_workbook)


BASE = Path(__file__).resolve().parent
LOCAL_ROSTER = BASE / '底库.xlsx'
EXISTING_ROSTER = BASE.parent / '群成员名单（学号与姓名）.xlsx'
TABLE_TYPES = ['xlsx', 'xls', 'csv']


def show_results(labels, rows, source):
    groups = {'已匹配': [], '待核对': [], '未匹配': []}
    for row in rows:
        status = row[-2]
        group = '待核对' if status.startswith('待核对') else '未匹配' if status == '未匹配' else '已匹配'
        groups[group].append(row)
    if not rows:
        st.info('没有可比对的姓名或学号。')
        return
    headers = [*labels, *RESULT_HEADERS]
    for tab, data in zip(st.tabs([f'{label} · {len(data)}' for label, data in groups.items()]), groups.values()):
        with tab:
            if data:
                st.dataframe(pd.DataFrame(data, columns=headers), hide_index=True, width='stretch')
            else:
                st.caption('此组没有记录')
    st.download_button('下载 Excel 结果', result_workbook(headers, groups),
                       file_name=f'{source}比对结果.xlsx', key=f'download-{source}', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def table_compare(roster):
    upload = st.file_uploader('上传需要比对的名单', type=TABLE_TYPES, key='incoming')
    if upload is None:
        return
    try:
        tables = {name: rows for name, rows in read_tables(upload.getvalue(), upload.name).items()
                  if any(any(clean(cell) for cell in row) for row in rows)}
    except Exception as exc:
        st.error(f'无法读取名单。{exc}')
        return
    if not tables:
        st.info('文件里没有可读取的名单。')
        return
    sheet = st.selectbox('工作表', list(tables))
    rows = tables[sheet]
    st.caption('如表头前面有标题或说明，请选择实际表头所在的行。')
    head = int(st.number_input('表头行号', min_value=1, max_value=len(rows), value=header_index(rows)+1,
                               key=f'header-{upload.file_id}-{sheet}'))-1
    width = max(map(len, rows))
    labels = columns(rows[head], width)
    options = [None, *range(width)]
    name_side, id_side = st.columns(2)
    selections = []
    for side, title, aliases in [(name_side, '姓名列', NAME_HEADERS), (id_side, '学号列', ID_HEADERS)]:
        default = next((i+1 for i, value in enumerate(rows[head]) if clean(value).lower() in aliases), 0)
        selections.append(side.selectbox(title, options, index=default,
                                         format_func=lambda i: '未选择' if i is None else labels[i],
                                         key=f'{upload.file_id}-{sheet}-{head}-{title}'))
    if selections == [None, None]:
        st.info('至少选择姓名列或学号列。')
        return
    if selections[0] is not None and selections[0] == selections[1]:
        st.warning('请为姓名和学号选择不同的列。')
        return
    show_results(labels, compare_rows(roster, rows[head+1:], width, *selections), '表格')


def recognize(data, layout):
    executable = os.environ.get('TESSERACT_CMD') or shutil.which('tesseract')
    if not executable:
        raise RuntimeError('找不到 Tesseract。请按 README 安装后重试。')
    with Image.open(BytesIO(data)) as source:
        image = ImageOps.exif_transpose(source).convert('L')
        if image.width < 1200:
            scale = min(1200 / image.width, 2)
            image = image.resize((round(image.width*scale), round(image.height*scale)))
        image = ImageOps.autocontrast(image)
        buffer = BytesIO()
        image.save(buffer, format='PNG')
    command = [executable, 'stdin', 'stdout', '-l', 'chi_sim+eng', '--psm', layout]
    result = subprocess.run(command, input=buffer.getvalue(), capture_output=True, timeout=45)
    if result.returncode:
        raise RuntimeError('图片识别失败。请确认 Tesseract 已安装 chi_sim 和 eng 语言包。')
    return result.stdout.decode('utf-8').strip()


def photo_compare(roster):
    image = st.file_uploader('上传名单照片或截图', type=['png', 'jpg', 'jpeg', 'webp'], key='photo')
    if image is not None:
        st.image(image.getvalue(), width=420)
    layout = st.radio('图片排版', ['整齐的表格或名单', '分散的文字'], horizontal=True)
    if st.button('识别图片', disabled=image is None):
        try:
            with st.spinner('正在本地识别…'):
                st.session_state['photo_text'] = recognize(image.getvalue(), '6' if layout.startswith('整齐') else '11')
                st.session_state.pop('confirmed_photo', None)
                if not st.session_state['photo_text']:
                    st.warning('这张图片未识别出文字，可以手动录入或换一张清晰的图片。')
        except (OSError, RuntimeError, subprocess.TimeoutExpired, ValueError) as exc:
            st.error(f'识别未完成。{exc}')
    with st.form('photo_review'):
        text = st.text_area('核对并编辑识别文本', key='photo_text', height=240,
                            placeholder='每行一人。可填写姓名，或填写学号和姓名。')
        st.caption('确认姓名和学号在同一行。更正时可用 Tab 分隔学号和姓名，只有姓名也可以。')
        submitted = st.form_submit_button('确认并比对')
    if submitted:
        st.session_state['confirmed_photo'] = text
    if 'confirmed_photo' in st.session_state:
        rows = photo_rows(st.session_state['confirmed_photo'])
        show_results(['学号', '姓名'], compare_rows(roster, rows, 2, 1, 0), '照片')


st.set_page_config(page_title='学生名单比对', page_icon='📋', layout='wide')
st.title('学生名单比对')
st.write('查询你负责的学生，或从新名单中筛出对应记录。')
with st.sidebar:
    st.header('我的学生底库')
    path = LOCAL_ROSTER if LOCAL_ROSTER.exists() else EXISTING_ROSTER
    upload = st.file_uploader('导入或更新底库', type=TABLE_TYPES, key='roster')
    try:
        roster = Roster.from_file(upload.getvalue(), upload.name) if upload else (
            Roster.from_file(path.read_bytes(), path.name) if path.exists() else None)
    except Exception as exc:
        roster = None
        st.error(f'底库读取失败。{exc}')
    if roster is not None:
        st.metric('底库记录', len(roster.people))
        st.caption(f'当前使用 {upload.name if upload else path.name}')
        if upload and st.button('保存为默认底库'):
            try:
                content = result_workbook(['学号', '姓名'], {'学生底库': roster.people})
                temporary = LOCAL_ROSTER.with_suffix('.tmp')
                temporary.write_bytes(content)
                temporary.replace(LOCAL_ROSTER)
                st.success('已保存，下次打开会自动加载。')
            except OSError as exc:
                st.error(f'保存失败。{exc}')
if roster is None:
    st.info('请先在左侧导入包含“学号”和“姓名”两列的底库。')
    st.stop()

search, table, photo = st.tabs(['单人查询', '表格比对', '照片识别'])
with search:
    query = st.text_input('输入完整姓名或学号')
    if clean(query):
        people = roster.lookup(query)
        if people:
            st.dataframe(pd.DataFrame(people, columns=['学号', '姓名']), hide_index=True, width='stretch')
        else:
            st.info('底库中没有找到对应记录。')
with table:
    table_compare(roster)
with photo:
    photo_compare(roster)
