import csv
import re
import unicodedata
from collections import defaultdict
from io import BytesIO, StringIO
from pathlib import Path

import openpyxl
import xlrd


NAME_HEADERS = {'姓名', '学生姓名', '名字', 'name'}
ID_HEADERS = {'学号', '学生学号', 'studentid', 'id'}
RESULT_HEADERS = ('比对结果', '底库对应记录')


def clean(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str('' if value is None else value)))


def student_id(value):
    return re.sub(r'(?<=\d)\.0+$', '', clean(value))


def read_tables(data, filename):
    suffix = Path(filename).suffix.lower()
    if suffix == '.csv':
        try:
            text = data.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = data.decode('gb18030')
        return {'CSV': list(csv.reader(StringIO(text)))}
    if suffix == '.xls':
        book = xlrd.open_workbook(file_contents=data)
        return {s.name: [s.row_values(i) for i in range(s.nrows)] for s in book.sheets() if s.nrows}
    if suffix == '.xlsx':
        with BytesIO(data) as buffer:
            book = openpyxl.load_workbook(buffer, read_only=True, data_only=True)
            try:
                return {s.title: list(s.values) for s in book.worksheets if s.max_row is None or s.max_row > 0}
            finally:
                book.close()
    raise ValueError('请选择 XLSX、XLS 或 CSV 文件。')


def header_index(rows):
    for index, row in enumerate(rows[:30]):
        if any(clean(cell).lower() in NAME_HEADERS | ID_HEADERS for cell in row):
            return index
    return 0


def columns(row, width):
    used = set()
    labels = []
    for index in range(width):
        value = row[index] if index < len(row) else None
        label = str(value).strip() if clean(value) else f'第 {index + 1} 列'
        label = label.replace('\n', '')
        if label in used or label in RESULT_HEADERS:
            label = f'{label}（第 {index + 1} 列）'
            while label in used:
                label += '＊'
        used.add(label)
        labels.append(label)
    return labels


def field(row, index):
    return row[index] if index is not None and index < len(row) else None


class Roster:
    def __init__(self, people):
        self.by_id = defaultdict(list)
        self.by_name = defaultdict(list)
        self.people = list(dict.fromkeys(people))
        for sid, name in self.people:
            if sid:
                self.by_id[sid].append((sid, name))
            if name:
                self.by_name[name].append((sid, name))

    @classmethod
    def from_file(cls, data, filename):
        for rows in read_tables(data, filename).values():
            if not rows:
                continue
            head = header_index(rows)
            labels = [clean(x).lower() for x in rows[head]]
            ids = next((i for i, h in enumerate(labels) if h in ID_HEADERS), None)
            names = next((i for i, h in enumerate(labels) if h in NAME_HEADERS), None)
            if ids is not None and names is not None:
                people = [(student_id(field(r, ids)), clean(field(r, names))) for r in rows[head+1:]]
                if any(sid or name for sid, name in people):
                    return cls((sid, name) for sid, name in people if sid or name)
        raise ValueError('底库需要包含“学号”和“姓名”两列，且至少有一条学生记录。')

    def lookup(self, query):
        query = clean(query)
        return list(dict.fromkeys(self.by_id.get(student_id(query), []) + self.by_name.get(query, [])))

    def compare(self, sid, name):
        sid, name = student_id(sid), clean(name)
        id_hits = self.by_id.get(sid, [])
        name_hits = self.by_name.get(name, [])
        if len(id_hits) == 1:
            if name and name != id_hits[0][1]:
                return '待核对：学号与姓名冲突', list(dict.fromkeys(id_hits + name_hits))
            return '学号匹配', id_hits
        if id_hits:
            exact = [person for person in id_hits if name and name == person[1]]
            if len(exact) == 1:
                return '学号与姓名匹配', exact
            return '待核对：底库学号重复', id_hits
        if sid and name_hits:
            return '待核对：姓名存在但学号不同', name_hits
        if len(name_hits) > 1:
            return '待核对：同名', name_hits
        if name_hits:
            return '仅姓名匹配', name_hits
        return '未匹配', []


def compare_rows(roster, rows, width, name_col, id_col):
    results = []
    for row in rows:
        name, sid = field(row, name_col), field(row, id_col)
        if not clean(name) and not student_id(sid):
            continue
        status, hits = roster.compare(sid, name)
        details = '；'.join(f'{n or "姓名空白"} {i or "学号空白"}' for i, n in hits)
        original = [field(row, i) for i in range(width)]
        if id_col is not None:
            original[id_col] = student_id(sid)
        results.append([*original, status, details])
    return results


def result_workbook(labels, groups):
    book = openpyxl.Workbook()
    book.remove(book.active)
    for group, rows in groups.items():
        sheet = book.create_sheet(group)
        for values in [labels, *rows]:
            sheet.append(values)
            for cell in sheet[sheet.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(color='FFFFFF', bold=True)
            cell.fill = openpyxl.styles.PatternFill('solid', fgColor='155C57')
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def photo_rows(text):
    rows = []
    for line in text.splitlines():
        if not line.strip():
            continue
        normalized = unicodedata.normalize('NFKC', line)
        if '\t' in normalized:
            sid, name = normalized.split('\t', 1)
            rows.append((student_id(sid), clean(name)))
            continue
        match = re.search(r'(?<!\d)\d(?: *\d){7,17}(?!\d)', normalized)
        if match:
            sid = student_id(match.group())
            rest = normalized[:match.start()] + normalized[match.end():]
            name = ''.join(re.findall(r'[\u3400-\u9fff·]', rest))
        else:
            sid, name = '', clean(normalized)
        if name not in NAME_HEADERS | ID_HEADERS or sid:
            rows.append((sid, name))
    return rows
