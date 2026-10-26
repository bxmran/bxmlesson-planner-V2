#!/usr/bin/env python3
"""Parse the TWO observed layouts of 重庆外语外事学院 official teaching progress table.
No network requests, no reading unrelated directories. Does not invent missing fields.

Output is a faithful JSON extraction for host AI authoring, not finished lesson prose.
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path
from docx import Document


def norm(t: str) -> str:
    return re.sub(r'\u00a0', ' ', (t or '')).strip()


def get_rows(path: Path):
    doc=Document(str(path))
    tables=[t for t in doc.tables if any('课次' in c.text and '课次'==norm(c.text) for r in t.rows[:6] for c in r.cells)]
    if len(tables)!=1:
        raise ValueError(f'无法唯一识别学校教学进度表：找到{len(tables)}张包含“课次”表头的表。请先核对模板。')
    table=tables[0]
    header_idx=next(i for i,row in enumerate(table.rows) if norm(row.cells[0].text)=='课次')
    heads=[norm(c.text) for c in table.rows[header_idx].cells]
    if '学时数' not in heads or not any('授课内容' in h for h in heads) or '作业布置' not in heads:
        raise ValueError('表头不匹配学校模板：必须含课次、学时数、授课内容、作业布置')
    h_hours=heads.index('学时数')
    h_content=next(i for i,h in enumerate(heads) if '授课内容' in h)
    h_hw=heads.index('作业布置')
    h_notes=heads.index('备注') if '备注' in heads else None
    items=[]
    for row in table.rows[header_idx+1:]:
        first=norm(row.cells[0].text)
        if not re.fullmatch(r'\d+',first): continue
        n=int(first)
        items.append({
            'number':n,
            'hours':norm(row.cells[h_hours].text),
            'topic':norm(row.cells[h_content].text),
            'homework':norm(row.cells[h_hw].text),
            'remarks':norm(row.cells[h_notes].text) if h_notes is not None else ''
        })
    if not items: raise ValueError('未找到有效课次编号。')
    numbers=[r['number'] for r in items]
    if len(set(numbers))!=len(numbers): raise ValueError(f'课次编号重复：{numbers}')
    return items


def select(items, spec):
    if not spec:return items
    wanted=set()
    for part in re.split(r'[,，\s]+',spec.strip()):
        if not part: continue
        match=re.fullmatch(r'(\d+)[-—~～至](\d+)',part)
        if match:
            a,b=map(int,match.groups()); wanted.update(range(min(a,b),max(a,b)+1))
        elif part.isdigit(): wanted.add(int(part))
        else: raise ValueError('课次范围应为 1-16 或 1-12,14-17 这样的格式')
    have={i['number'] for i in items}
    missing=sorted(wanted-have)
    if missing:raise ValueError(f'选择的课次不在计划表内：{missing}')
    return [i for i in items if i['number'] in wanted]


def extract(path: Path, only=None):
    items=select(get_rows(path),only)
    # Course metadata deliberately omitted: no teacher/student names or locations.
    return {'source_type':'school_progress_table', 'lessons':items,
            'selected_count':len(items), 'selected_numbers':[i['number'] for i in items]}


def main():
    p=argparse.ArgumentParser(description='按学校官方教学进度表识别全部实际课次。')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True)
    p.add_argument('--only',help='可选；例如 1-16 或 1-12,14-17；不提供即全部课次')
    a=p.parse_args()
    data=extract(a.input,a.only)
    a.out.parent.mkdir(parents=True,exist_ok=True)
    a.out.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print('已识别',len(data['lessons']),'次；编号：',data['selected_numbers'])

if __name__=='__main__': main()
