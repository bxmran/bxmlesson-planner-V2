#!/usr/bin/env python3
"""Fill the exact school lesson-only template. Does NOT make a cover or a course homepage.
Requires authored content JSON (the AI host is responsible for drafting).
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn

FIVE=Pt(10.5)
BODY_LINE_HEIGHT=Pt(20)  # Consistent, comfortable spacing in all filled text
ITEM_AFTER=Pt(2)           # Small visual separation for numbered tasks
DEFAULT_TEMPLATE=Path(__file__).resolve().parents[1]/'assets/school_lesson_only_blank.docx'
TIME_ROWS=[33,34,35,36,37]
ACTIVITY_ROWS=[30,33,34,35,36,37,40]
# Names from the actual school template cells; no field guesses.
INDEX={
 'chapter_short':(1,3), 'form':(1,8),'hours':(2,3),'location':(2,8),
 'foundation':(5,3),'cognitive_practice':(6,3),'learning_characteristics':(7,3),
 'knowledge':(9,3),'ability':(10,3),'literacy':(11,3),
 'key_points':(13,3),'key_solutions':(14,3),'hard_points':(15,3),'hard_solutions':(16,3),
 'teaching_methods':(21,3),'learning_methods':(22,3),
 'resources':(24,2),'equipment':(25,2),'preparation':(26,2),
 'homework_text':(44,0), 'reflection_achievement':(46,3),
 'reflection_issues':(47,3),'reflection_improvements':(48,3),
}
ACTIVITY_FIELDS={'phase':0,'time':1,'content':3,'method':4,'teacher':5,'student':7,'ethics':8,'goals':9}
REQUIRED_TEXT_FIELDS=['chapter_short','foundation','cognitive_practice','learning_characteristics','knowledge','ability','literacy',
'key_points','key_solutions','hard_points','hard_solutions','teaching_methods','learning_methods','resources','equipment',
'preparation','homework_text','reflection_achievement','reflection_issues','reflection_improvements']
ASSESS_FIELDS=['module','weight','evidence','criteria','evaluator','feedback']
# Content-quality checks for ordinary teaching sessions. An exam can be concise,
# but a teaching session must not pass by listing bare method names.
MIN_CONTENT_CHARS={'teaching_methods':180,'learning_methods':110,'preparation':110}
MAX_CONTENT_CHARS={'teaching_methods':310,'learning_methods':220,'preparation':190}
METHOD_TOTAL_RANGE=(330,460)  # 教法 + 学法，约400字
MIN_NUMBERED_ITEMS={'teaching_methods':3,'learning_methods':2,'preparation':3}
NUMBERED_ITEM=re.compile(r'(?m)^\s*(?:[1-9]\d*[.、．]|[（(][1-9]\d*[)）])\s*')


def exam_only_lesson(lesson):
    title=str(lesson.get('chapter_short','')).strip().replace(' ','')
    # "能力检测评讲/考核讲评" is still instruction, not a pure exam.
    return bool(re.match(r'^(?:第\d+次课)?(?:期末考试|期中考试|随堂考试|考试|结课考试)(?:$|[（(:：])',title))


def verify_core_pedagogy_detail(lesson):
    if exam_only_lesson(lesson):
        return
    n=lesson['number']
    labels={'teaching_methods':'教法','learning_methods':'学法','preparation':'课前预习任务'}
    lengths={}
    for field, minimum in MIN_CONTENT_CHARS.items():
        content=str(lesson.get(field,'')).strip()
        size=len(flatten_text(content))
        lengths[field]=size
        if size<minimum:
            raise ValueError(f'第{n}次课{labels[field]}只有{size}字，需至少{minimum}字；须写出本课具体任务，不能只列方法名称')
        maximum=MAX_CONTENT_CHARS[field]
        if size>maximum:
            raise ValueError(f'第{n}次课{labels[field]}有{size}字，超过精练上限{maximum}字，请删减重复细节而不删教学要点')
        count=len(NUMBERED_ITEM.findall(content))
        if count<MIN_NUMBERED_ITEMS[field]:
            raise ValueError(f'第{n}次课{labels[field]}需要至少{MIN_NUMBERED_ITEMS[field]}项编号的针对性任务，实际{count}项')
    total=lengths['teaching_methods']+lengths['learning_methods']
    lo,hi=METHOD_TOTAL_RANGE
    if not lo<=total<=hi:
        raise ValueError(f'第{n}次课教法＋学法共{total}字，须控制在{lo}—{hi}字（目标约400字）')


def flatten_text(p):
    return re.sub(r'\s+', '',str(p or ''))


def write_cell(cell, content, *, bold=False, align=None, leading=False):
    # Target a fillable cell only. Never overwrite fixed school headings or table labels.
    content='' if content is None else str(content)
    cell.text=''
    lines=content.split('\n')
    for i,line in enumerate(lines):
        p=cell.paragraphs[0] if i==0 else cell.add_paragraph()
        # One consistent, slightly airy rhythm across teaching analysis,
        # methodology, implementation, assessment, homework and reflection.
        # Leave school-owned headings and label cells untouched.
        p.paragraph_format.space_before=Pt(0)
        has_item_marker=bool(NUMBERED_ITEM.match(line) or line.startswith('• '))
        p.paragraph_format.space_after=ITEM_AFTER if has_item_marker and i<len(lines)-1 else Pt(0)
        p.paragraph_format.line_spacing=BODY_LINE_HEIGHT
        if align is not None:p.alignment=align
        if not line:continue
        run=p.add_run(line)
        run.font.name='宋体'
        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
        run.font.size=FIVE
        run.bold=bold
        run.font.highlight_color=None


def write_activity(table, row_i, activity):
    for key,col in ACTIVITY_FIELDS.items():
        val=activity.get(key,'')
        write_cell(table.rows[row_i].cells[col],val,
                   bold=(key=='phase'),
                   align=WD_ALIGN_PARAGRAPH.CENTER if key in ('phase','time','goals') else WD_ALIGN_PARAGRAPH.LEFT,
                   leading=True)


def validate_draft(data):
    lessons=data.get('lessons')
    if not isinstance(lessons,list) or not lessons:raise ValueError('JSON 必须包含非空 lessons 数组')
    nums=[]
    for lesson in lessons:
        n=lesson.get('number')
        if not isinstance(n,int) or n<1: raise ValueError(f'课次编号无效：{n}')
        nums.append(n)
        if any(not isinstance(lesson.get(k),str) or not lesson[k].strip() for k in REQUIRED_TEXT_FIELDS):
            missed=[k for k in REQUIRED_TEXT_FIELDS if not isinstance(lesson.get(k),str) or not lesson[k].strip()]
            raise ValueError(f'第{n}次课缺少必须撰写的内容：{missed}')
        verify_core_pedagogy_detail(lesson)
        imp=lesson.get('implementation') or {}
        if not isinstance(imp.get('before'),dict) or not isinstance(imp.get('after'),dict):
            raise ValueError(f'第{n}次课缺少课前或课后教学活动')
        during=imp.get('during')
        if not isinstance(during,list) or len(during)!=5:raise ValueError(f'第{n}次课须提供5个课堂实施环节')
        for stage in [imp['before']]+during+[imp['after']]:
            missing=[k for k in ACTIVITY_FIELDS if not isinstance(stage.get(k),str)]
            if missing:raise ValueError(f'第{n}次课教学实施内容缺失：{missing}')
        hours_string=str(lesson.get('hours','')).strip()
        if not hours_string:
            # Do not fabricate minute values for an unscheduled practice week.
            if any(re.search(r'\d+\s*分钟',stage['time']) for stage in during):
                raise ValueError(f'第{n}次课无计划学时，不应自行编造课堂分钟数')
        elif re.fullmatch(r'[1-9]\d*(?:\s*学时)?',hours_string):
            h=int(re.match(r'\d+',hours_string).group())
            if h!=2:
                mvals=[]
                for stage in during:
                    mt=re.fullmatch(r'\s*(\d+)\s*分钟\s*',stage['time'])
                    if not mt:raise ValueError(f'第{n}次课应写明五个教学环节的分钟数')
                    mvals.append(int(mt.group(1)))
                if sum(mvals)!=45*h:raise ValueError(f'第{n}次课课中时间应合计{45*h}分钟，实际为{sum(mvals)}')
        ass=lesson.get('assessment')
        if not isinstance(ass,list) or len(ass)!=3:raise ValueError(f'第{n}次课须有3个过程性评价模块')
        for item in ass:
            if any(not isinstance(item.get(k),str) for k in ASSESS_FIELDS):
                raise ValueError(f'第{n}次课评价模块缺失内容')
        values=[]
        for item in ass:
            mt=re.fullmatch(r'\s*(\d+)\s*%\s*',item.get('weight',''))
            if mt:values.append(int(mt.group(1)))
        if len(values)==3 and sum(values)!=100:
            raise ValueError(f'第{n}次课过程性评价权重合计{sum(values)}%，应为100%')
        strat=lesson.get('design_thought')
        if not isinstance(strat,dict) or any(not isinstance(strat.get(k),str) or not strat[k].strip() for k in ('intro','before','during','after')):
            raise ValueError(f'第{n}次课教学设计思路须包含总述、课前、课中、课后')
        # Only check duration for conventional 2-hour meetings; practice week without a stated duration must not be invented.
        if str(lesson.get('hours','')).strip() in ('2','2学时'):
            durations=[]
            for stage in during:
                tm=re.fullmatch(r'\s*(\d+)\s*分钟\s*',stage['time'])
                if not tm:raise ValueError(f'第{n}次课课中时间应写如“10分钟”')
                durations.append(int(tm.group(1)))
            if sum(durations)!=90:raise ValueError(f'第{n}次课课中时间总和{sum(durations)}分钟，不是90分钟')
    if len(set(nums))!=len(nums):raise ValueError(f'重复课次：{nums}')
    # Budget the content authored by the model. Avoid counting fixed column headings as writing volume.
    # Usable Chinese character count is a simplified non-whitespace count.
    total=0
    for lesson in lessons:
        def collect(x):
            if isinstance(x,dict):return ''.join(collect(v) for k,v in x.items() if k not in ('form','location'))
            if isinstance(x,list):return ''.join(collect(v) for v in x)
            if isinstance(x,str):return x
            return ''
        nchar=len(flatten_text(collect(lesson)))
        total+=nchar
    ceiling=85000*len(lessons)//16
    if total>ceiling:
        raise ValueError(f'正文填写字数约{total}，超过{len(lessons)}次课按比例计算上限{ceiling}。请精简后重新生成。')
    return {'count':len(lessons),'numbers':nums,'estimated_fill_chars':total,'ceiling':ceiling}


def fill_table(table, lesson):
    for key,(row_i,col_i) in INDEX.items():
        val=lesson.get(key,'')
        # No teacher name, course cover, personal identifying info, or guessed room info.
        write_cell(table.rows[row_i].cells[col_i],val)
    s=lesson['design_thought']
    design='\n'.join([s['intro'],'• 课前：'+s['before'],'• 课中：'+s['during'],'• 课后：'+s['after']])
    write_cell(table.rows[19].cells[0],design)
    stages=lesson['implementation']
    write_activity(table,30,stages['before'])
    for i,v in zip(TIME_ROWS,stages['during']):write_activity(table,i,v)
    write_activity(table,40,stages['after'])
    score=table.rows[42].cells[0].tables[0]
    assert len(score.columns)==6 and len(score.rows)==4
    for c in score.rows[0].cells:
        # Existing template label retained exactly, only table header style standardized per user instruction.
        for p in c.paragraphs:
            for r in p.runs:
                if r.text:r.font.size=FIVE;r.bold=True
    for r_i,rowdata in enumerate(lesson['assessment'],start=1):
        for j,key in enumerate(ASSESS_FIELDS):
            write_cell(score.rows[r_i].cells[j],rowdata.get(key,''),
                       align=WD_ALIGN_PARAGRAPH.CENTER if j in (0,1,4) else WD_ALIGN_PARAGRAPH.LEFT)


def check_schedule_fidelity(draft, schedule):
    """Fail closed if lessons, stated hours or assigned work mismatch the uploaded schedule."""
    source = {x['number']: x for x in schedule['lessons']}
    result = {x['number']: x for x in draft['lessons']}
    if list(result) != list(source):
        raise ValueError(f"课次不匹配：计划表={list(source)}，教案={list(result)}")
    for n, row in source.items():
        lesson=result[n]
        stated=str(row.get('hours','')).strip()
        if str(lesson.get('hours','')).strip()!=stated:
            raise ValueError(f'第{n}次课学时与进度表不一致：{stated} != {lesson.get("hours","")}')
        source_hw=flatten_text(row.get('homework',''))
        if source_hw and source_hw not in flatten_text(lesson.get('homework_text','')):
            raise ValueError(f'第{n}次课缺少原计划表布置的作业原文，请将其原样收入 homework_text')
        if not row.get('topic','').strip():
            raise ValueError(f'第{n}次课原进度表授课内容为空，请人工核实')


def generate(input_json:Path, output_docx:Path, template_path:Path=DEFAULT_TEMPLATE, schedule_json:Path|None=None):
    data=json.loads(input_json.read_text(encoding='utf-8'))
    report=validate_draft(data)
    if schedule_json is not None:
        schedule=json.loads(schedule_json.read_text(encoding='utf-8'))
        check_schedule_fidelity(data,schedule)
        report['schedule_fidelity']='PASS'
    original=Document(str(template_path))
    if len(original.tables)!=1 or len(original.tables[0].rows)!=49:raise ValueError('不是本技能内置的49行学校课次模板')
    body=original._element.body
    # Keep document styles and section geometry; regenerate content only as complete lesson blocks.
    for node in list(body):
        if node.tag!=qn('w:sectPr'):body.remove(node)
    for ix,lesson in enumerate(data['lessons']):
        part=Document(str(template_path))
        p=part.paragraphs[0]
        p.clear(); run=p.add_run(f"第{lesson['number']}次课")
        run.font.name='宋体';run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
        run.bold=True;run.font.size=Pt(22)
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        fill_table(part.tables[0],lesson)
        if ix:
            blank=original.add_paragraph()
            blank.add_run().add_break(WD_BREAK.PAGE)
        # Insert before sectPr to keep OOXML body order Word-compliant.
        sect=body[-1] if len(body) and body[-1].tag==qn('w:sectPr') else None
        ins=(lambda el: body.insert(len(body)-1,el)) if sect is not None else body.append
        ins(deepcopy(p._p))
        ins(deepcopy(part.tables[0]._tbl))
        sig=part.paragraphs[1]
        ins(deepcopy(sig._p))
    original.core_properties.author=''
    original.core_properties.last_modified_by=''
    output_docx.parent.mkdir(parents=True,exist_ok=True)
    original.save(str(output_docx))
    # Confirm all lessons included, no homepage and no names invented.
    final=Document(str(output_docx))
    if len(final.tables)!=len(data['lessons']):raise AssertionError('丢失课次表格')
    if any('教案（首页）' in p.text for p in final.paragraphs):raise AssertionError('错误生成了课程首页')
    report['output']=str(output_docx)
    return report


def main():
    ap=argparse.ArgumentParser(description='使用学校原定稿版式，只输出第X次课教案')
    ap.add_argument('input',type=Path);ap.add_argument('--out',type=Path,required=True)
    ap.add_argument('--template',type=Path,default=DEFAULT_TEMPLATE)
    ap.add_argument('--schedule',type=Path,help='parse_schedule.py 输出的计划表JSON，启用课次/学时/作业一致性校验')
    args=ap.parse_args()
    print(json.dumps(generate(args.input,args.out,args.template,args.schedule),ensure_ascii=False,indent=2))

if __name__=='__main__':main()
