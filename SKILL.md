---
name: bxmlesson-planner-v2
version: 2.3.0
summary: 上传教学进度表，自动生成指定或全部课次的完整 Word 教案。
description: 上传教学进度表，自动生成指定或全部课次的完整 Word 教案。
---

# bxmlesson-planner V2

**功能：** 上传教学进度表，生成指定或全部课次的完整、可编辑 Word 教案。

**使用方法：** 上传计划表，说“生成第1—16次课”或“按计划表全部生成”。

**内部执行（每次必须执行，无需向用户展示细节）：**
1. 先完整读取并严格执行 `references/DETAILED_RULES.md`、`references/AUTHORING.md`、`references/FIELDS.md`、`references/SCHEMA.md` 和 `references/QUALITY_EXAMPLES.md`。
2. 用 `scripts/parse_schedule.py` 读取进度表，逐次撰写符合字段规范的结构化内容。
3. 多课次任务内部按每批2—4次课撰写并合并为完整 `draft.json`，不要向用户拆分交付。
4. 用 `scripts/render_school_docx.py` 填入 `assets/school_lesson_only_blank.docx`；程序必须逐课检查可见正文字符数，普通授课课次不足4,500字符时补写并重新校验，全部通过后才输出完整 Word。

**原则：** 只做课次教案；不加封面，不补猜测信息，不重新绘制模板，不固定课次数量。需要脚本执行能力，无法运行时须如实说明。
