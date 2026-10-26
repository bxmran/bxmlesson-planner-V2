# 学校教案49行固定填空对照表（最终定稿版）

只修改如下填写位置，其他标题/栏目名称/表格结构原封不动。行号为程序的0起始索引，用户不需手动操作。实际位置和代码映射见 `scripts/render_school_docx.py:INDEX`。

| 模板位置 | JSON字段 | 内容要求 |
|---|---|---|
| 一、教学基本信息—授课章节 | `chapter_short` | 章/项目＋节/主题简写，不粘贴目录全文 |
| 授课形式 | `form` | 仅当原资料明确给出时填，未知留空 |
| 计划学时 | `hours` | 原计划表数字/单位原样；实践周无学时则留空 |
| 授课地点 | `location` | 没有就空，绝不填上一门课的教室 |
| 二、学情分析3格 | `foundation`, `cognitive_practice`, `learning_characteristics` | 先备知识、实践认知、学习特点；可以合理写预计学习困难但不要假装测量过 |
| 二、教学目标3格 | `knowledge`, `ability`, `literacy` | 知识能力素养；尽量可观察、可评价，课后产出对齐 |
| 二、重难点及措施4格 | `key_points`, `key_solutions`, `hard_points`, `hard_solutions` | 分别写重点及解决措施、难点及解决措施 |
| 三、教学设计思路 | `design_thought.intro/before/during/after` | 一段短总述+课前、课中、课后3条带分项格式 |
| 三、教法学法 | `teaching_methods`, `learning_methods` | **教法3—4条，建议220—270字符**：每条具体到本次课的材料、组织步骤、反馈；**学法2—3条，建议130—170字符**：写学生做什么、形成什么产出、如何复核，不只是列举方法名。教法与学法合计目标约400字，过短或过长均会报错 |
| 三、教学资源、设备、课前任务 | `resources`, `equipment`, `preparation` | **课前任务3—4项，建议130—170字符**，分别写阅读与术语、情境资料准备、预习检测/疑问、成果与课堂利用；不假设教师有某专用平台账号；未证实硬件不冒填 |
| 四、课前1行 | `implementation.before` | 环节、时长、内容、方法、教师活动、学生活动、思政、目标对应 |
| 四、课中5行 | `implementation.during` | **5个相互递进的课中环节**，2学时总共90分钟，不必固定环节名称，应贴合课程 |
| 四、课后1行 | `implementation.after` | 课后巩固、教师批改反馈、学生修订等 |
| 五、考核3行 | `assessment` | 3行依次“评价模块/权重/载体/评价要点/评价主体/反馈方式”，填本次课过程评价 |
| 六、作业 | `homework_text` | 原进度表作业原文完整保留，再补基础、提高、实践等必要分层任务 |
| 七、反思3格 | `reflection_achievement`,`reflection_issues`,`reflection_improvements` | 判断依据、可能难点、改进闭环；不编造真实效果 |

课次标题及授课教师签名处为原模板固定结构，不属于让 AI 生成个人资料的填空。
