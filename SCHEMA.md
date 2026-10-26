# 给AI使用的 JSON 结构说明

`draft.json` 必须是 `{"lessons":[...]}`，每个 lesson 对应一张学校原表。代码中的字段有固定映射，不能把英文键翻译成中文键。

```json
{
  "lessons": [
    {
      "number": 1,
      "hours": "2",
      "chapter_short": "项目一 1.1 课程导入",
      "form": "",
      "location": "",
      "foundation": "学生具备基础学习经验……",
      "cognitive_practice": "学生能阅读示例但流程判断需要练习……",
      "learning_characteristics": "适合使用可视化任务卡……",
      "knowledge": "1. 掌握……\n2. 理解……",
      "ability": "1. 能够……",
      "literacy": "1. 强化……",
      "key_points": "核心概念与决策依据……",
      "key_solutions": "1. 采用对照图示……",
      "hard_points": "如何迁移至具体业务情境……",
      "hard_solutions": "1. 分步示范……",
      "design_thought": {"intro":"总述……","before":"准备材料和预习……","during":"讲解、案例、任务……","after":"巩固、反馈、修订……"},
      "teaching_methods": "1. ……",
      "learning_methods": "1. ……",
      "resources": "教材与课堂工作单……",
      "equipment": "多媒体与白板……",
      "preparation": "阅读章节、整理问题……",
      "implementation": {
        "before": {"phase":"自主预习","time":"课前","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标1"},
        "during": [
          {"phase":"情境导入","time":"10分钟","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标1"},
          {"phase":"知识讲解","time":"30分钟","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标1"},
          {"phase":"案例探究","time":"25分钟","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标2"},
          {"phase":"任务实践","time":"20分钟","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标2、目标3"},
          {"phase":"课堂小结","time":"5分钟","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标1、目标3"}
        ],
        "after": {"phase":"拓展提升","time":"课后","content":"……","method":"……","teacher":"……","student":"……","ethics":"……","goals":"目标2"}
      },
      "assessment": [
        {"module":"课前预习","weight":"20%","evidence":"……","criteria":"……","evaluator":"教师、平台","feedback":"……"},
        {"module":"课中表现","weight":"50%","evidence":"……","criteria":"……","evaluator":"教师、小组","feedback":"……"},
        {"module":"课后任务","weight":"30%","evidence":"……","criteria":"……","evaluator":"教师","feedback":"……"}
      ],
      "homework_text": "[必须保留原计划表作业原文]\n……",
      "reflection_achievement": "教学设计以……检验……",
      "reflection_issues": "需重点关注……",
      "reflection_improvements": "拟在下次课……"
    }
  ]
}
```

这是结构示意而非可交付正文，省略号需要由 AI 按真实课次填写。每节不能复用示例文字。

`parse_schedule.py` 产生 `schedule.json`，其每条仅有 `number, hours, topic, homework, remarks`。正式输出时执行 `render_school_docx.py --schedule schedule.json`，从而核实授课次数、原计划学时和作业。**缺少原作业原文会中止生成**。
