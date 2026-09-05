---
name: study-fast
description: Create a two-part evidence-based learning system from course, job, project, research-progress, or meeting materials. The study handbook combines knowledge with likely questions and answers; the teacher narrative patiently explains the subject in connected prose. Use for rapid learning, exams, interviews, onboarding, meetings, and defenses. Do not use for ordinary document editing or standalone debugging.
---

# Study Fast

把当前材料转化为可核验、可复习、可表达的准备包。默认生成“带问题及答案的学习手册”和“老师视角的耐心讲解稿”；只有真实的代码、数据、实验、调试或技术设计需求才生成实操指南。

## 确认任务边界

先判断用户要求的是聊天式分析、材料审查、问答演练、模拟会议/面试，还是生成文档。只要求分析或审查时不得创建文件。不得代替用户提交考试、参加面试、对外发送材料、联系第三方或发布未授权信息。

## 选择主场景

选择一个主场景，必要时记录一个辅场景：

- `knowledge_point`：知识点、方法、工具或论文主题速学。
- `exam_prep`：课程、教材、题型或考试范围复习。
- `job_interview`：JD 匹配、项目表达、行为与技术面试。
- `work_onboarding`：职责、业务、流程、系统或代码库上手。
- `weekly_progress`：课题、实验、研究或项目周进展汇报。
- `meeting_prep`：组会、答辩、评审、客户会或内部会议。

只读取 [scenario-playbooks.md](references/scenario-playbooks.md) 中所选场景的一节。能从材料安全推断时继续并标注推断；只有输出性质会因此改变时才向用户确认。

## 选择输出

- 完整准备包使用 `selection.output_mode=full_pack`：生成 `study_handbook` 和 `teaching_narrative`；`practice_needed=true` 时再生成 `practice_guide`。
- 用户明确指定文档时使用 `selection.output_mode=selected`，并在 `requested_documents` 中列出所需类型。不得为了凑数量生成空洞内容。
- 仅当明确存在代码面试、SQL/数据任务、现场编程、系统设计、调试、实验复现、技术演示或代码走读时设置 `practice_needed=true`。

详细内容边界见 [output-contract.md](references/output-contract.md)。生成两份核心文档时还要读取 [learning-design.md](references/learning-design.md)，按可观察目标和用户可用时间设计学习闭环。

## 读取材料并建立证据

先只读盘点文件类型、大小、相关性和访问限制，再按价值读取。Office、PDF、表格、幻灯片或代码仓库使用对应技能或可靠解析器；不要粗暴剥离 Office XML，也不要无差别扫描大型目录。

对进入文档的事实建立 `evidence` 账本：

- `VERIFIED`：由当前材料、代码、日志、结果或可靠来源直接支持。
- `USER_STATED`：用户陈述但尚未独立核验。
- `INFERRED`：基于材料的合理推断。
- `UNKNOWN`：缺失、冲突、不可访问或无法确认。

需要在文档中展示账本时使用 `evidence_table` 块并引用证据 ID。不得编造成绩、贡献、指标、研究结论、面试流程、职责或会议决定；数值保留定义、基线、时间/版本和来源。

使用个人或未公开材料时完整读取 [privacy-and-evidence.md](references/privacy-and-evidence.md)。

## 编写内容

围绕用户最终需要完成的行为组织内容：理解、回忆、作答、演示、决策或汇报。

- `study_handbook` 用于检索、知识地图、行动和应对提问。必须包含与当前场景相关的可能问题，以及每道问题的答案、参考作答或证据受限的回答框架；不能只列问题。
- `teaching_narrative` 把学习者放在学生位置，由一位耐心、经验丰富的老师或直属上级从头讲解。以自然连续的段落为主，解释为什么、是什么、如何运作、哪里容易出错以及怎样理解和表达；只在确实有助于理解时使用少量分点和表格，不写成第二份提纲或题库。
- `practice_guide` 只用于可运行或可复现的实操。

事实、建议、示例和未知项必须分开。题目、案例和练习随当前材料变化，不得塞入固定算法、岗位、课程、技术栈或会议话术。

## 配置和生成

以 [pack-config.example.json](assets/pack-config.example.json) 为起点；编辑器可使用 [pack-config.schema.json](assets/pack-config.schema.json)。任务配置、生成结果和渲染文件必须放在技能目录之外。

模板：

- `formal_zh_cn`：宋体和 Times New Roman，适合中文 Windows/Word 环境。
- `formal_cross_platform`：Noto Serif CJK SC 和 Liberation Serif，适合跨平台环境；仍需通过渲染确认实际字体可用。

先校验，再生成：

```text
python scripts/build_pack.py --config path/to/pack.json --check-only
python scripts/build_pack.py --config path/to/pack.json --output-dir path/to/output
```

使用 `--list-templates` 查看模板。默认拒绝覆盖现有文件；只有用户明确需要时才使用 `--overwrite`。

## 验证

对每份最终 DOCX：

1. 确认文件非空并能由 `python-docx` 重新打开。
2. 运行 `scripts/audit_docx.py` 检查页面、字体、颜色、元数据和隐私。
3. 使用 documents 技能的 `render_docx.py` 渲染逐页 PNG，并检查每一页的裁切、重叠、空白页、标题、表格、字体替换和页眉页脚。
4. 对照材料复核事实、证据、引用、代码、命令、数据和结论；修改后重新审计与渲染。

缺少 LibreOffice 或所需字体时，只能报告结构检查结果，并明确视觉验证未完成。

## 交付

说明读取的材料、主场景、生成文件、实操指南生成或跳过的原因、主要证据、未知缺口、隐私候选和实际验证。只交付用户要求的最终文档，不交付配置、渲染页或临时审计文件，除非用户明确要求。
