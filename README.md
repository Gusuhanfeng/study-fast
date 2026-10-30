# Study Fast

教材式学习手册、连贯知识讲解与按需的题库、模拟卷和实操训练。

Evidence-grounded learning handbooks, connected explanatory readers, optional question banks, mock exams and practical interview exercises.

## Outputs

| Type | Purpose |
| --- | --- |
| study_handbook | 在约定范围内独立完整的教材 / self-contained textbook |
| teaching_narrative | 少标题、连贯段落的知识讲解 / connected explanatory prose |
| question_bank | 按题型或能力组织的题目与答案 / questions and answers by type |
| mock_exam | 分值校验、集中答案的模拟卷 / scored mock exam with answer section |
| practice_guide | 代码、SQL、实验或系统设计 / executable or reproducible practice |

`full_pack` preserves the original two core documents, plus practice when needed. `selected` supports any explicitly requested combination. Theory question banks do not require `practice_needed=true`.

## Usage

```text
python -m pip install -r requirements.txt
python scripts/build_pack.py --config examples/question-bank.json --check-only
python scripts/build_pack.py --config examples/question-bank.json --output-dir ../study-output
python scripts/audit_docx.py ../study-output/question-bank.docx --config examples/question-bank.json
python scripts/self_check.py
python -m unittest discover -s tests -v
```

See [skill instructions](SKILL.md), [output contract](references/output-contract.md), [assessment design](references/assessment-design.md), [layout](references/formal-document-template.md), and [schema](assets/pack-config.schema.json). Core, question-bank and mock-exam example configurations are bundled. They demonstrate the interface, not complete course content.

## Layout and compatibility

Fixed A4, Chinese/Latin font mapping, two-character body indent, native Word equations, three-line tables and no running header. Question banks default to academic blue questions and black answers. Per-document layout supports header/cover/contents switches, black questions, legacy grid tables, and adjacent or separate answers. Old configurations remain readable; visual defaults intentionally change. Pass the same config to the auditor for explicit overrides.

Equations accept math-only OMML; rich questions and answers support paragraphs, tables, equations and code. Every question requires an answer. Mock exam points must match the declared total. Internal evidence and inferred-question provenance are not automatically printed.

## Verification and privacy

Read the latest artifact before editing; never overwrite manual changes using a stale generator. Outputs and private configurations stay outside the repository. Native Word or another available renderer may produce page images for complete visual review; structural tests do not prove pagination, mathematical validity or learning completeness. Do not upload private course, job or project sources without authorization.

Contributions: [CONTRIBUTING.md](CONTRIBUTING.md). Security: [SECURITY.md](SECURITY.md). License: [MIT](LICENSE).
