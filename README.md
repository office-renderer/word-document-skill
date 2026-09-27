# word-document-skill

这是一个用于 **Word 文档排版、规范化、校验和自动修复** 的独立 Skill 仓库。

它的职责很明确：**只处理 Word 格式，不负责生成、改写或补充正文内容。**

核心目标是让 AI 不再自己记忆大量 Word 排版细节，而是通过固定脚本完成：

```text
已有内容 / Word 文档
        ↓
format_docx.py
        ↓
validate_docx.py
        ↓
PASS → 完成

FAIL
        ↓
生成 validation.json
        ↓
format_docx.py --fix
        ↓
重新 validate
```

如果连续两轮自动修复后仍然存在同一问题，应停止重复修复，进入诊断处理，而不是继续盲目循环。

## 仓库结构

```text
word-document-skill/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ assets/
│  └─ 公文排版Word模板.dotx
├─ config/
│  └─ rules.json
├─ references/
│  └─ template-spec.md
└─ scripts/
   ├─ format_docx.py
   └─ validate_docx.py
```

其中：

- `assets/公文排版Word模板.dotx`：排版模板，是模板格式的权威来源。
- `config/rules.json`：机器执行的排版规则，formatter 和 validator 共用。
- `scripts/format_docx.py`：负责首次排版和按校验报告自动修复。
- `scripts/validate_docx.py`：只负责检查，不修改文档。
- `references/template-spec.md`：模板说明，主要用于排版诊断和维护。
- `SKILL.md`：规定整个 Skill 的调用流程和边界。

## 基本用法

首次排版：

```bash
python scripts/format_docx.py output.docx --in-place
```

校验并生成报告：

```bash
python scripts/validate_docx.py output.docx --report validation.json
```

如果校验失败，可根据报告自动修复：

```bash
python scripts/format_docx.py output.docx --fix validation.json --in-place
python scripts/validate_docx.py output.docx --report validation.json
```

最终结构检查或排查异常时，可使用：

```bash
python scripts/validate_docx.py output.docx --strict
```

## 使用原则

- 不修改正文实质内容。
- 不在 AI 提示词中重复维护大量格式规则。
- 格式规则统一放在 `config/rules.json`。
- validator 只检查，不修改。
- formatter 负责首次排版和定向修复。
- Word→PDF 不属于日常校验，只在明确需要版式检查时执行。
- 自动修复连续两轮仍未解决的问题，应转入诊断，而不是继续循环。
