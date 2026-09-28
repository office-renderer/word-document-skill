# word-document-skill

这是一个给 AI 使用的 **Word 文档排版 Skill**。

它不负责写正文，也不负责改内容；它只负责把已有的 `.docx` 按固定模板进行：

- 排版；
- 格式规范化；
- 校验；
- 根据校验结果自动修复。

这个仓库可以被 ChatGPT、Codex、Claude Code 或其他能够访问仓库文件并执行 Python 的 AI Agent 使用。

---

## 1. 这个仓库是什么

它不是 Word 插件，也不是单独的桌面软件。

它本质上由三部分组成：

```text
SKILL.md
    ↓
告诉 AI 应该怎么工作

format_docx.py
    ↓
真正负责修改 Word 格式

validate_docx.py
    ↓
负责检查结果，并输出修复报告
```

模板和规则分别放在：

```text
assets/公文排版Word模板.dotx
config/rules.json
```

因此 AI 不需要自己记“字号是多少、缩进是多少、表格多宽、分页选项怎么设”。

AI 只需要按照 `SKILL.md` 调用脚本。

---

## 2. 使用前提

使用这个 Skill 的 AI 至少需要具备两项能力：

1. 能读取这个 GitHub 仓库；
2. 能在工作目录中执行 Python。

Python 环境需要：

```bash
pip install python-docx lxml
```

当前处理对象是已有的 `.docx` 文件。

正文内容应该在使用本 Skill 之前已经确定好。

---

## 3. 标准工作流程

所有 AI 都使用同一个流程：

```text
已有 DOCX
    ↓
读取 SKILL.md
    ↓
format_docx.py
    ↓
validate_docx.py
    ↓
PASS
    └─ 完成

FAIL
    ↓
生成 validation.json
    ↓
format_docx.py --fix
    ↓
重新 validate
```

首次排版：

```bash
python scripts/format_docx.py 文件.docx --in-place
```

校验并生成报告：

```bash
python scripts/validate_docx.py 文件.docx --report validation.json
```

如果校验失败：

```bash
python scripts/format_docx.py 文件.docx --fix validation.json --in-place
python scripts/validate_docx.py 文件.docx --report validation.json
```

最多自动修复两轮。

如果两轮之后仍然存在同一错误码，应停止重复修复，进入诊断，而不是继续盲目循环。此时优先判断是 formatter 未覆盖该情况，还是 validator 对不可见 OOXML 结构产生了误报。

---

## 4. ChatGPT 怎么使用

前提是 ChatGPT 当前的工作环境能够访问这个仓库和目标 Word 文件。

可以直接告诉 ChatGPT：

```text
请使用 office-renderer/word-document-skill 处理这个 Word 文档。

先读取仓库中的 SKILL.md。
不要修改正文内容，只处理 Word 排版。

按照 Skill 的流程执行：
format → validate → 如果失败则按 validation.json 自动修复 → 再 validate。

最多自动修复两轮。
除非我明确要求检查版式，否则不要转 PDF。
```

如果 ChatGPT 已经位于这个仓库的工作目录中，也可以更简短：

```text
使用当前仓库的 word-document-skill 处理这个 DOCX。
先读 SKILL.md，然后严格按里面的闭环执行。
不要改正文内容。
```

重点是：**让 ChatGPT 先读 `SKILL.md`，不要让它自己重新设计一套 Word 排版规则。**

---

## 5. Codex 怎么使用

最适合的方式是让 Codex 直接在这个仓库中工作。

例如工作目录中有：

```text
word-document-skill/
目标文件.docx
```

然后告诉 Codex：

```text
Read SKILL.md first.

Use this repository only to format 目标文件.docx.
Do not rewrite the document content.

Run the formatter, then the validator.
If validation fails, feed validation.json back to format_docx.py and validate again.
Stop after two repair cycles.
Do not render PDF unless explicitly requested.
```

Codex 不需要自己分析具体 Word 参数。

它只需要负责：

```text
读规则
→ 执行脚本
→ 看 validator 结果
→ 按报告修复
```

---

## 6. Claude 怎么使用

对于 Claude Code 或其他能够直接操作本地仓库的 Claude，使用方式与 Codex 相同。

先把这个仓库放在 Claude 可以访问的工作目录中，然后告诉它：

```text
First read SKILL.md in this repository.

Use this skill only for Word formatting.
Do not change substantive document content.

Process 目标文件.docx using the required loop:
format → validate → fix from validation.json → validate again.

Stop after two failed repair cycles and report the remaining validator codes.
Do not invent new formatting rules.
```

如果使用的 Claude 环境不能执行本地 Python，那么它只能阅读规则，不能完整执行这个 Skill。

---

## 7. 其他 AI Agent 怎么使用

并不要求必须是 GPT、Codex 或 Claude。

任何 Agent 只要满足：

```text
能读取仓库
+
能执行 Python
+
能访问目标 DOCX
```

就可以使用。

统一入口只有一个：

```text
先读 SKILL.md
```

不要让不同 Agent 分别理解 `template-spec.md`、`rules.json` 后自己实现排版。

正常情况下：

- AI 负责调用；
- `format_docx.py` 负责修改；
- `validate_docx.py` 负责检查；
- `rules.json` 负责保存机器规则；
- DOTX 模板负责提供模板格式。

---

## 8. 仓库结构

```text
word-document-skill/
├─ README.md
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

各文件职责：

- `README.md`：给使用者看，说明这个仓库怎么使用；
- `SKILL.md`：给 AI 看，规定实际执行流程；
- `公文排版Word模板.dotx`：模板格式的权威来源；
- `rules.json`：formatter 和 validator 共用的机器规则；
- `format_docx.py`：首次排版和自动修复；
- `validate_docx.py`：只检查，不修改；
- `template-spec.md`：排版异常时用于诊断，不是日常必读文件。

---

## 9. 最重要的原则

不要让 AI 自己记忆和执行大量 Word 格式细节。

正确方式是：

```text
AI 负责调用
Python 负责排版
Validator 负责检查
模板和 rules.json 负责定义规则
```

如果 validator 连续两轮仍然发现同一问题，说明当前程序还没有覆盖这种情况。

此时应该分析问题并更新 Skill 的规则或脚本，而不是让 AI 无限重复尝试。


## Word Online 字体兼容

本仓库不会为了兼容浏览器版 Word 把 `仿宋_GB2312` 替换成 `FangSong`。

实际对照测试表明，问题来自同一个 Word run 内同时包含中文与英文/数字时，Word Online 可能与桌面 Word 采用不同的字体槽解析方式。

因此 formatter 会自动：

```text
中文 run      → 当前段落样式的东亚字体
英文/数字 run → 当前段落样式的西文字体（通常 Times New Roman）
```

例如：

```text
原始一个 run：
2021年以来，北京北矿……

格式化后：
[2021]        Times New Roman
[年以来，北京北矿……]  仿宋_GB2312
```

文字内容不发生改变，只调整 Word 内部 run 边界和字体槽。validator 会检查混合 run 和显式字体设置，并通过 `font` 修复组交回 formatter 自动修复。
