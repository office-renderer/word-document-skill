# word-document-skill

这是一个给 ChatGPT、Codex、Claude Code 等 AI Agent 使用的 **Word 排版 Skill**。

它只负责把已有 DOCX 按固定模板进行排版、校验和自动修复，**不负责撰写、改写或补充正文内容**。

## 最简单的用法

AI 进入本仓库后先读 `SKILL.md`，然后运行：

```bash
python scripts/run_skill.py 文件.docx
```

这个入口会自动完成：

```text
首次排版
  ↓
快速校验
  ↓
有可修问题 → 定向修复
  ↓
重新校验
  ↓
PASS 或达到自动修复上限后停止
```

自动修复次数由 `config/rules.json` 控制，AI 不需要自己数轮次。

最终结构检查可用：

```bash
python scripts/run_skill.py 文件.docx --strict
```

PDF 渲染不属于日常校验，只有明确需要版式检查时才执行。

## 给不同 Agent 的提示词

ChatGPT / Codex / Claude Code 都可以直接使用：

```text
先读取 SKILL.md。
只处理这个 DOCX 的 Word 排版，不修改正文内容。
运行 scripts/run_skill.py。
如果自动修复停止并仍有 validator 错误码，报告错误码，不要自行设计新的修复流程。
```

前提是 Agent 能读取仓库、访问目标 DOCX，并能执行 Python。

依赖：

```bash
pip install python-docx lxml
```

## 仓库结构

```text
word-document-skill/
├─ README.md
├─ SKILL.md
├─ agents/openai.yaml
├─ assets/公文排版Word模板.dotx
├─ config/rules.json
├─ references/template-spec.md
└─ scripts/
   ├─ run_skill.py
   ├─ format_docx.py
   └─ validate_docx.py
```

职责很简单：

- 模板：定义版式；
- `rules.json`：保存机器规则；
- `format_docx.py`：应用格式和定向修复；
- `validate_docx.py`：只检查；
- `run_skill.py`：执行完整闭环；
- `template-spec.md`：只有排查异常或维护 Skill 时才看。

## 已验证的 Word Online 字体处理

桌面 Word 能在同一个 run 中按字符自动选择中文和西文字体，但 Word Online / OneDrive 曾出现解析不一致。

最终验证有效的处理是：

```text
中文 run      → 保留模板中文字体，例如 仿宋_GB2312
英文/数字 run → 保留模板西文字体，例如 Times New Roman
```

formatter 会拆分普通的中西文混合 run，并显式写入字体槽，同时保留原 run 已经存在的直接字体设置。

单纯把 `仿宋_GB2312` 改成 `FangSong`、修改字体表或只设置 `themeFontLang` 都不能替代这一处理。

## 设计原则

AI 负责调用，Python 负责排版，validator 负责验收，模板负责定义样式。

出现两轮仍无法修复的问题时，应先判断是 formatter 未覆盖，还是 validator 误报，再更新 Skill；不要无限重复同一修复。
