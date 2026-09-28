# word-document-skill

这是一个给 ChatGPT、Codex、Claude Code 等 AI Agent 使用的 **Word 排版 Skill**。它只处理已有 DOCX 的排版、校验和修复，不负责撰写或改写正文。

## 最简单的用法

```bash
python scripts/run_skill.py 文件.docx
```

现在采用 **validate-first**：

```text
先校验
  ↓
已经合格 → 直接结束，不重写 DOCX

发现问题
  ↓
只执行 validation.json 对应的修复组
  ↓
重新校验
  ↓
PASS 或达到自动修复上限
```

这对 200～300 页的大文档尤其重要：已经正确的文档不会为了“再格式化一次”而完整保存一遍。

最终结构检查：

```bash
python scripts/run_skill.py 文件.docx --strict
```

依赖：

```bash
pip install -r requirements.txt
```

## 给 AI Agent 的提示词

```text
先读取 SKILL.md。
只处理 DOCX 排版，不修改正文内容。
运行 scripts/run_skill.py。
保留已有横向/竖向分节。
如果自动修复停止并仍有 validator 错误码，报告错误码，不要自行设计新的修复流程。
```

## 三个重要处理原则

### 1. 横向大表按所在 section 处理

程序不再把所有 section 强制改成竖向。已有横向 section 会保留横向，页面尺寸按 A4 横向规范化；页边距等仍按模板执行。

表格宽度不再使用全篇统一的 8845 twips 上限，而是读取表格实际所在 section：

```text
竖向 section → 使用竖向版心宽度
横向 section → 使用横向版心宽度
```

只有超出所在 section 版心的表格才缩小，窄表不强制撑满。

### 2. 大文档先校验再修改

runner 不再一上来执行完整 formatter。先运行只读 validator，只有发现具体问题后才调用对应修复组。

`--strict` 也只在快速校验通过以后执行一次，避免大型文档在每轮修复时重复做结构深检。

### 3. 中西文分 Run 最小化

只处理真正同时包含：

- 中文/东亚文字；
- ASCII 英文字母或数字；

的普通文本 run。

空格、括号、百分号、句号等中性字符不会单独触发拆分，而是附着到相邻文字片段。纯中文、纯英文/数字 run 不拆，也不会为了“规范化”而额外写入字体槽。

对真正拆分的 mixed run，保留原有直接字体；缺失字体从段落样式和 `docDefaults` 继承。

## 仓库结构

```text
word-document-skill/
├─ README.md
├─ SKILL.md
├─ requirements.txt
├─ agents/openai.yaml
├─ assets/公文排版Word模板.dotx
├─ config/rules.json
├─ references/template-spec.md
└─ scripts/
   ├─ run_skill.py
   ├─ format_docx.py
   └─ validate_docx.py
```

## Word Online 字体经验

已验证：单纯把 `仿宋_GB2312` 改成 `FangSong`、修改 `fontTable.xml` 或只设置 `themeFontLang` 都不能替代 mixed-run 拆分。

正式方案仍是：**保留模板字体 + mixed run 最小拆分 + 拆分后的 run 显式字体槽**。


## 回归测试

修改 formatter / validator 后运行：

```bash
python -m unittest discover -s tests -p "test_*.py"
```

当前回归测试固定覆盖三类曾经实际出现的问题：

- 合并单元格遍历不得漏单元格；
- TOC/标题等样式的 `basedOn` 继承链必须一次同步收敛；
- 页脚修复后，default/even/first 引用及首页不同设置必须与模板语义一致。
