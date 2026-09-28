# 公文排版Word模板.dotx 参考说明

本文件只用于排版诊断和 Skill 维护。正常执行不要默认读取。

`assets/公文排版Word模板.dotx` 是模板格式事实来源；运行时规则见 `config/rules.json`。

## 页面与版心

模板基准为 A4 纵向：

- 页面：`11906 × 16838` twips；
- 上/下/左/右边距：`2098 / 1984 / 1587 / 1474` twips；
- 页眉/页脚：`850 / 1417` twips；
- 纵向正文可用宽度约 15.6 cm；
- 文档网格：`linePitch=576`。

已有 section 的方向属于文档结构，不应被模板无条件覆盖。纵向 section 使用模板 A4；横向 section 使用旋转后的 A4，并保留横向方向。表格宽度按其所在 section 的 `页面宽度 - 左右边距` 计算。

## 正文与标题

正文 `正文（默认）`：

- 中文：仿宋_GB2312；
- 西文：Times New Roman；
- 16 pt（三号）；
- 首行缩进 `640` twips；
- 固定行距 28.8 pt。

| Word 样式 | 实际用途 | 中文字体 | 字号 |
|---|---|---|---|
| Heading 1 | 文档主标题 | 方正小标宋简体 | 22 pt |
| Heading 2 | 正文一级标题 | 黑体 | 16 pt |
| Heading 3 | 正文二级标题 | 楷体_GB2312 | 16 pt |
| Heading 4 | 正文三级标题 | 仿宋_GB2312，加粗 | 16 pt |
| Heading 5 | 正文四级标题 | 仿宋_GB2312 | 16 pt |

可见编号采用 `一、`、`（一）`、`1.`、`（1）`，不假定为 Word 自动多级编号。

## Word Online mixed-run 经验

对照测试结论：

- `仿宋_GB2312 → FangSong`：无效；
- 修改 `fontTable.xml`：无效；
- 单独设置 `themeFontLang eastAsia=zh-CN`：不足；
- 真正有效的是拆分同一普通 run 内混合的东亚文字和 ASCII 字母/数字；
- 拆分后恢复 `仿宋_GB2312` 仍正常。

为避免长文档 run 数膨胀，正式规则采用最小拆分：

- 纯中文：不拆；
- 纯英文/数字：不拆；
- 只有标点/空格差异：不拆；
- 真正同时含东亚文字和 ASCII 字母/数字：拆；
- 中性标点附着到相邻片段，不自行形成 run；
- 含域、绘图、制表符、换行等特殊 OOXML 的 run 不拆。

## 大文档执行策略

正常入口是 validate-first：

```text
quick validate
→ 仅对发现的问题执行 fix_group
→ quick validate
→ 可选 strict 一次
```

已经通过 quick validation 的 DOCX 不重写。仅元数据需要修复时直接修改 ZIP 包，不通过 python-docx 重存整篇文档。

## 页码、表格与校验口径

模板奇偶页页码格式为 `— PAGE —`。多 section 文档的页面设置与表格宽度按 section 检查；页脚结构目前只检查并修复 final/default section，避免修一个页脚问题时重写所有分节页脚。

表格只缩小超出版心的情况，不把较窄表格强制撑满。合并表格中的 `vMerge` 续接单元格不作为独立可见单元格校验。

## OOXML 安全

不要使用 `xml.etree.ElementTree` 对 `document.xml` 或 `styles.xml` 做整体重写。格式修改使用 `python-docx + lxml` 做局部操作。
