# 公文排版Word模板.dotx 参考说明

本文件只用于排版诊断和 Skill 维护。正常执行不要默认读取。

`assets/公文排版Word模板.dotx` 是模板格式事实来源；运行时规则见 `config/rules.json`。

## 页面与版心

- A4 纵向：`11906 × 16838` twips。
- 上/下/左/右边距：`2098 / 1984 / 1587 / 1474` twips。
- 页眉/页脚：`850 / 1417` twips。
- 正文可用宽度约 15.6 cm。
- 文档网格：`linePitch=576`，即 28.8 pt。

## 正文与标题

正文 `正文（默认）`：

- 中文：仿宋_GB2312。
- 西文：Times New Roman。
- 16 pt（三号）。
- 首行缩进 `640` twips。
- 固定行距 28.8 pt。

标题映射：

| Word 样式 | 实际用途 | 中文字体 | 字号 |
|---|---|---|---|
| Heading 1 | 文档主标题 | 方正小标宋简体 | 22 pt |
| Heading 2 | 正文一级标题 | 黑体 | 16 pt |
| Heading 3 | 正文二级标题 | 楷体_GB2312 | 16 pt |
| Heading 4 | 正文三级标题 | 仿宋_GB2312，加粗 | 16 pt |
| Heading 5 | 正文四级标题 | 仿宋_GB2312 | 16 pt |

可见编号采用 `一、`、`（一）`、`1.`、`（1）`，不假定为 Word 自动多级编号。

## 页码

模板启用奇偶页不同页脚：

- 偶数页左对齐；
- 奇数页右对齐；
- 页码格式 `— PAGE —`；
- 宋体 14 pt。

## Word Online 字体兼容经验

多轮 OneDrive / Word Online 对照测试得到：

- 把 `仿宋_GB2312` 改成 `FangSong`：无效；
- 同步修改 `fontTable.xml`：无效；
- 单独把 `themeFontLang eastAsia` 改成 `zh-CN`：不足以解决；
- 把普通文本中的中文和 ASCII 英文/数字拆为不同 run：有效；
- 拆 run 后恢复中文 `仿宋_GB2312`：仍然正常。

因此正式方案是：

```text
保留模板字体
+ themeFontLang eastAsia=zh-CN
+ 普通中西文混合 run 拆分
+ 显式字体槽
```

如果原 run 已经有直接字体设置，formatter 保留该设置；只有缺失的字体槽才从段落样式和 docDefaults 补齐。

含域、绘图、制表符、换行等特殊 OOXML 的 run 不做这种拆分。

## 表格与校验口径

- 表格最大宽度按版心限制；只缩小超宽表格，不把窄表格强制撑满。
- 合并表格中的 `vMerge` 续接单元格不作为独立可见单元格校验。
- 页脚校验比较影响渲染的语义，不要求无关 XML 结构逐节点相同。

## OOXML 安全

不要使用 `xml.etree.ElementTree` 对 `document.xml` 或 `styles.xml` 做整体重写。格式修改使用 `python-docx + lxml` 做局部操作。
