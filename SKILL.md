---
name: word-section-splitter
description: "从 Microsoft Word 导出的 HTML 文件中提取章节主体内容，保留表格结构，清理 Word 格式伪影，合并被分页拆分的表格，统一排版（移除内联字号/行距/字体），智能居中/左对齐表格单元格，将提示块转换为可折叠 details 元素，输出结构化 HTML。触发词：Word拆分、章节提取、Word转HTML、书本格式化、表格合并、续表清理、Word cleanup、section split。"
---

# WordSectionSplitter v3.0

从 Microsoft Word 导出的 HTML（gb2312/utf-8）中提取结构化章节内容，完整保留表格，清理 Word 格式伪影，统一排版，输出现代无衬线风格 HTML。

## 核心功能

1. **章节识别** — "第x节" → `<h2>`，"一、二、三…" → `<h3>`，"第x章" → `<h2 class="chapter">`（去重）
2. **表格完整保留** — 所有 `TableNormal` 表格完整保留，colspan/rowspan 不丢失
3. **续表合并** — 自动检测"续表"标记，合并被 Word 分页拆分的表格，跳过重复表头
4. **排版统一** — 移除所有内联 `font-size`/`line-height`/`font-family`，由 CSS 统一控制
5. **Word 格式清理** — 移除 `<o:p>`、`mso-*` CSS 属性、极端 letter-spacing、mso-spacerun、lang 属性、空 span、空段落
6. **表格单元格智能样式** — ≤6 字居中 + nowrap（标签格），>6 字左对齐 + word-break 自动换行（内容格），换行后左对齐更整齐
7. **CJK 空白规范化** — 移除中日韩字符之间及字符与标点之间的多余空格
8. **可折叠提示块** — 7 个已知关键词（记忆线索/老马识途/主观点拨/一招制敌/归纳总结/考点提示/易错提醒）+ 通用 `[xxx]`/`【xxx】` 模式 → `<details class="collapsible-tip" open>`
9. **思维导图/目录过滤** — 自动定位 TOC 结束和正文开始位置，跳过思维导图区域
10. **现代无衬线排版** — Segoe UI / Roboto / Noto Sans SC / PingFang SC，16px，line-height 1.6，max-width 1000px
11. **隐式列表转换** — 自动检测连续编号段落（1. 2. 3. / (一)(二)(三) / ①②③），转为 `<ol class="content-list">`，CSS 禁用计数器避免双重编号

## 三条核心规则

1. **续表合并** — 删除"续表"标记后，原先被拆分的两个表格合并为一个逻辑表格
2. **层级标记不被折叠** — "一、二、三"层级标记不会出现在 [记忆线索] 等折叠块内部
3. **所有 [xxx]/【xxx】括号内容** — 制作成可折叠 `<details>` 块

## 使用方式

### 方式一：直接处理 1-50.htm

```bash
cd C:/Users/Singing/.claude/skills/word-section-splitter
python process_1_50_v3.py
```

输出：`C:/Users/Singing/Desktop/产品及广告文案/法理学1-50/1-50-processed.html`

### 方式二：Python CLI 处理任意文件

```bash
python section_splitter.py input.html -o output.html
```

### 方式三：作为 Python 库使用

```python
from section_splitter import WordSectionSplitter

splitter = WordSectionSplitter()
html, stats = splitter.process('input.htm', output_path='output.html')
```

## 处理流程

1. 解析 HTML → DOM 树（BeautifulSoup，自动检测 gb2312/gbk/utf-8 编码）
2. 从 `WordSection` div 中按文档顺序提取 `<p>` 和 `<table>` 元素
3. 合并被"续表"标记拆分的表格（检测重复表头并跳过）
4. 定位目录结束位置 → 定位正文开始位置
5. 遍历元素构建输出 HTML（表格→清理+智能样式，思维导图→跳过，提示→`<details>`包裹）
6. 后处理：CJK 间空白规范化 + HTML 标签间文本空白规范化
7. 隐式列表转换：连续编号段落 → `<ol class="content-list"><li>...</li></ol>`
8. 注入统一 CSS → 输出完整 HTML

## 关键要素总结

### 排版统一原则
- **CSS 集中控制**：所有排版属性（字号/行距/字体）由 CSS 统一定义，禁止内联覆盖
- **清理策略**：`clean_style_string()` 剥离 `font-size`/`line-height`/`font-family`
- **CSS 参数**：`font-family: 'Segoe UI', Roboto, 'Helvetica Neue', 'Noto Sans SC', 'PingFang SC', 'Microsoft YaHei', sans-serif`
- **CSS 参数**：`font-size: 16px`, `line-height: 1.6`, `max-width: 1000px`

### 表格处理两原则
- **标签格**（≤6 字）：`text-align: center; white-space: nowrap`
- **内容格**（>6 字）：`text-align: left; word-break: break-all`（换行后左对齐更整齐）
- 所有单元格统一适用以上规则，th/td/rowspan/colspan 均不例外

### Word 格式清理清单
- `<o:p>` 标签 → 删除
- `mso-*` CSS 属性 → 删除
- `letter-spacing` 绝对值 <0.5pt 或 >2.0pt → 删除
- `mso-spacerun:yes` → 删除
- `lang` 属性 → 删除
- 空 `<span>`（无 style/class/id）→ unwrap
- 空 `<p>` 标签 → 删除
- `tab-stops` → 删除
- 续表标记（"续表"/"续 表"）→ 删除，表格合并

### 提示块检测机制
- 已知关键词列表：`['老马识途', '记忆线索', '主观点拨', '一招制敌', '归纳总结', '考点提示', '易错提醒']`
- 通用模式：`[xxx]` 或 `【xxx】`，2-10 个非数字字母字符
- 短文本（<20 字）→ 标记行，收集后续内容直到遇到下一个标题/提示
- 长文本 → 整段一起包裹在 `<details>` 中

### 边界检测
- **TOC 结束**：优先检测 `<w:sdt>` 标签（Word TOC 容器）；回退检测密集节标题块
- **正文开始**：TOC 后第一个"第x节"后紧跟"一、"或长文本（>60 字）的位置
- **提示块断点**：节/章标题、一/二/三 层级标记、(一)(二)(三) 括号标记、另一个提示块
- **思维导图断点**：节标题、长层级标记（>10 字）、长文本（>80 字）

## 参考文件

- `section_splitter.py` — 核心库 + CLI（完整管道实现）
- `process_1_50_v3.py` — 处理 1-50.htm 的薄封装
- `test_input.html` / `test_output.html` — 测试用例
