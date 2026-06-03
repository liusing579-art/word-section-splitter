#!/usr/bin/env python3
"""
WordSectionSplitter v3.0 — Python CLI + Library
从 Microsoft Word 导出的 HTML 中提取章节主体内容。

核心能力:
  - 章节识别: 第x节 → h2, 一/二/三... → h3, 第x章 → h2.chapter（去重）
  - 表格保留: TableNormal 表格完整保留，自动合并"续表"拆分的表格（18对→46表）
  - 表格智能样式: ≤6字居中+nowrap（标签格），>6字左对齐+word-break（内容格），th 始终居中
  - 排版统一: 移除所有内联 font-size/line-height/font-family，由 CSS 统一控制
  - Word 格式清理: 移除 <o:p>, mso-*, letter-spacing 极端值, mso-spacerun, lang, 空标签, 空段落
  - CJK 空白规范化: 移除中日韩字符间及字符与标点间的多余空格
  - 可折叠提示块: 7 个已知关键词 + 通用 [xxx]/【xxx】模式 → <details class="collapsible-tip" open>
  - 目录/思维导图过滤: 自动定位 TOC 结束和正文开始位置，跳过思维导图区域
  - 现代无衬线排版: Segoe UI / Noto Sans SC / PingFang SC, 16px, line-height 1.6

用法:
  python section_splitter.py input.html -o output.html
  python section_splitter.py input.html --stdout

依赖: pip install beautifulsoup4
"""

import re
import json
import argparse
import sys
from copy import deepcopy
from pathlib import Path

try:
    from bs4 import BeautifulSoup, Tag, NavigableString, Comment
except ImportError:
    print("错误：需要 BeautifulSoup4。请运行: pip install beautifulsoup4")
    sys.exit(1)


# ===== 书本风格 CSS =====
BOOK_CSS = """
:root {
  --page-bg: #f5f5f5; --text-color: #1a1a1a; --heading-color: #000;
  --accent: #1a5f8a; --accent-light: #d0d8e0; --table-border: #999;
  --table-header-bg: #e8ecf0; --table-stripe: #f7f8fa;
  --details-bg: #f4f6f8; --details-border: #b0bec5; --details-summary-bg: #e8ecf0;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  background: var(--page-bg); color: var(--text-color);
  font-family: 'Segoe UI', Roboto, 'Helvetica Neue', 'Noto Sans SC', 'PingFang SC', 'Microsoft YaHei', sans-serif;
  font-size: 16px; line-height: 1.6; padding: 40px 24px 60px; max-width: 1000px; margin: 0 auto;
  -webkit-font-smoothing: antialiased;
}
.book-content {
  background: #fff; border: 1px solid #e0e0e0; border-radius: 4px;
  padding: 48px 40px;
}
@media (max-width: 640px) {
  body { padding: 12px; } .book-content { padding: 24px 16px; }
}
.book-content h2 {
  font-size: 1.6em; color: var(--heading-color); margin: 32px 0 20px;
  padding-bottom: 10px; border-bottom: 2px solid var(--accent-light);
  text-align: center; font-weight: 700; letter-spacing: 1px;
}
.book-content h2.chapter { border-bottom-color: var(--accent); }
.book-content h3 {
  font-size: 1.25em; color: var(--heading-color); margin: 28px 0 16px; font-weight: 700;
}
.book-content h4, .book-content h5, .book-content h6 {
  font-size: 1.1em; color: var(--heading-color); margin: 20px 0 12px; font-weight: 600;
}
.book-content p { margin: 0 0 12px 0; }
.book-content table {
  width: 100%; table-layout: fixed; border-collapse: collapse; margin: 20px 0;
  font-size: 0.95em; border: 1.5px solid var(--table-border); overflow-x: auto;
}
.book-content table th {
  background: var(--table-header-bg); color: var(--heading-color);
  font-weight: 700; padding: 8px 10px; border: 1px solid var(--table-border); text-align: center;
}
.book-content table td {
  padding: 8px 10px; border: 1px solid var(--table-border); vertical-align: middle;
}
.book-content table tr:nth-child(even) td { background: var(--table-stripe); }
.book-content ol {
  padding-left: 0; list-style: none; counter-reset: item; margin: 12px 0;
}
.book-content ol li {
  counter-increment: item; margin-bottom: 20px; position: relative; padding-left: 2em;
}
.book-content ol li::before {
  content: counter(item) ". "; position: absolute; left: 0; font-weight: 600;
}
.book-content ol.content-list { list-style: none; padding-left: 0; margin: 12px 0; }
.book-content ol.content-list li { margin-bottom: 6px; }
.book-content ol.content-list li::before { content: none; }
.book-content ul { margin: 12px 0 12px 2em; }
.book-content li { margin-bottom: 6px; }
.book-content li p { display: inline; }
.book-content img { max-width: 100%; height: auto; display: block; margin: 16px auto; border-radius: 2px; }
details.collapsible-tip {
  margin: 20px 0; border: 1.5px solid var(--details-border);
  border-radius: 6px; background: var(--details-bg); overflow: hidden;
}
details.collapsible-tip[open] { background: #fff; }
details.collapsible-tip > summary {
  cursor: pointer; padding: 12px 16px; background: var(--details-summary-bg);
  font-weight: 700; color: var(--accent); font-size: 1.05em; user-select: none;
  list-style: none; display: flex; align-items: center; gap: 6px;
}
details.collapsible-tip > summary::-webkit-details-marker { display: none; }
details.collapsible-tip > summary::before {
  content: "\\25B6"; display: inline-block; font-size: 0.7em;
  transition: transform 0.2s; margin-right: 4px;
}
details.collapsible-tip[open] > summary::before { transform: rotate(90deg); }
details.collapsible-tip > :not(summary) { padding: 8px 16px 12px; }
.book-content b, .book-content strong { color: #000; }
.book-content u { text-decoration-color: #999; text-underline-offset: 2px; }
"""


# ===== 工具函数 =====

def _remove_cjk_spaces(text):
    """移除 CJK 字符之间及 CJK 与 CJK 标点之间的空格。"""
    CJK = '一-鿿㐀-䶿豈-﫿　-〿＀-￯⼀-⿟⺀-⻿'
    text = re.sub(f'([{CJK}])\\s+([{CJK}])', r'\1\2', text)
    text = re.sub(f'\\s+([、。，；：？！」』）】》""''])', r'\1', text)
    text = re.sub(f'([「『（【《])\\s+', r'\1', text)
    return text


def clean_txt(text):
    """规范化文本：替换不换行空格，合并空白，移除 CJK 间空格。"""
    text = text.replace('\xa0', ' ').replace('&nbsp;', ' ')
    text = re.sub(r'\s+', ' ', text)
    text = _remove_cjk_spaces(text)
    return text.strip()


def clean_style_string(style):
    """清理单个 style 属性字符串中的 Word 伪影。
    移除所有与统一 CSS 冲突的排版属性（字号/行距/字体），由 CSS 统一控制。"""
    if not style:
        return style
    # Remove mso-* properties
    for _ in range(4):
        new_style = re.sub(r'\s*mso-[^:;]+:[^;]*;?', '', style)
        if new_style == style:
            break
        style = new_style
    # Remove Word-specific properties
    style = re.sub(r'\s*tab-stops:[^;]*;?', '', style)
    style = re.sub(r'\s*mso-spacerun:yes;?', '', style)

    # Strip letter-spacing — unified CSS handles character spacing
    style = re.sub(r'\s*letter-spacing:[^;]*(?:;|$)', '', style)

    # Strip typography properties — unified CSS handles these
    style = re.sub(r'\s*font-size:[^;]*(?:;|$)', '', style)
    style = re.sub(r'\s*line-height:[^;]*(?:;|$)', '', style)
    style = re.sub(r'\s*font-family:[^;]*(?:;|$)', '', style)

    # Strip spacing properties — unified CSS handles these
    style = re.sub(r'\s*margin-(?:top|bottom|left|right):[^;]*(?:;|$)', '', style)
    style = re.sub(r'\s*text-indent:[^;]*(?:;|$)', '', style)
    # Strip border properties (except border-collapse) — CSS handles borders uniformly
    style = re.sub(r'\s*border(?!-collapse)(?:-\w+)?\s*:\s*[^;]*(?:;|$)', '', style)
    # Strip padding on table cells — CSS handles padding uniformly
    style = re.sub(r'\s*padding(?:-(?:top|bottom|left|right))?\s*:\s*[^;]*(?:;|$)', '', style)

    # Clean up
    style = re.sub(r';;+', ';', style)
    style = re.sub(r'^;+', '', style)
    style = re.sub(r';+$', '', style)
    return style.strip()


def clean_word_formatting(elem):
    """清理元素中的 Word 格式伪影（原地修改）。"""
    for tag in elem.find_all('o:p'):
        tag.decompose()
    for child in list(elem.descendants):
        if isinstance(child, Comment):
            child.extract()
    if elem.get('style'):
        cleaned = clean_style_string(elem['style'])
        if cleaned:
            elem['style'] = cleaned
        else:
            del elem['style']
    for tag in elem.find_all(style=True):
        style = tag.get('style', '')
        if not style:
            continue
        cleaned = clean_style_string(style)
        if cleaned:
            tag['style'] = cleaned
        else:
            del tag['style']
    for tag in elem.find_all('span', style=re.compile(r'mso-spacerun')):
        tag.unwrap()
    for tag in elem.find_all(attrs={'lang': True}):
        del tag['lang']
    for span in list(elem.find_all('span')):
        if not span.get('style') and not span.get('class') and not span.get('id'):
            if len(span.attrs) == 0:
                span.unwrap()
    for tag in list(elem.find_all('p')):
        text = tag.get_text().strip()
        if not text:
            tag.decompose()
    for tag in elem.find_all('p'):
        text = clean_txt(tag.get_text())
        if text in ['续 表', '续表']:
            tag.decompose()
    # Strip Word table attributes that conflict with CSS (root + descendants)
    for table in [elem] + elem.find_all('table'):
        for attr in ['width', 'cellpadding', 'cellspacing', 'align', 'valign', 'border']:
            if table.get(attr):
                del table[attr]
    return elem


def _normalize_html_text(html_string):
    """后处理完整 HTML 字符串，规范化标签间文本的空白。"""
    protected = {}

    def _protect(m):
        key = f'__PROTECTED_{len(protected)}__'
        protected[key] = m.group(0)
        return key

    html_string = re.sub(r'<style[^>]*>.*?</style>', _protect, html_string, flags=re.DOTALL)
    html_string = re.sub(r'<script[^>]*>.*?</script>', _protect, html_string, flags=re.DOTALL)
    html_string = html_string.replace('\xa0', ' ').replace('&nbsp;', ' ')

    def _normalize_segment(text):
        if not text.strip():
            return ''
        text = re.sub(r'\s+', ' ', text)
        text = _remove_cjk_spaces(text)
        return text.strip()

    parts = re.split(r'(<[^>]*>)', html_string)
    result = []
    for part in parts:
        if part.startswith('<'):
            result.append(part)
        else:
            result.append(_normalize_segment(part))
    html_string = ''.join(result)
    for key, content in protected.items():
        html_string = html_string.replace(key, content)
    return html_string


def _style_table_cells(table_elem):
    """
    统一表格单元格样式:
    - ≤6 个可见字符 → 居中 + nowrap（标签/关键词单元格）
    - >6 个字符 → 左对齐 + word-break（长文本自动换行，换行后左对齐更整齐）
    """
    for cell in table_elem.find_all(['td', 'th']):
        text = cell.get_text().strip()
        style = cell.get('style', '')
        style = re.sub(r'\s*white-space:[^;]*;?', '', style)
        style = re.sub(r'\s*text-align:[^;]*;?', '', style)
        style = re.sub(r'\s*word-break:[^;]*;?', '', style)
        style = style.strip('; ')

        if len(text) <= 6:
            style = (style + '; text-align:center; white-space:nowrap').strip('; ')
        else:
            style = (style + '; text-align:left; word-break:break-all').strip('; ')

        style = re.sub(r';;+', ';', style)
        cell['style'] = style


# ===== 元素提取 =====

def extract_elements(body):
    """
    从 WordSection div 中按文档顺序提取所有 <p> 和 <table> 元素。
    也提取 <w:sdt>（Word TOC 容器）中嵌套的 <p> 元素。
    返回: [{type, element, text, html, ...}]
    """
    result = []
    sections = body.find_all('div', class_=re.compile(r'WordSection'))

    for section in sections:
        for child in section.children:
            if not isinstance(child, Tag):
                continue
            if child.name == 'p':
                raw = child.get_text()
                text = clean_txt(raw)
                result.append({
                    'type': 'p', 'element': child, 'text': text, 'html': str(child),
                })
            elif child.name == 'table':
                raw = child.get_text()
                text = clean_txt(raw)
                cls = ' '.join(child.get('class', []))
                is_content_table = 'TableNormal' in cls
                result.append({
                    'type': 'table', 'element': child, 'text': text,
                    'html': str(child), 'is_content_table': is_content_table,
                })
            elif child.name == 'w:sdt':
                for inner_p in child.find_all('p'):
                    raw = inner_p.get_text()
                    text = clean_txt(raw)
                    result.append({
                        'type': 'p', 'element': inner_p, 'text': text,
                        'html': str(inner_p), 'from_w_sdt': True,
                    })
    return result


# ===== 表格合并 =====

def merge_split_tables(elements):
    """
    合并被 Word 分页拆分的表格。
    检测"续表"标记，将后续表格的行合并到前一表格，跳过重复表头。
    """
    merged = []
    i = 0
    merges_done = 0

    def _is_skip_paragraph(txt):
        if not txt:
            return True
        if txt in ['续 表', '续表', '续  表']:
            return True
        if re.match(r'^第[一二三四五六七八九十\d]+章', txt):
            return True
        if '法律硕士考试背诵宝典' in txt:
            return True
        if re.match(r'^第[一二三四五六七八九十\d]+章.*续表', txt):
            return True
        if re.match(r'^[\s/0-9]+$', txt):
            return True
        return False

    while i < len(elements):
        elem = elements[i]

        if elem['type'] != 'table' or not elem.get('is_content_table'):
            merged.append(elem)
            i += 1
            continue

        j = i + 1
        found_cont = False

        while j < len(elements):
            nxt = elements[j]
            if nxt['type'] == 'p' and _is_skip_paragraph(nxt['text']):
                if nxt['text'] in ['续 表', '续表', '续  表'] or '续表' in nxt['text']:
                    found_cont = True
                j += 1
                continue
            break

        if found_cont and j < len(elements) and elements[j]['type'] == 'table' and elements[j].get('is_content_table'):
            table1 = elem['element']
            table2 = elements[j]['element']

            rows1 = table1.find_all('tr')
            t1_header_text = rows1[0].get_text().strip() if rows1 else ''

            rows2 = table2.find_all('tr')
            skip_first = False
            if rows2 and rows1:
                t2_first_text = rows2[0].get_text().strip()
                if t2_first_text == t1_header_text:
                    skip_first = True
                if rows2[0].find('th'):
                    skip_first = True

            tbody = table1.find('tbody')
            target = tbody if tbody else table1

            for idx, row in enumerate(rows2):
                if skip_first and idx == 0:
                    continue
                target.append(row)

            elem['text'] = clean_txt(table1.get_text())
            elem['html'] = str(table1)
            merged.append(elem)
            i = j + 1
            merges_done += 1
        else:
            merged.append(elem)
            i += 1

    if merges_done:
        print(f'  Merged {merges_done} split tables')
    return merged


# ===== 边界检测 =====

def find_toc_end(elements):
    """找到目录结束位置。"""
    last_sdt = -1
    for i, e in enumerate(elements):
        if e.get('from_w_sdt'):
            last_sdt = i

    if last_sdt >= 0:
        toc_end = last_sdt + 1
        while toc_end < len(elements) and elements[toc_end]['type'] == 'p' and not elements[toc_end]['text']:
            toc_end += 1
        return toc_end

    para_texts = [(i, e['text']) for i, e in enumerate(elements) if e['type'] == 'p']
    best_end_elem = 0
    idx = 0
    para_count = len(para_texts)

    while idx < para_count:
        elem_idx, text = para_texts[idx]
        if re.search(r'第[一二三四五六七八九十\d]+[节章]', text):
            block_start = idx
            sec_count = 0
            last_idx = idx
            j = idx
            while j < para_count and j - block_start < 100:
                _, t = para_texts[j]
                if re.search(r'第[一二三四五六七八九十\d]+[节章]', t) and re.search(r'/\s*\d+', t):
                    sec_count += 1
                    last_idx = j
                elif not t:
                    pass
                else:
                    if sec_count <= 5:
                        break
                j += 1
            if sec_count > 10:
                best_end_elem = max(best_end_elem, para_texts[last_idx][0] + 1)
        idx += 1

    return best_end_elem


def find_content_start(elements, toc_end):
    """找到目录之后正文开始的位置。"""
    for i in range(toc_end, len(elements)):
        if elements[i]['type'] != 'p':
            continue
        text = elements[i]['text']
        if re.search(r'第[一二三四五六七八九十\d]+节', text):
            for j in range(i + 1, min(i + 10, len(elements))):
                if elements[j]['type'] != 'p':
                    continue
                nt = elements[j]['text']
                if nt and re.match(r'[一二三四五六七八九十]+[、.．]', nt):
                    return i
                if len(nt) > 60:
                    return i
    return toc_end


# ===== 提示检测 =====

def has_tip(text):
    """检测 [] 或 【】 中的提示关键词。返回关键词或 None。"""
    known = ['老马识途', '记忆线索', '主观点拨', '一招制敌', '归纳总结', '考点提示', '易错提醒']
    for kw in known:
        if f'[{kw}]' in text or f'【{kw}】' in text:
            return kw
    m = re.search(r'[\[【]([^\]】]{2,10})[\]】]', text)
    if m:
        candidate = m.group(1).strip()
        if not re.match(r'^[0-9a-zA-Z\s]+$', candidate):
            return candidate
    return None


# ===== HTML 构建 =====

def _convert_implicit_lists(html_string):
    """
    后处理：将连续编号段落（1. 2. 3. / (一)(二)(三) / ①②③）转为 <ol class="content-list">。
    文本自带编号，CSS 已禁用 ::before 计数器避免双重编号。
    """
    soup = BeautifulSoup(html_string, 'html.parser')
    book = soup.find('div', class_='book-content')
    if not book:
        return html_string

    # 编号检测模式（统一为 (pattern, type_key)）
    patterns = [
        (r'^\d+[\.、．]', 'arabic'),           # 1. 2. 3. 或 1、2、3、
        (r'^[\(（][一二三四五六七八九十]+[\)）]', 'cjk-paren'),  # (一)(二)(三)
        (r'^[①②③④⑤⑥⑦⑧⑨⑩]', 'circled'),            # ①②③
    ]

    def _is_numbered(p_tag):
        """检查 <p> 是否以编号开头，返回 (prefix, type_key) 或 None。"""
        text = p_tag.get_text().strip()
        if not text:
            return None
        for pat, ptype in patterns:
            m = re.match(pat, text)
            if m:
                return (m.group(), ptype)
        return None

    children = list(book.children)
    new_children = []
    i = 0

    while i < len(children):
        child = children[i]

        if not isinstance(child, Tag) or child.name != 'p':
            new_children.append(child)
            i += 1
            continue

        result = _is_numbered(child)
        if not result:
            new_children.append(child)
            i += 1
            continue

        prefix, ptype = result

        # 收集连续且相同模式的编号段落
        run = [child]
        j = i + 1
        while j < len(children):
            sibling = children[j]
            if not isinstance(sibling, Tag) or sibling.name != 'p':
                break
            next_result = _is_numbered(sibling)
            if not next_result:
                break
            next_prefix, next_type = next_result
            if next_type != ptype:
                break
            run.append(sibling)
            j += 1

        if len(run) < 2:
            # 少于 2 项不构成列表
            new_children.extend(run)
            i = j
            continue

        # 创建 <ol class="content-list">，包裹 run 中的 <p> 为 <li>
        ol_tag = soup.new_tag('ol')
        ol_tag['class'] = 'content-list'
        for p_tag in run:
            li_tag = soup.new_tag('li')
            for inner in list(p_tag.children):
                li_tag.append(inner)
            ol_tag.append(li_tag)

        new_children.append(ol_tag)
        i = j

    # 替换 book 的子节点
    book.clear()
    for child in new_children:
        book.append(child)

    return str(soup)


def build_html(elements, content_start):
    """从内容元素构建输出 HTML。"""
    output_soup = BeautifulSoup('', 'html.parser')
    book = output_soup.new_tag('div')
    book['class'] = 'book-content'

    stats = {'sections': 0, 'subsections': 0, 'tips': 0, 'chapters': 0, 'tables': 0, 'merged': 0}
    seen_chapters = set()

    i = content_start
    while i < len(elements):
        elem = elements[i]

        # ---- 表格 ----
        if elem['type'] == 'table':
            if elem.get('is_content_table'):
                table_copy = deepcopy(elem['element'])
                table_copy = clean_word_formatting(table_copy)
                _style_table_cells(table_copy)
                book.append(table_copy)
                stats['tables'] += 1
            i += 1
            continue

        # ---- 段落 ----
        text = elem['text']

        if not text:
            i += 1
            continue

        if text in ['续 表', '续表', '续  表']:
            i += 1
            continue

        if re.match(r'^[\s/0-9]+$', text):
            i += 1
            continue

        # 思维导图区域
        mind_map_indicators = ['思维导图', '导图']
        is_mindmap = any(ind in text for ind in mind_map_indicators)

        if is_mindmap:
            i += 1
            while i < len(elements):
                if elements[i]['type'] == 'table':
                    i += 1
                    continue
                nt = elements[i]['text']
                if not nt:
                    i += 1
                    continue
                if re.search(r'第[一二三四五六七八九十\d]+节', nt):
                    break
                if re.match(r'[一二三四五六七八九十]+[、.．]', nt) and len(nt) > 10:
                    break
                if len(nt) > 80:
                    break
                i += 1
            continue

        # 页眉噪声
        if '法律硕士考试背诵宝典' in text and len(text) < 50:
            i += 1
            continue
        if re.match(r'^第[一二三四五六七八九十\d]+章.*续表', text):
            i += 1
            continue

        # 章节分隔
        chap_match = re.search(r'第([一二三四五六七八九十\d]+)章', text)
        if chap_match and not re.search(r'第[一二三四五六七八九十\d]+节', text):
            if '续表' in text:
                i += 1
                continue
            chap_num = chap_match.group(1)
            if chap_num not in seen_chapters:
                h2 = output_soup.new_tag('h2')
                h2['class'] = 'chapter'
                h2.string = text
                book.append(h2)
                stats['chapters'] += 1
                seen_chapters.add(chap_num)
            i += 1
            continue

        # 节标题
        if re.search(r'第[一二三四五六七八九十\d]+节', text):
            h2 = output_soup.new_tag('h2')
            h2.string = text
            book.append(h2)
            stats['sections'] += 1
            i += 1
            continue

        # 次级标题（一、xxx）
        if re.match(r'^[\s　]*[一二三四五六七八九十]+[、.．]', text):
            h3 = output_soup.new_tag('h3')
            h3.string = text
            book.append(h3)
            stats['subsections'] += 1
            i += 1
            continue

        # 提示块
        tip_kw = has_tip(text)
        if tip_kw:
            if len(text) < 20:
                details = output_soup.new_tag('details')
                details['class'] = 'collapsible-tip'
                details['open'] = ''
                summary = output_soup.new_tag('summary')
                summary.string = f'[{tip_kw}]'
                details.append(summary)

                i += 1
                while i < len(elements):
                    if elements[i]['type'] == 'table':
                        table_copy = deepcopy(elements[i]['element'])
                        table_copy = clean_word_formatting(table_copy)
                        _style_table_cells(table_copy)
                        details.append(table_copy)
                        stats['tables'] += 1
                        i += 1
                        continue
                    nt = elements[i]['text']
                    if not nt:
                        i += 1
                        continue
                    if re.search(r'第[一二三四五六七八九十\d]+[节章]', nt):
                        break
                    if re.match(r'[一二三四五六七八九十]+[、.．]', nt):
                        break
                    if re.match(r'^[\(（][一二三四五六七八九十]+[\)）]', nt):
                        break
                    if has_tip(nt):
                        break
                    p_copy = deepcopy(elements[i]['element'])
                    p_copy = clean_word_formatting(p_copy)
                    details.append(p_copy)
                    i += 1

                book.append(details)
                stats['tips'] += 1
                continue
            else:
                details = output_soup.new_tag('details')
                details['class'] = 'collapsible-tip'
                details['open'] = ''
                summary = output_soup.new_tag('summary')
                summary.string = f'[{tip_kw}]'
                details.append(summary)
                p_copy = deepcopy(elem['element'])
                p_copy = clean_word_formatting(p_copy)
                details.append(p_copy)
                book.append(details)
                stats['tips'] += 1
                i += 1
                continue

        # 常规段落
        p_copy = deepcopy(elem['element'])
        p_copy = clean_word_formatting(p_copy)
        book.append(p_copy)
        i += 1

    return book, stats


# ===== 主类 =====

class WordSectionSplitter:
    """Word HTML 章节提取器（程序化调用入口）。"""

    def __init__(self, input_encoding='utf-8'):
        self.input_encoding = input_encoding

    @staticmethod
    def _detect_encoding(filepath):
        """自动检测文件编码（gb2312/gbk/utf-8）。"""
        for enc in ['utf-8', 'gb2312', 'gbk', 'gb18030']:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    f.read()
                return enc
            except (UnicodeDecodeError, UnicodeError):
                continue
        return 'utf-8'

    def process(self, input_path, output_path=None, json_path=None):
        """
        处理单个 HTML 文件。

        参数:
          input_path: 输入文件路径
          output_path: HTML 输出路径（可选）
          json_path: JSON 输出路径（可选）

        返回: (html_string, stats_dict)
        """
        encoding = self._detect_encoding(input_path)
        print(f'读取: {input_path} (编码: {encoding})')

        with open(input_path, 'r', encoding=encoding, errors='replace') as f:
            raw = f.read()

        soup = BeautifulSoup(raw, 'html.parser')
        body = soup.body if soup.body else soup
        elements = extract_elements(body)

        p_count = sum(1 for e in elements if e['type'] == 'p')
        t_count = sum(1 for e in elements if e['type'] == 'table')
        ct_count = sum(1 for e in elements if e['type'] == 'table' and e.get('is_content_table'))
        print(f'元素总数: {len(elements)} (p: {p_count}, 表格: {t_count}, 内容表格: {ct_count})')

        elements = merge_split_tables(elements)

        p_count = sum(1 for e in elements if e['type'] == 'p')
        t_count = sum(1 for e in elements if e['type'] == 'table')
        ct_count = sum(1 for e in elements if e['type'] == 'table' and e.get('is_content_table'))
        print(f'合并后: {len(elements)} (p: {p_count}, 表格: {t_count}, 内容表格: {ct_count})')

        toc_end = find_toc_end(elements)
        print(f'TOC 结束位置: {toc_end}')

        content_start = find_content_start(elements, toc_end)
        print(f'正文开始位置: {content_start}')

        if content_start < len(elements):
            e = elements[content_start]
            t = e['text'][:120] if e['text'] else '[table]'
            print(f'正文首元素: [{content_start}] <{e["type"]}> {t}')

        book, stats = build_html(elements, content_start)

        output_html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>{BOOK_CSS}</style>
</head>
<body>
{str(book)}
</body>
</html>'''

        output_html = _normalize_html_text(output_html)
        output_html = _convert_implicit_lists(output_html)

        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(output_html)
            print(f'HTML 已保存至: {output_path}')

        if json_path:
            json_output = json.dumps({'html': output_html}, ensure_ascii=False, indent=2)
            with open(json_path, 'w', encoding='utf-8') as f:
                f.write(json_output)
            print(f'JSON 已保存至: {json_path}')

        print(f'输入: {len(raw)/1024:.1f} KB → 输出: {len(output_html)/1024:.1f} KB')
        for k, v in stats.items():
            print(f'  {k}: {v}')

        return output_html, stats


# ===== CLI =====

def main():
    parser = argparse.ArgumentParser(
        description='WordSectionSplitter v3.0 — 从 Word HTML 提取章节内容',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  python section_splitter.py input.html -o output.html
  python section_splitter.py input.html -o output.html --json output.json
  python section_splitter.py input.html --json output.json
  python section_splitter.py input.html --stdout
        '''
    )
    parser.add_argument('input', help='输入的 Word HTML 文件路径')
    parser.add_argument('-o', '--output', help='输出 HTML 文件路径')
    parser.add_argument('--json', help='输出 JSON 文件路径')
    parser.add_argument('--stdout', action='store_true', help='直接输出 HTML 到标准输出')

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'错误：文件不存在 — {args.input}')
        sys.exit(1)

    splitter = WordSectionSplitter()
    html_output, stats = splitter.process(
        str(input_path),
        output_path=args.output,
        json_path=args.json,
    )

    if args.stdout or (not args.output and not args.json):
        print(html_output)

    print('完成。')


if __name__ == '__main__':
    main()
