#!/usr/bin/env python3
"""
Process 1-50.htm — 法律硕士考试背诵宝典（法理学）
调用 section_splitter 核心库 + 项目专属后处理。
"""

from section_splitter import WordSectionSplitter

INPUT = 'C:/Users/Singing/Desktop/背诵宝典/1-50/1-50.htm'
OUTPUT = 'C:/Users/Singing/Desktop/产品及广告文案/法理学1-50/1-50-processed.html'


def post_process(path):
    """项目专属：合并[主观点拨]提示块，删除中间的'第一章 法本体论'标题。"""
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 定位第二个 [主观点拨] details 块
    tag = '[主观点拨]'
    p1 = content.find(tag)
    p2 = content.find(tag, p1 + 1)
    ds = content.rfind('<details', 0, p2)
    de_tmp = content.find('</details>', p2)
    summary_end = content.find('</summary>', ds) + len('</summary>')

    # 清理 intro 段落中的重复 <span> 标签
    intro = content[summary_end:de_tmp]
    s1s = intro.find('<span')
    s1e = intro.find('</span>', s1s) + len('</span>')
    s2s = intro.find('<span', s1e)
    s2e = intro.find('</span>', s2s) + len('</span>')
    intro_clean = intro[:s1s] + intro[s1e:s2s] + intro[s2e:]

    # 提取 details 后的 <ol>（含段落 1、2）
    de = de_tmp + len('</details>')
    ol_start = content.find('<ol class="content-list">', de)
    ol_end = content.find('</ol>', ol_start) + len('</ol>')
    ol_html = content[ol_start:ol_end]

    # 跳过 <h2> 第一章 法本体论，提取段落 3
    h2_start = content.find('<h2 class="chapter">', ol_end)
    h2_end = content.find('</h2>', h2_start) + len('</h2>')
    p3_txt = content.find('分析法学派对现代自然法学派', h2_end)
    p3_start = content.rfind('<p', 0, p3_txt)
    p3_end = content.find('</p>', p3_txt) + len('</p>')

    # 构建新的 details 块
    new_block = (
        '<details class="collapsible-tip" open="">'
        '<summary>[主观点拨]</summary>'
        + intro_clean + ol_html + content[p3_start:p3_end] +
        '</details>'
    )

    content = content[:ds] + new_block + content[p3_end:]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('  后处理: 主观点拨合并完成')


def main():
    splitter = WordSectionSplitter()
    splitter.process(INPUT, output_path=OUTPUT)
    post_process(OUTPUT)


if __name__ == '__main__':
    main()
