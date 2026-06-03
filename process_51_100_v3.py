#!/usr/bin/env python3
"""
Process 51-100.htm — 法律硕士考试背诵宝典（第二部分）
薄封装脚本，调用 section_splitter 核心库处理指定文件。
"""

from section_splitter import WordSectionSplitter

INPUT = 'C:/Users/Singing/Desktop/背诵宝典/51-100.htm'
OUTPUT = 'C:/Users/Singing/Desktop/产品及广告文案/法理学51-100/51-100-processed.html'


def main():
    splitter = WordSectionSplitter()
    splitter.process(INPUT, output_path=OUTPUT)


if __name__ == '__main__':
    main()
