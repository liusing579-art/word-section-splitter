#!/usr/bin/env python3
"""Process 101-150.htm — 法律硕士考试背诵宝典（第三部分）"""

from section_splitter import WordSectionSplitter

INPUT = 'C:/Users/Singing/Desktop/背诵宝典/101-150.htm'
OUTPUT = 'C:/Users/Singing/Desktop/产品及广告文案/法理学101-150/101-150-processed.html'


def main():
    splitter = WordSectionSplitter()
    splitter.process(INPUT, output_path=OUTPUT)


if __name__ == '__main__':
    main()
