# -*- coding: utf-8 -*-
"""
国企公文格式 生成器
====================
复制「标准版模板」docx → 仅替换 body 内 <w:p> → 保留末位 <w:sectPr> 与页脚引用
→ 重新打包。从而 100% 继承标准版的页面设置与页脚页码。

用法:
  python gen_gongwen.py -t 模板.docx -c 内容.json -o 输出.docx

内容 JSON 格式: 二维数组 [[kind, text], ...]
  kind ∈ {title, date, zhusong, body, h1, h2, h3}
  - title   标题（可多行：每个数组项一行，按倒梯形排布，首行最长）
  - date    日期副标题，例 "(2026 年 08 月 16 日)"
  - zhusong 主送单位，例 "主送单位：市体育局"
  - body    正文段落（仿宋 / 两端对齐 / 首行缩进2字符）
  - h1      一级标题（一、二、三…，黑体加粗）
  - h2      二级标题（（一）（二）…，楷体）
  - h3      三级标题（（1）（2）…，仿宋）
  date 之后脚本会自动插入一个空行，无需在 JSON 写空项。
"""
import json
import os
import zipfile
import shutil
import argparse
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def w(t):
    return '{%s}%s' % (W, t)

# ---- 格式规范（与标准版模板 100% 一致） ----
SPEC = {
    'title':   dict(font='方正小标宋简体', sz=44, bold=False, jc='center', ind=0,   rsp=False),
    'date':    dict(font='楷体_GB2312',   sz=32, bold=False, jc='center', ind=0,   rsp=False),
    'zhusong': dict(font='仿宋_GB2312',   sz=32, bold=False, jc='left',   ind=0,   rsp=False),
    'body':    dict(font='仿宋_GB2312',   sz=32, bold=False, jc='both',   ind=640, rsp=True),
    'h1':      dict(font='黑体',          sz=32, bold=True,  jc='left',   ind=0,   rsp=True),
    'h2':      dict(font='楷体_GB2312',   sz=32, bold=False, jc='left',   ind=0,   rsp=True),
    'h3':      dict(font='仿宋_GB2312',   sz=32, bold=False, jc='left',   ind=0,   rsp=True),
}


def make_run(font, sz, bold, text, with_rspacing):
    r_el = etree.Element(w('r'))
    rPr = etree.SubElement(r_el, w('rPr'))
    rFonts = etree.SubElement(rPr, w('rFonts'))
    rFonts.set(w('ascii'), font)
    rFonts.set(w('hAnsi'), font)
    rFonts.set(w('eastAsia'), font)
    b = etree.SubElement(rPr, w('b'))
    if not bold:
        b.set(w('val'), '0')
    sz_el = etree.SubElement(rPr, w('sz'))
    sz_el.set(w('val'), str(sz))
    if with_rspacing:
        sp = etree.SubElement(rPr, w('spacing'))
        sp.set(w('val'), '8')
    t = etree.SubElement(r_el, w('t'))
    t.text = text
    return r_el


def make_empty_para():
    """日期后的空行：仅 keepNext + spacing line=560 exact，无文字、无对齐。"""
    p = etree.Element(w('p'))
    pPr = etree.SubElement(p, w('pPr'))
    etree.SubElement(pPr, w('keepNext'))
    sp = etree.SubElement(pPr, w('spacing'))
    sp.set(w('before'), '0')
    sp.set(w('after'), '0')
    sp.set(w('line'), '560')
    sp.set(w('lineRule'), 'exact')
    return p


def make_para(pkind, text):
    s = SPEC[pkind]
    p = etree.Element(w('p'))
    pPr = etree.SubElement(p, w('pPr'))
    # 所有段落都 keepNext（标题/日期/主送/各级标题/正文一致；小标题不分页）
    etree.SubElement(pPr, w('keepNext'))
    # 正文/各级标题额外 keepLines + widowControl
    if pkind in ('body', 'h1', 'h2', 'h3'):
        etree.SubElement(pPr, w('keepLines')).set(w('val'), '0')
        etree.SubElement(pPr, w('widowControl'))
    sp = etree.SubElement(pPr, w('spacing'))
    sp.set(w('before'), '0')
    sp.set(w('after'), '0')
    sp.set(w('line'), '560')
    sp.set(w('lineRule'), 'exact')
    ind = etree.SubElement(pPr, w('ind'))
    ind.set(w('firstLine'), str(s['ind']))
    jc = etree.SubElement(pPr, w('jc'))
    jc.set(w('val'), s['jc'])
    p.append(make_run(s['font'], s['sz'], s['bold'], text, s['rsp']))
    return p


def build(template, content, output):
    tmp = output + ".tmp"
    if os.path.exists(tmp):
        shutil.rmtree(tmp)
    os.makedirs(tmp, exist_ok=True)
    with zipfile.ZipFile(template, 'r') as z:
        z.extractall(tmp)

    doc_xml = os.path.join(tmp, 'word', 'document.xml')
    tree = etree.parse(doc_xml)
    root = tree.getroot()
    body = root.find(w('body'))
    children = list(body)
    sectPr = children[-1] if children and children[-1].tag == w('sectPr') else None
    for c in children:
        body.remove(c)
    if sectPr is not None:
        body.append(sectPr)

    for kind, text in content:
        p = make_para(kind, text)
        sectPr.addprevious(p)
        if kind == 'date':
            sectPr.addprevious(make_empty_para())

    tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 输出：若原文件被 Word 占用则改用 _修正.docx
    final = output
    base, ext = os.path.splitext(output)
    try:
        if os.path.exists(final):
            os.remove(final)
    except PermissionError:
        final = base + '_修正' + ext
    with zipfile.ZipFile(final, 'w', zipfile.ZIP_DEFLATED) as z:
        for rd, _, fs in os.walk(tmp):
            for fl in fs:
                fp = os.path.join(rd, fl)
                arc = os.path.relpath(fp, tmp)
                z.write(fp, arc)
    shutil.rmtree(tmp, ignore_errors=True)
    return final


def main():
    ap = argparse.ArgumentParser(description='国企公文格式生成器')
    ap.add_argument('-t', '--template', required=True, help='标准版模板 docx 路径')
    ap.add_argument('-c', '--content', required=True, help='内容 JSON 路径 [[kind,text],...]')
    ap.add_argument('-o', '--output', required=True, help='输出 docx 路径')
    args = ap.parse_args()

    if not os.path.exists(args.template):
        raise SystemExit('模板不存在: %s' % args.template)
    with open(args.content, 'r', encoding='utf-8') as f:
        content = json.load(f)

    final = build(args.template, content, args.output)
    print('生成成功:', final)
    print('段落数量:', len(content))


if __name__ == '__main__':
    main()
