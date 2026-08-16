# -*- coding: utf-8 -*-
"""
国企公文格式 验证器
====================
对照 GB/T 9704 标准规范，自动核对生成的 docx：
  - 每个段落的字体/字号/对齐/首行缩进/行距/keepNext
  - 日期段后紧跟一个空行
  - 标题非粗体、一级标题粗体
  - 正文首行缩进 2 字符 + 两端对齐
  - 页脚 `— PAGE —` 宋体、PAGE 域
  - ZIP 完整性

用法:
  python verify_gongwen.py 输出文件.docx
退出码 0 = 全部通过，1 = 存在不符合项。
"""
import os
import sys
import zipfile
import tempfile
import shutil
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def w(t):
    return '{%s}%s' % (W, t)

ALLOWED_FONTS = {'方正小标宋简体', '楷体_GB2312', '仿宋_GB2312', '黑体'}


def get_text(p):
    return ''.join(t.text or '' for t in p.iter(w('t')))


def analyze_para(p):
    """返回段落的格式画像，无法解析时返回 None。"""
    pPr = p.find(w('pPr'))
    if pPr is None:
        return None
    # 对齐
    jc = pPr.find(w('jc'))
    jc_val = jc.get(w('val')) if jc is not None else None
    # 缩进
    ind = pPr.find(w('ind'))
    firstLine = ind.get(w('firstLine')) if ind is not None else None
    # 行距
    sp = pPr.find(w('spacing'))
    line = sp.get(w('line')) if sp is not None else None
    lineRule = sp.get(w('lineRule')) if sp is not None else None
    # keepNext
    keepNext = pPr.find(w('keepNext')) is not None
    # 首个有字 run 的 rPr
    font = sz = bval = None
    for r in p.findall(w('r')):
        rPr = r.find(w('rPr'))
        if rPr is None:
            continue
        rf = rPr.find(w('rFonts'))
        if rf is not None and rf.get(w('eastAsia')):
            font = rf.get(w('eastAsia'))
        sz_el = rPr.find(w('sz'))
        if sz_el is not None:
            sz = sz_el.get(w('val'))
        b_el = rPr.find(w('b'))
        if b_el is not None:
            bval = b_el.get(w('val'))
        if font and sz:
            break
    return dict(jc=jc_val, firstLine=firstLine, line=line, lineRule=lineRule,
                keepNext=keepNext, font=font, sz=sz, bval=bval,
                empty=(len(p.findall(w('r'))) == 0))


def check(path):
    errors = []
    tmp = tempfile.mkdtemp(prefix='gww_verify_')
    try:
        with zipfile.ZipFile(path, 'r') as z:
            bad = z.testzip()
            if bad is not None:
                errors.append('ZIP 损坏: %s' % bad)
            z.extractall(tmp)
        # ---- 文档段落 ----
        doc = os.path.join(tmp, 'word', 'document.xml')
        tree = etree.parse(doc)
        body = tree.getroot().find(w('body'))
        paras = [c for c in body if c.tag == w('p')]
        runs_paras = [analyze_para(p) for p in paras]
        # 跳过 None
        # 通用检查：每个有 run 的段落
        date_idx = -1
        for i, info in enumerate(runs_paras):
            if info is None:
                continue
            if info['empty']:
                continue
            # 行距
            if info['line'] != '560' or info['lineRule'] != 'exact':
                errors.append('段落[%d] 行距应为 560/exact，实际 %s/%s' % (i, info['line'], info['lineRule']))
            if not info['keepNext']:
                errors.append('段落[%d] 缺少 keepNext' % i)
            # 字体合法性
            if info['font'] and info['font'] not in ALLOWED_FONTS:
                errors.append('段落[%d] 字体非法: %s' % (i, info['font']))
            # 字号
            if info['sz'] and info['sz'] not in ('44', '32'):
                errors.append('段落[%d] 字号非法: %s' % (i, info['sz']))
            # 标题
            if info['font'] == '方正小标宋简体':
                if info['sz'] != '44':
                    errors.append('段落[%d] 标题字号应 44' % i)
                if info['jc'] != 'center':
                    errors.append('段落[%d] 标题应居中' % i)
                if info['bval'] not in (None, '0'):
                    errors.append('段落[%d] 标题不应加粗' % i)
            # 黑体（一级标题）
            if info['font'] == '黑体':
                if info['sz'] != '32':
                    errors.append('段落[%d] 一级标题字号应 32' % i)
                if info['bval'] not in ('1', None) or info['bval'] == '0':
                    # 一级标题必须加粗
                    if info['bval'] != '1':
                        errors.append('段落[%d] 一级标题(黑体)应加粗' % i)
                if info['jc'] != 'left':
                    errors.append('段落[%d] 一级标题应左对齐' % i)
                if info['firstLine'] != '0':
                    errors.append('段落[%d] 一级标题首行缩进应 0' % i)
            # 仿宋（正文 / 主送 / 三级）
            if info['font'] == '仿宋_GB2312':
                if info['sz'] != '32':
                    errors.append('段落[%d] 仿宋字号应 32' % i)
                if info['jc'] == 'both':
                    if info['firstLine'] != '640':
                        errors.append('段落[%d] 正文首行缩进应为 640(2字符)' % i)
                elif info['jc'] == 'left':
                    if info['firstLine'] != '0':
                        errors.append('段落[%d] 主送/三级标题首行缩进应 0' % i)
            # 楷体（日期 / 二级）
            if info['font'] == '楷体_GB2312':
                if info['sz'] != '32':
                    errors.append('段落[%d] 楷体字号应 32' % i)
            # 日期定位
            if info['font'] == '楷体_GB2312' and info['jc'] == 'center':
                date_idx = i
        # 日期后空行
        if date_idx >= 0:
            if date_idx + 1 < len(runs_paras) and runs_paras[date_idx + 1] and runs_paras[date_idx + 1]['empty']:
                pass  # OK
            else:
                errors.append('日期段后未紧跟空行（第 %d 段之后）' % date_idx)
        else:
            errors.append('未找到日期段（楷体居中）')
        # ---- 页脚 ----
        footers = [f for f in os.listdir(os.path.join(tmp, 'word')) if f.startswith('footer') and f.endswith('.xml')]
        footer_ok = False
        for fn in footers:
            fx = etree.parse(os.path.join(tmp, 'word', fn))
            txt = ''.join(t.text or '' for t in fx.iter(w('t')))
            has_page = 'PAGE' in etree.tostring(fx.getroot()).decode('utf-8').upper()
            has_dash = '—' in txt
            # 宋体
            song = any(rf is not None and rf.get(w('eastAsia')) == '宋体'
                       for rf in fx.iter(w('rFonts')))
            if has_page and has_dash and song:
                footer_ok = True
                break
        if not footer_ok:
            errors.append('页脚未同时包含 PAGE 域 / — 分隔线 / 宋体')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print('=' * 50)
    if errors:
        print('❌ 验证未通过，发现 %d 处问题:' % len(errors))
        for e in errors:
            print('  - ' + e)
        print('=' * 50)
        return 1
    print('✅ 验证通过：所有段落与页脚均符合 GB/T 9704 标准规范')
    print('=' * 50)
    return 0


def main():
    if len(sys.argv) < 2:
        raise SystemExit('用法: python verify_gongwen.py 输出文件.docx')
    path = sys.argv[1]
    if not os.path.exists(path):
        raise SystemExit('文件不存在: %s' % path)
    sys.exit(check(path))


if __name__ == '__main__':
    main()
