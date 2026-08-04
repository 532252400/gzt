"""
100×100 打印标签生成功能
==========================
功能: 上传手机壳装箱Excel，生成10×10cm打印标签（HTML文件，可打印）
"""

import os, re, openpyxl
from collections import defaultdict


def get_date_from_name(name):
    """从文件名提取日期"""
    m = re.search(r'(\d{8})', name)
    return m.group(1)[:4]+'-'+m.group(1)[4:6]+'-'+m.group(1)[6:8] if m else None


def extract_name(name):
    """从文件名提取名称"""
    parts = name.split('-')
    for i,p in enumerate(parts):
        if any('\u4e00'<=c<='\u9fff' for c in p):
            orig = '-'.join(parts[i:]).replace('.xlsx','')
            m = re.search(r'(\d{8})', orig)
            return orig.replace(m.group(1),'').strip('-'), m.group(1) if m else (orig, None)
    return name.replace('.xlsx',''), None


def run_lbl100(fp):
    """
    生成100×100mm打印标签
    
    参数:
        fp: Excel文件路径
        
    返回:
        str: 结果消息
        
    Excel格式:
        方案A（有SKU列）:
            A列=SKU, B列=品名, C列=型号, D列=包数, E列=单包数量, F列=采购量, G列=箱号
        方案B（无SKU列）:
            C列=SKU, D列=品名, E列=型号, F列=包数, G列=数量, H列=总量, I列=箱号
    """
    wb = openpyxl.load_workbook(fp)
    ws = wb.active
    bn = os.path.basename(fp)
    nc, fd = extract_name(bn)
    ff = get_date_from_name(bn) or '0000-00-00'
    
    boxes = defaultdict(list)
    for r in range(2, ws.max_row + 1):
        sku = str(ws.cell(r, 1).value or '').strip()
        if not sku:
            # 方案B: SKU在C列
            sku = str(ws.cell(r, 3).value or '').strip()
            if not sku: 
                continue
            name = str(ws.cell(r, 4).value or '').strip()
            model = str(ws.cell(r, 5).value or '').strip()
            pk = str(ws.cell(r, 6).value or '')
            qty = str(ws.cell(r, 7).value or '')
            total = str(ws.cell(r, 8).value or '')
            box = str(ws.cell(r, 9).value or '').strip()
        else:
            # 方案A: SKU在A列
            name = str(ws.cell(r, 2).value or '').strip()
            model = str(ws.cell(r, 3).value or '').strip()
            pk = str(ws.cell(r, 4).value or '')
            qty = str(ws.cell(r, 5).value or '')
            total = str(ws.cell(r, 6).value or '')
            box = str(ws.cell(r, 7).value or '').strip()
        
        if box.isdigit():
            boxes[int(box)].append((sku, name, model, pk, qty, total))
    
    tb = len(boxes)  # 总箱数
    ts = sum((len(v) + 5) // 6 for v in boxes.values())  # 总张数
    
    sp = [(b, len(boxes[b]), (len(boxes[b]) + 5) // 6) 
          for b in sorted(boxes.keys()) 
          if (len(boxes[b]) + 5) // 6 > 1]
    
    # 生成HTML
    lh = ''
    for idx, bn_ in enumerate(sorted(boxes.keys()), 1):
        for ps in range(0, len(boxes[bn_]), 6):
            ch = boxes[bn_][ps:ps+6]
            lh += '<div class="l"><div class="bh">(第' + str(idx) + '箱/总' + str(tb) + '箱)</div>'
            lh += '<table><tr><th>SKU</th><th>品名</th><th>型号</th><th>包数</th><th>单包数量</th><th>采购量</th><th>箱号</th></tr>'
            for sku, name, model, pk, qty, total in ch:
                lh += '<tr><td style="font-size:10px">' + sku + '</td>'
                lh += '<td style="font-size:12px">' + name + '</td>'
                lh += '<td style="font-size:10px">' + model + '</td>'
                lh += '<td style="font-size:12px;text-align:center">' + pk + '</td>'
                lh += '<td style="font-size:12px;text-align:center">' + qty + '</td>'
                lh += '<td style="font-size:12px;text-align:center">' + total + '</td>'
                lh += '<td style="font-size:12px;text-align:center">' + str(idx) + '</td></tr>'
            lh += '</table></div>'
    
    style = '''
    <style>
    @page{size:100mm 100mm;margin:0}
    body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;margin:0;padding:0}
    .l{width:100mm;height:100mm;padding:1.5mm;page-break-after:always;overflow:hidden;
       display:flex;flex-direction:column}
    .l:last-child{page-break-after:auto}
    .bh{text-align:center;font-weight:bold;font-size:11px;padding:1mm 0}
    table{width:100%;border-collapse:collapse;table-layout:fixed}
    th,td{padding:0.8mm 0.5mm;border:0.7px solid #000;line-height:1.15;word-break:break-all}
    th{font-size:9px;text-align:center}
    th:nth-child(1),td:nth-child(1){width:18%}
    th:nth-child(2),td:nth-child(2){width:28%}
    th:nth-child(3),td:nth-child(3){width:14%}
    th:nth-child(4),td:nth-child(4){width:8%}
    th:nth-child(5),td:nth-child(5){width:10%}
    th:nth-child(6),td:nth-child(6){width:10%}
    th:nth-child(7),td:nth-child(7){width:12%}
    .np{text-align:center;padding:8px;background:#fff3cd;border-bottom:2px solid #ffc107}
    @media print{.np{display:none}}
    </style>'''
    
    html = ('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>'
            + nc + '标签</title>' + style + '</head><body>'
            + '<div class="np"><strong>' + nc + '标签</strong> | 10x10cm | '
            + str(ts) + '张 | <button onclick="window.print()" '
            + 'style="font-size:15px;padding:5px 18px">打印</button></div>'
            + lh + '</body></html>')
    
    # 输出到同级目录
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    fout = os.path.join(output_dir, ff + '-标签-' + nc + '.html')
    with open(fout, 'w', encoding='utf-8') as f:
        f.write(html)
    
    msg = '✅ ' + nc + ' 共 ' + str(ts) + ' 张\n文件：' + os.path.basename(fout)
    if sp:
        msg += '\n拆箱：' + ', '.join(['箱' + str(b) + '(' + str(c) + '款→' + str(p) + '张)' for b, c, p in sp])
    
    return msg


# ====== 独立测试 ======
if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        result = run_lbl100(sys.argv[1])
        print(result)
    else:
        print("用法: python 功能_100x100标签.py <Excel文件路径>")
