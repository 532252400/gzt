"""
100×30 提货标签生成功能
==========================
功能: 上传司机提货信息Excel，生成100×30mm提货标签（HTML文件，可打印）
"""

import os, openpyxl


def run_lbl30(fp):
    """
    生成100×30mm提货标签
    
    参数:
        fp: Excel文件路径
        
    返回:
        str: 结果消息
        
    Excel格式:
        每行一条记录，前12列分别为:
        区域/仓库代码、客户名、箱数、重量、体积、日期、时间、
        司机名、车牌号、电话、备注
    """
    wb = openpyxl.load_workbook(fp)
    ws = wb.active
    
    labels = []
    for r in range(2, ws.max_row + 1):
        if not str(ws.cell(r, 1).value or '').strip():
            continue
        labels.append([str(ws.cell(r, c).value or '').strip() for c in range(1, 12)])
    
    num = len(labels)
    pages = (num + 1) // 2
    
    # 生成HTML
    lbs = ''
    for pi in range(0, num, 2):
        lbs += '<div class="p">'
        for j in range(2):
            i = pi + j
            if i < num:
                l = labels[i]
                lbs += '<div class="lb">'
                lbs += '<div class="rg">' + l[0] + '</div>'
                lbs += '<div class="bk">' + l[1] + '</div>'
                lbs += '<div class="dt">' + l[2] + '箱--' + l[3] + 'kg--' + l[4] + 'm3</div>'
                lbs += '<div class="dt tm">' + l[5][:10] + '--' + l[6][:5] + '</div>'
                lbs += '<div class="dt">' + l[7] + '--' + l[8] + '--' + l[9] + '</div>'
                lbs += '<div class="ph">' + l[10] + '</div>'
                lbs += '</div>'
        lbs += '</div>'
    
    style = '''
    <style>
    @page{size:100mm 30mm;margin:0}
    body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;margin:0;padding:0}
    .p{width:100mm;height:30mm;display:flex;page-break-after:always}
    .p:last-child{page-break-after:auto}
    .lb{width:50mm;height:30mm;padding-left:3mm;padding-top:2.5mm;overflow:hidden;line-height:1.15}
    .rg{font-size:14px;font-weight:bold}
    .bk{font-size:12px}
    .dt{font-size:12px}
    .tm{font-weight:bold}
    .ph{font-size:12px}
    .np{text-align:center;padding:8px;background:#fff3cd;border-bottom:2px solid #ffc107}
    @media print{.np{display:none}}
    </style>'''
    
    html = ('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>标签-提货信息</title>'
            + style + '</head><body>'
            + '<div class="np"><strong>标签-提货信息</strong> | 100x30mm | '
            + str(num) + '张 | 2列/页 | <button onclick="window.print()" '
            + 'style="font-size:15px;padding:5px 18px">打印</button></div>'
            + lbs + '</body></html>')
    
    # 输出到同级目录
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    fout = os.path.join(output_dir, '标签-提货信息.html')
    with open(fout, 'w', encoding='utf-8') as f:
        f.write(html)
    
    return '✅ ' + str(num) + '张标签，' + str(pages) + '页\n文件：' + os.path.basename(fout)


# ====== 独立测试 ======
if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        result = run_lbl30(sys.argv[1])
        print(result)
    else:
        print("用法: python 功能_100x30标签.py <Excel文件路径>")
