"""
收货单据核对功能
==================
功能: 上传 Receiving Record Excel，核对通知收货 vs 到货数据
"""

import os, openpyxl
from collections import defaultdict


def run_rc(fp):
    """
    核对收货单据
    
    参数:
        fp: Excel文件路径
        
    返回:
        str: 核对结果消息
        
    Excel格式要求:
        Sheet2 表:
        第1-2行=表头
        第3行起=数据
        
        B列=通知收货SKU  D列=通知收货数量
        F列=到货SKU     H列=到货数量
        C列=通知收货品名  G列=到货品名
    """
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb['Sheet2']
    
    # 读取通知收货数据（左侧）
    left = {}
    for r in range(3, ws.max_row + 1):
        sku = str(ws.cell(r, 2).value or '').strip()
        qty = ws.cell(r, 4).value
        if sku and qty is not None:
            try:
                left[sku] = left.get(sku, 0) + int(float(str(qty)))
            except:
                pass
    
    # 读取到货数据（右侧）
    right = {}
    for r in range(3, ws.max_row + 1):
        sku = str(ws.cell(r, 6).value or '').strip()
        qty = ws.cell(r, 8).value
        if sku and qty is not None:
            try:
                right[sku] = right.get(sku, 0) + int(float(str(qty)))
            except:
                pass
    
    # 读取品名
    names = {}
    for r in range(3, ws.max_row + 1):
        for sc, nc in [(2, 3), (6, 7)]:
            sku = str(ws.cell(r, sc).value or '').strip()
            nm = str(ws.cell(r, nc).value or '').strip()
            if sku:
                names[sku] = nm
    
    # 核对
    all_s = set(list(left.keys()) + list(right.keys()))
    match = sum(1 for s in all_s if left.get(s, 0) == right.get(s, 0))
    diff = [(s, names.get(s, ''), left.get(s, 0), right.get(s, 0))
            for s in sorted(all_s) if left.get(s, 0) != right.get(s, 0)]
    
    msg = '✅ 核对完成\n总SKU: ' + str(len(all_s)) + ' | 一致: ' + str(match) + ' | 差异: ' + str(len(diff))
    
    if diff:
        msg += '\n\n差异明细：'
        for s, nm, l, r in diff:
            msg += '\n' + s + ' | ' + nm[:20] + '\n  通知 ' + format(l, ',') + ' → 到货 ' + format(r, ',') + ' (差额 ' + format(l - r, ',') + ')'
    
    return msg


# ====== 独立测试 ======
if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        result = run_rc(sys.argv[1])
        print(result)
    else:
        print("用法: python 功能_收货单据核对.py <Excel文件路径>")
        print("注意: Excel需包含'Sheet2'表")
