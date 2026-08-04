"""Replace WORKSHOP_PAGE in test version with original simplified version"""
import os

filepath = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'

# Read current file
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find WORKSHOP_PAGE boundaries
start = content.find("WORKSHOP_PAGE = '''")
end = content.find("WORKSHOP_ADMIN = '''")
if start < 0 or end < 0:
    print("Could not find boundaries")
    exit(1)

# New simplified workshop page (using raw string to avoid escape issues)
new_ws = r"""WORKSHOP_PAGE = '''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"><title>车间加工</title><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f5f5f5;color:#333;padding:0;max-width:500px;margin:0 auto}
.hd{background:linear-gradient(135deg,#2d3748,#4a5568);color:#fff;padding:14px 16px;position:sticky;top:0;z-index:10}
.hd h1{font-size:17px}.hd p{font-size:11px;color:#cbd5e0;margin-top:2px}
.sel{width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;margin:10px 12px;width:calc(100% - 24px);background:#fff}
.filter-bar{display:flex;gap:4px;margin:0 12px 8px}
.filter-bar button{flex:1;padding:6px;border:1px solid #3182ce;border-radius:6px;font-size:12px;cursor:pointer;background:#fff;color:#3182ce;font-weight:500}
.filter-bar button.on{background:#e53e3e;border-color:#e53e3e;color:#fff}
.pri-badge{display:inline-block;background:#e53e3e;color:#fff;font-size:9px;padding:1px 5px;border-radius:3px;margin-left:6px;vertical-align:middle}
.li{list-style:none;padding:0;margin:0}
.li li{background:#fff;margin:8px 12px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.06);overflow:hidden}
.li .top{display:flex;align-items:center;padding:10px 12px 6px;gap:8px}
.li .st{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.li .st.pd{background:#a0aec0}.li .st.pr{background:#d69e2e;animation:bl 1s infinite}
.li .st.cp{background:#38a169}
@keyframes bl{50%{opacity:.4}}
.li .sk{font-size:13px;font-weight:600;flex:1}
.li .nm{font-size:11px;color:#718096;margin:0 12px 8px}
.li .qt{font-size:11px;color:#718096;margin:0 12px 2px}
.li .btns{display:flex;gap:6px;padding:6px 12px 10px}
.btns button{flex:1;padding:8px;border:none;border-radius:6px;font-size:12px;cursor:pointer;font-weight:500}
.btn-s{background:#3182ce;color:#fff}.btn-c{background:#38a169;color:#fff}.btn-d{background:#e2e8f0;color:#718096;cursor:not-allowed}
.ov{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:50;display:none;align-items:center;justify-content:center}
.ov .bx{background:#fff;border-radius:12px;padding:20px;width:300px;max-width:90%}
.ov .bx h3{font-size:15px;margin-bottom:12px}
.ov .bx input{width:100%;padding:10px;border:2px solid #3182ce;border-radius:8px;font-size:16px;text-align:center;margin-bottom:12px}
.ov .bx button{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;margin:0 4px}
.na{text-align:center;font-size:12px;color:#999;padding:40px 20px}
.cr{font-size:11px;color:#999;text-align:center;padding:12px;margin-top:8px}
.bt{font-size:11px;color:#cbd5e0;text-decoration:none}</style></head><body>
<div class="hd"><h1>🔧 车间加工列表</h1><p id="batchInfo">加载中...</p></div>
<select class="sel" id="batchSel" onchange="loadJobs()"></select>
<div class="filter-bar"><button id="f_all" class="on" onclick="setFilter('all')">📋 全部</button><button id="f_priority" onclick="setFilter('priority')">⭐ 优先</button></div>
<ul class="li" id="jobList"></ul>
<div class="ov" id="modal"><div class="bx"><h3 id="mTitle">完成加工</h3><p style="font-size:12px;color:#666;margin-bottom:6px" id="mSku"></p><input id="mQty" type="number" placeholder="输入完成数量" min="1" onkeydown="if(event.key==='Enter')confirmComplete()"><button class="btn-s" onclick="confirmComplete()">确定完成</button><button onclick="closeModal()" style="background:#e2e8f0">取消</button></div></div>
<p class="cr"><a class="bt" href="/workshop_admin">管理后台</a> | <a class="bt" href="/">工作台</a></p>
<script>
var curItemId = 0;
var curFilter = 'all';
function setFilter(mode){
    curFilter=mode;
    document.querySelectorAll('.filter-bar button').forEach(function(b){b.className=''});
    document.getElementById('f_'+mode).className='on';
    loadJobs();
}
async function loadJobs(){
    var bid=document.getElementById('batchSel').value;
    if(!bid) return;
    var r=await fetch('/workshop_data?batch='+bid);var items=await r.json();
    var ul=document.getElementById('jobList');ul.innerHTML='';
    if(!items.length){ul.innerHTML='<div class="na">暂无加工单<br>请先在工作台上传</div>';return}
    if(curFilter==='priority') items=items.filter(function(i){return i.priority==1});
    items.forEach(function(i){
        var li=document.createElement('li');
        var statusLabel={pending:'待处理',processing:'加工中',completed:'已完成'};
        var stClass = i.status;
        var btns='';
        if(i.status==='pending'){btns='<button class="btn-s" onclick="startJob('+i.id+')">\u25b6 \u5f00\u59cb\u52a0\u5de5</button>';}
        else if(i.status==='processing'){btns='<button class="btn-c" onclick="openComplete('+i.id+',\\''+i.sku+'\\')">\u2714 \u5b8c\u6210\u52a0\u5de5</button>';}
        else{btns='<button class="btn-d">\u2714 \u5df2\u5b8c\u6210</button>';}
        var info='<div class="qt">\u6570\u91cf: '+i.qty+'</div>';
        if(i.status==='processing' && i.worker) info+='<div class="qt">\u5de5\u4eba: '+i.worker+' | \u5f00\u59cb: '+i.started.substr(11,5)+'</div>';
        if(i.status==='completed' && i.done_qty) info+='<div class="qt">\u5b8c\u6210: '+i.done_qty+'\u4ef6 | '+i.completed.substr(11,5)+'</div>';
        var priBadge = i.priority ? '<span class="pri-badge">\u2605\u4f18</span>' : '';
        li.innerHTML='<div class="top"><span class="st '+stClass+'"></span><span class="sk">'+i.sku+priBadge+'</span><span style="font-size:10px;color:#999">'+statusLabel[i.status]+'</span></div>'+'<div class="nm">'+(i.name||'')+'</div>'+info+'<div class="btns">'+btns+'</div>';
        ul.appendChild(li);
    });
}
async function startJob(id){
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker','\u8f66\u95f4\u5de5\u4eba');
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u2705');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
function openComplete(id,sku){curItemId=id;document.getElementById('mSku').textContent=sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}
function closeModal(){document.getElementById('modal').style.display='none';}
async function confirmComplete(){
    var qty=document.getElementById('mQty').value;
    if(!qty||parseInt(qty)<=0){alert('\u8bf7\u8f93\u5165\u6709\u6548\u6570\u91cf');return}
    var fd=new FormData();fd.append('action','complete_job');fd.append('batch_name',curItemId);fd.append('done_qty',qty);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();closeModal();if(d.status==='ok'){alert(d.message);loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
    var sel=document.getElementById('batchSel');
    bs.forEach(function(b){var o=document.createElement('option');o.value=b.id;o.textContent=b.name+' \u00b7 '+b.count+'\u9879';sel.appendChild(o);});
    if(bs.length){document.getElementById('batchInfo').textContent='\u5f53\u524d: '+bs[0].name;document.getElementById('batchSel').value=bs[0].id;loadJobs();}
    else document.getElementById('batchInfo').textContent='\u6682\u65e0\u52a0\u5de5\u5355';
});
</script></body></html>'''"""

# Replace the old WORKSHOP_PAGE with the new one
old_ws = content[start:end]
new_content = content[:start] + new_ws + content[end:]

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("WORKSHOP_PAGE replaced successfully!")
