"""Restore feature-rich workshop page with all original features (scan, QR code, people bar, etc.)"""
import os

filepath = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find WORKSHOP_PAGE boundaries
start = content.find("WORKSHOP_PAGE = '''")
end = content.find("WORKSHOP_ADMIN = '''")
if start < 0 or end < 0:
    print("Could not find boundaries")
    exit(1)

# The full feature-rich workshop page from ws_final.txt
new_ws = r"""WORKSHOP_PAGE = '''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"><title>车间加工</title><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f5f5f5;color:#333;padding:0;max-width:500px;margin:0 auto}
.hd{background:linear-gradient(135deg,#2d3748,#4a5568);color:#fff;padding:14px 16px;position:sticky;top:0;z-index:10}
.hd h1{font-size:17px}.hd p{font-size:11px;color:#cbd5e0;margin-top:2px}
.sel{width:100%;padding:8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;margin:10px 12px;width:calc(100% - 24px);background:#fff}
.filter-bar{display:flex;gap:4px;margin:4px 12px 8px}
.filter-bar button{flex:1;padding:6px;border:1px solid #3182ce;border-radius:6px;font-size:12px;cursor:pointer;background:#fff;color:#3182ce;font-weight:500}
.filter-bar button.on{background:#e53e3e;border-color:#e53e3e;color:#fff}
.pri-badge{display:inline-block;background:#e53e3e;color:#fff;font-size:9px;padding:1px 5px;border-radius:3px;margin-left:6px;vertical-align:middle}
.li{list-style:none;padding:0;margin:0}
.li li{background:#fff;margin:8px 12px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.06);overflow:hidden}
.li .top{display:flex;align-items:center;padding:10px 12px 6px;gap:8px}
.li .st{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.li .st.pending{background:#a0aec0}.li .st.processing{background:#d69e2e;animation:bl 1s infinite}.li .st.paused{background:#ed8936}.li .st.completed{background:#38a169}
@keyframes bl{50%{opacity:.4}}
.li .sk{font-size:13px;font-weight:600;flex:1}
.li .nm{font-size:11px;color:#718096;margin:0 12px 8px}
.li .qt{font-size:11px;color:#718096;margin:0 12px 2px}
.li .btns{display:flex;gap:6px;padding:6px 12px 10px}
.btns button{flex:1;padding:8px;border:none;border-radius:6px;font-size:11px;cursor:pointer;font-weight:500}
.btn-s{background:#3182ce;color:#fff}.btn-c{background:#38a169;color:#fff}.btn-x{background:#e53e3e;color:#fff}.btn-d{background:#e2e8f0;color:#718096;cursor:not-allowed}
.ov{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:50;display:none;align-items:center;justify-content:center}
.ov .bx{background:#fff;border-radius:12px;padding:20px;width:300px;max-width:90%}
.ov .bx h3{font-size:15px;margin-bottom:12px}
.ov .bx input{width:100%;padding:10px;border:2px solid #3182ce;border-radius:8px;font-size:16px;text-align:center;margin-bottom:12px}
.ov .bx button{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;margin:0 4px}
.na{text-align:center;font-size:12px;color:#999;padding:40px 20px}
.cr{font-size:11px;color:#999;text-align:center;padding:12px;margin-top:8px}
.bt{font-size:11px;color:#cbd5e0;text-decoration:none}
.pri-item{border-left:3px solid #e53e3e}
.pri-notice{background:#fff5f5;border:1px solid #fed7d7;border-radius:8px;padding:10px 14px;margin:8px 12px;display:none;font-size:12px;color:#e53e3e}
.people-row{display:flex;gap:8px;align-items:center;margin:8px 12px}
.people-row input{flex:1;padding:6px 10px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;width:60px}
.people-row button{padding:6px 12px;border:none;border-radius:6px;font-size:11px;cursor:pointer;background:#3182ce;color:#fff}
.people-row label{font-size:11px;color:#718096}
.tool-row{display:flex;gap:4px;margin:4px 12px 8px}
.tool-row button{flex:1;padding:6px;border:1px solid #e2e8f0;border-radius:6px;font-size:11px;cursor:pointer;background:#fff;color:#4a5568}
.tool-row button:hover{background:#f7fafc;border-color:#3182ce}
.qr-box{text-align:center;padding:10px}
.qr-box img{width:200px;height:200px;margin:10px 0}
.qr-box button{padding:8px 16px;border:none;border-radius:6px;font-size:12px;cursor:pointer;background:#3182ce;color:#fff;margin:0 4px}
#scanVideo{width:100%;max-height:300px;background:#000;border-radius:8px}
.scan-result{font-size:14px;font-weight:600;color:#38a169;text-align:center;padding:8px}
.scan-input-row{display:flex;gap:8px;margin-top:8px}
.scan-input-row input{flex:1;padding:8px;border:2px solid #3182ce;border-radius:6px;font-size:14px;text-align:center}
.scan-input-row button{padding:8px 16px;border:none;border-radius:6px;font-size:12px;cursor:pointer;background:#3182ce;color:#fff}
</style></head><body>
<div class="hd"><h1>\U0001f527 \u8f66\u95f4\u52a0\u5de5</h1><p id="batchInfo">\u52a0\u8f7d\u4e2d...</p></div>
<div id="priNotice" class="pri-notice">\u2b50 \u6709 <span id="priCount">0</span> \u4e2a\u4f18\u5148\u8ba2\u5355\u5f85\u5904\u7406</div>
<div class="filter-bar"><button id="f_pending" class="on" onclick="setFilter('pending')">\U0001f4cb \u5f85\u5904\u7406</button><button id="f_processing" onclick="setFilter('processing')">\U0001f527 \u52a0\u5de5\u4e2d</button><button id="f_completed" onclick="setFilter('completed')">\u2705 \u5df2\u5b8c\u6210</button><button id="f_priority" onclick="setFilter('priority')">\u2b50 \u4f18\u5148</button></div>
<div class="people-row"><span>\U0001f465 \u4eba\u6570:</span><input id="peopleInput" type="number" min="1" value="1" placeholder="\u4eba\u6570"><button onclick="savePeople()">\u8bbe\u7f6e</button><label id="peopleLabel"></label></div>
<div class="tool-row"><button onclick="showQR()">\U0001f4f1 \u663e\u793a\u4e8c\u7ef4\u7801</button><button onclick="startScan()">\U0001f4f7 \u626b\u7801\u67e5\u627e</button></div>
<ul class="li" id="jobList"></ul>
<!-- Complete Modal -->
<div class="ov" id="modal"><div class="bx"><h3 id="mTitle">\u5b8c\u6210\u52a0\u5de5</h3><p style="font-size:12px;color:#666;margin-bottom:6px" id="mSku"></p><input id="mQty" type="number" placeholder="\u8f93\u5165\u5b8c\u6210\u6570\u91cf" min="1" onkeydown="if(event.key==='Enter')confirmComplete()"><button class="btn-c" onclick="confirmComplete()">\u786e\u5b9a\u5b8c\u6210</button><button class="btn-x" onclick="closeModal()">\u53d6\u6d88</button></div></div>
<!-- QR Modal -->
<div class="ov" id="qrModal"><div class="bx"><h3>\U0001f4f1 \u626b\u7801\u6253\u5f00\u8f66\u95f4\u9875\u9762</h3><div class="qr-box"><img id="qrImg" src=""><br><button onclick="downloadQR()">\u2b07 \u4e0b\u8f7d\u4e8c\u7ef4\u7801</button><button onclick="document.getElementById('qrModal').style.display='none'" style="background:#e2e8f0;color:#333">\u5173\u95ed</button></div></div></div>
<!-- Scan Modal -->
<div class="ov" id="scanModal"><div class="bx"><h3>\U0001f4f7 \u626b\u7801\u67e5\u627e</h3><video id="scanVideo" autoplay playsinline></video><div id="scanSku" class="scan-result"></div><div id="scanManual" style="display:none"><div class="scan-input-row"><input id="scanInput" placeholder="\u8f93\u5165SKU\u7f16\u7801" onkeydown="if(event.key==='Enter')scanManualSearch()"><button onclick="scanManualSearch()">\u67e5\u627e</button></div></div><p style="font-size:11px;color:#999;text-align:center;margin-top:8px"><a href="#" onclick="stopScan();return false" style="color:#e53e3e">\u5173\u95ed\u626b\u7801</a></p></div></div>
<p class="cr"><a class="bt" href="/workshop_admin">\u7ba1\u7406\u540e\u53f0</a> | <a class="bt" href="/">\u5de5\u4f5c\u53f0</a></p>
<script>
var curItemId = 0;
var curFilter = 'pending';
function setFilter(mode){
    curFilter=mode;
    document.querySelectorAll('.filter-bar button').forEach(function(b){b.className=''});
    document.getElementById('f_'+mode).className='on';
    loadJobs();
}
async function loadJobs(){
    var ppl=localStorage.getItem('default_people')||'1';
    var r=await fetch('/workshop_data?people='+ppl);var items=await r.json();
    window.allItems=items;
    var ul=document.getElementById('jobList');ul.innerHTML='';
    if(!items.length){ul.innerHTML='<div class="na">\u6682\u65e0\u52a0\u5de5\u5355<br>\u8bf7\u5148\u5728\u5de5\u4f5c\u53f0\u4e0a\u4f20</div>';return}
    var info=document.getElementById('batchInfo');
    if(info) info.textContent='\u5168\u90e8\u52a0\u5de5\u5355 ('+items.length+'\u9879)';
    var priPending=items.filter(function(i){return i.status==='pending' && i.priority==1}).length;
    var el=document.getElementById('priNotice');
    if(el){if(priPending>0){el.style.display='block';document.getElementById('priCount').textContent=priPending}else el.style.display='none'}
    var ti=items;
    document.getElementById('f_pending').textContent='\U0001f4cb \u5f85\u5904\u7406 ('+ti.filter(function(x){return x.status==='pending'}).length+')';
    document.getElementById('f_processing').textContent='\U0001f527 \u52a0\u5de5\u4e2d ('+ti.filter(function(x){return x.status==='processing'}).length+')';
    document.getElementById('f_completed').textContent='\u2705 \u5df2\u5b8c\u6210 ('+ti.filter(function(x){return x.status==='completed'}).length+')';
    document.getElementById('f_priority').textContent='\u2b50 \u4f18\u5148 ('+ti.filter(function(x){return x.priority==1}).length+')';
    if(curFilter==='all') items=items.filter(function(i){return i.status!=='completed'});
    if(curFilter==='pending') items=items.filter(function(i){return i.status==='pending'});
    if(curFilter==='priority') items=items.filter(function(i){return i.priority==1});
    if(curFilter==='processing') items=items.filter(function(i){return i.status==='processing'});
    if(curFilter==='completed') items=items.filter(function(i){return i.status==='completed'});
    items.forEach(function(i){
        var li=document.createElement('li');
        if(i.priority) li.className='pri-item';
        var statusLabel={pending:'\u5f85\u5904\u7406',processing:'\u52a0\u5de5\u4e2d',paused:'\u5df2\u6682\u505c',completed:'\u5df2\u5b8c\u6210'};
        var stClass = i.status;
        var btns='';
        if(i.status==='pending'){btns='<button class="btn-s" onclick="startJob('+i.id+')">\u25b6 \u5f00\u59cb\u52a0\u5de5</button>';}
        else if(i.status==='processing'){btns='<button class="btn-c" onclick="openComplete(this)" data-qty="'+i.qty+'">\u2714 \u5b8c\u6210</button><button class="btn-x" style="background:#d69e2e" onclick="pauseJob('+i.id+')">\u23f8 \u6682\u505c</button><button class="btn-x" onclick="cancelJob('+i.id+')">\u2716 \u53d6\u6d88</button>';}
        else if(i.status==='paused'){btns='<button class="btn-s" onclick="resumeJob('+i.id+')">\u25b6 \u53d6\u6d88\u6682\u505c</button><button class="btn-x" onclick="cancelJob('+i.id+')">\u2716 \u53d6\u6d88</button>';}
        else{btns='<button class="btn-d">\u2714 \u5df2\u5b8c\u6210</button>';}
        var info='<div class="qt">\u6570\u91cf: '+i.qty+'</div>';
        if(i.status==='completed' && i.done_qty) info+='<div class="qt">\u5b8c\u6210: '+i.done_qty+'\u4ef6 | '+i.completed.substr(11,5)+'</div>';
        var priBadge = i.priority ? '<span class="pri-badge">\u2b50\u4f18\u5148</span>' : '';
        li.innerHTML='<div class="top"><span class="st '+stClass+'"></span><span class="sk">'+i.sku+priBadge+'</span><span style="font-size:10px;color:#999">'+statusLabel[i.status]+'</span></div>'+'<div class="nm">'+(i.name||'')+'</div>'+info+'<div class="btns">'+btns+'</div>';
        ul.appendChild(li);
    });
}
async function startJob(id){
    var ppl=document.getElementById('peopleInput')?document.getElementById('peopleInput').value:'';
    if(!ppl||parseInt(ppl)<1)ppl='1';
    localStorage.setItem('default_people',ppl);
    var wn='\u8f66\u95f4\u5de5\u4eba x'+ppl;
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker',wn);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u2705');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
function savePeople(){
    var v=document.getElementById('peopleInput').value;
    if(v&&parseInt(v)>=1){localStorage.setItem('default_people',v);document.getElementById('peopleLabel').textContent='\u4e0a\u73ed: '+v+'\u4eba';}
}
async function cancelJob(id){
    if(!confirm('\u786e\u5b9a\u53d6\u6d88\u6b64\u52a0\u5de5\u4efb\u52a1\uff1f'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u2705');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
async function pauseJob(id){
    var fd=new FormData();fd.append('action','pause_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u23f8');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
async function resumeJob(id){
    var fd=new FormData();fd.append('action','resume_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u25b6');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
function openComplete(btn){curItemId=btn.dataset.id;curOrderQty=btn.dataset.qty||'';document.getElementById('mSku').textContent=btn.dataset.sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}
function fillQty(){document.getElementById('mQty').value=curOrderQty||''}
function closeModal(){document.getElementById('modal').style.display='none';}
async function confirmComplete(){
    var qty=document.getElementById('mQty').value;
    if(!qty||parseInt(qty)<=0){alert('\u8bf7\u8f93\u5165\u6709\u6548\u6570\u91cf');return}
    var fd=new FormData();fd.append('action','complete_job');fd.append('batch_name',curItemId);fd.append('done_qty',qty);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();closeModal();if(d.status==='ok'){alert(d.message);loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}
var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='\u4e0a\u73ed: '+savedPpl+'\u4eba';}
document.getElementById('batchInfo').textContent='\u5168\u90e8\u52a0\u5de5\u5355';
loadJobs();
function autoRefresh(){loadJobs();setTimeout(autoRefresh,2000);}
autoRefresh();
function showQR(){
    var url=window.location.href;
    document.getElementById('qrImg').src='https://api.qrserver.com/v1/create-qr-code/?size=300x300&data='+encodeURIComponent(url);
    document.getElementById('qrModal').style.display='flex';
}
function downloadQR(){
    var url=window.location.href;
    var a=document.createElement('a');a.href='https://api.qrserver.com/v1/create-qr-code/?size=500x500&data='+encodeURIComponent(url);a.download='workshop_qr.png';a.click();
}
var scanStream=null;
function startScan(){
    document.getElementById('scanModal').style.display='flex';
    var v=document.getElementById('scanVideo');
    if(navigator.mediaDevices&&navigator.mediaDevices.getUserMedia){
        navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}}).then(function(s){
            scanStream=s;v.srcObject=s;v.play();
            sL();
        }).catch(function(){document.getElementById('scanManual').style.display='block'});
    }else document.getElementById('scanManual').style.display='block';
}
function sL(){
    var v=document.getElementById('scanVideo');
    if(!v.videoWidth) return;
    if('BarcodeDetector' in window){
        var d=new BarcodeDetector({formats:['ean_13','ean_8','code_128','code_39']});
        d.detect(v).then(function(c){if(c.length>0){sF(c[0].rawValue);return}setTimeout(sL,500)}).catch(function(){setTimeout(sL,500)});
    }else setTimeout(sL,500);
}
function sF(code){
    document.getElementById('scanSku').textContent='\u2705 '+code;
    setTimeout(function(){
    document.getElementById('scanModal').style.display='none';
    document.getElementById('scanManual').style.display='none';
    if(scanStream){scanStream.getTracks().forEach(function(t){t.stop()});scanStream=null}
    var list=window.allItems||[];code=code.toUpperCase();var mt=null;
    for(var i=0;i<list.length;i++){if(list[i].sku.toUpperCase()===code||list[i].sku.includes(code)){mt=list[i];break}}
    if(mt){
        setFilter('all');
        setTimeout(function(){
            var cs=document.querySelectorAll('#jobList li');
            for(var j=0;j<cs.length;j++){
                var se=cs[j].querySelector('.sk');
                if(se&&se.textContent.includes(mt.sku)){cs[j].scrollIntoView({behavior:'smooth',block:'center'});cs[j].style.background='#fefcbf';cs[j].style.borderLeft='4px solid #e53e3e';setTimeout(function(){try{cs[j].style.background='';cs[j].style.borderLeft=''}catch(e){}},3000);break}
            }
        },300);
    }else alert('\u26a0 \u672a\u627e\u5230: '+code);},100);
}
function stopScan(){
    document.getElementById('scanModal').style.display='none';
    document.getElementById('scanManual').style.display='none';
    if(scanStream){scanStream.getTracks().forEach(function(t){t.stop()});scanStream=null}
}
function scanManualSearch(){
    var c=document.getElementById('scanInput').value.trim().toUpperCase();
    if(c) sF(c);
}
</script></body></html>'''"""

# Replace the old WORKSHOP_PAGE with the new one
old_ws = content[start:end]
new_content = content[:start] + new_ws + '\n' + content[end:]

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Feature-rich workshop page restored!")
