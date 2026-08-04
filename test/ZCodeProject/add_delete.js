const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');
let unix = data.replace(/\r\n/g, '\n');

// ============================================================
// 1. Server handler: add delete_jobs action
// ============================================================
const serverOld = `            if action == 'cancel_job':
                item_id = int(batch_name) if batch_name.isdigit() else 0
                if item_id:
                    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                    c.execute("UPDATE job_items SET status='pending', worker='', started_at=NULL, completed_at=NULL, completed_qty=NULL WHERE id=? AND status='processing'", (item_id,))
                    conn.commit(); conn.close()
                    self._json({'status':'ok','message':'\u2705 \u5df2\u53d6\u6d88\u52a0\u5de5\uff0c\u53ef\u91cd\u65b0\u5f00\u59cb'})
                    return
                self._json({'status':'error','message':'\u274c \u65e0\u6548ID'})
                return
            if action == 'set_priority':`;

const serverNew = `            if action == 'cancel_job':
                item_id = int(batch_name) if batch_name.isdigit() else 0
                if item_id:
                    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                    c.execute("UPDATE job_items SET status='pending', worker='', started_at=NULL, completed_at=NULL, completed_qty=NULL WHERE id=? AND status='processing'", (item_id,))
                    conn.commit(); conn.close()
                    self._json({'status':'ok','message':'\u2705 \u5df2\u53d6\u6d88\u52a0\u5de5\uff0c\u53ef\u91cd\u65b0\u5f00\u59cb'})
                    return
                self._json({'status':'error','message':'\u274c \u65e0\u6548ID'})
                return
            if action == 'delete_jobs':
                ids_str = self._get_post('ids', '')
                if ids_str:
                    ids = [int(x) for x in ids_str.split(',') if x.strip().isdigit()]
                    if ids:
                        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                        placeholders = ','.join(['?'] * len(ids))
                        c.execute('DELETE FROM job_items WHERE id IN (' + placeholders + ')', ids)
                        conn.commit(); conn.close()
                        self._json({'status':'ok','message':'\u2705 \u5df2\u5220\u9664 ' + str(len(ids)) + ' \u9879'})
                        return
                self._json({'status':'error','message':'\u274c \u65e0\u6548\u7684ID\u5217\u8868'})
                return
            if action == 'set_priority':`;

unix = unix.replace(serverOld, serverNew);
console.log('1. Server delete_jobs handler: OK');

// ============================================================
// 2. ADMIN CSS: Add delete bar styles (before .ov)
// ============================================================
const cssOld = `.btn-pri-off{background:#fff;color:#e53e3e;border:1px solid #e53e3e!important;flex:0!important;padding:7px 9px!important}</style>`;
const cssNew = `.btn-pri-off{background:#fff;color:#e53e3e;border:1px solid #e53e3e!important;flex:0!important;padding:7px 9px!important}
.del-bar{display:flex;align-items:center;gap:6px;margin:4px 12px 6px;padding:6px 8px;background:#fff;border-radius:6px;font-size:11px;border:1px solid #fed7d7}
.del-bar label{display:flex;align-items:center;gap:3px;cursor:pointer;color:#4a5568}
.del-bar label input{cursor:pointer}
.del-btn{background:#e53e3e;color:#fff;border:none;border-radius:4px;padding:4px 10px;font-size:11px;cursor:pointer;margin-left:auto}
.del-btn:hover{background:#c53030}
.cb-item{width:14px;height:14px;cursor:pointer;margin-right:4px;flex-shrink:0}</style>`;

unix = unix.replace(cssOld, cssNew);
console.log('2. Admin CSS: OK');

// ============================================================
// 3. ADMIN HTML: Add delete bar after batch list
// ============================================================
const htmlOld = `</div>
<ul class="li" id="jobList"></ul>
<div class="ov" id="modal"><div class="bx"><h3 id="mTitle">完成加工</h3><p style="font-size:12px;color:#666;margin-bottom:6px" id="mSku"></p><input id="mQty" type="number" placeholder="输入完成数量" min="1" onkeydown="if(event.key==='Enter')confirmComplete()"><button class="btn-s" onclick="confirmComplete()">确定完成</button><button onclick="closeModal()" style="background:#e2e8f0">取消</button></div></div>
<p class="cr"><a class="bt" href="/workshop">车间页面</a> | <a class="bt" href="/">工作台</a></p>`;

const htmlNew = `</div>
<div class="del-bar"><label><input type="checkbox" id="selectAll" onchange="toggleAll()"> 全选</label><span id="selCount" style="color:#999">0项</span><button class="del-btn" onclick="deleteSelected()">🗑 删除选中</button></div>
<ul class="li" id="jobList"></ul>
<div class="ov" id="modal"><div class="bx"><h3 id="mTitle">完成加工</h3><p style="font-size:12px;color:#666;margin-bottom:6px" id="mSku"></p><input id="mQty" type="number" placeholder="输入完成数量" min="1" onkeydown="if(event.key==='Enter')confirmComplete()"><button class="btn-s" onclick="confirmComplete()">确定完成</button><button onclick="closeModal()" style="background:#e2e8f0">取消</button></div></div>
<p class="cr"><a class="bt" href="/workshop">车间页面</a> | <a class="bt" href="/">工作台</a></p>`;

unix = unix.replace(htmlOld, htmlNew);
console.log('3. Admin HTML delete bar: OK');

// ============================================================
// 4. ADMIN JS: Add checkbox to each item + deleteSelected function
// ============================================================
// Update li.innerHTML to include checkbox in the top div
const itemOld = `li.innerHTML='<div class="top"><span class="st '+stClass+'"></span><span class="sk">'+i.sku+priBadge+'</span><span style="font-size:10px;color:#999">'+statusLabel[i.status]+'</span></div>'+'<div class="nm">'+(i.name||'')+'</div>'+info+'<div class="btns">'+priBtn+btns+'</div>';`;

const itemNew = `li.innerHTML='<div class="top"><input type="checkbox" class="cb-item" data-id="'+i.id+'" onchange="updateSelCount()"><span class="st '+stClass+'"></span><span class="sk">'+i.sku+priBadge+'</span><span style="font-size:10px;color:#999">'+statusLabel[i.status]+'</span></div>'+'<div class="nm">'+(i.name||'')+'</div>'+info+'<div class="btns">'+priBtn+btns+'</div>';`;

unix = unix.replace(itemOld, itemNew);
console.log('4. Item checkbox: OK');

// ============================================================
// 5. ADMIN JS: Add deleteSelected + toggleAll + updateSelCount functions
// ============================================================
const funcOld = `async function cancelJob(id){
    if(!confirm('确定取消此加工任务？'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}
function openComplete(btn)`;  // Using partial match since the admin page is what we want

const funcNew = `async function cancelJob(id){
    if(!confirm('确定取消此加工任务？'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}
function toggleAll(){
    var checked=document.getElementById('selectAll').checked;
    document.querySelectorAll('.cb-item').forEach(function(c){c.checked=checked;});
    updateSelCount();
}
function updateSelCount(){
    var n=document.querySelectorAll('.cb-item:checked').length;
    document.getElementById('selCount').textContent=n+'项';
}
function deleteSelected(){
    var ids=[];
    document.querySelectorAll('.cb-item:checked').forEach(function(c){ids.push(c.getAttribute('data-id'));});
    if(!ids.length){alert('请先选择要删除的项');return;}
    if(!confirm('确定删除选中的 '+ids.length+' 项？此操作不可撤销！'))return;
    var fd=new FormData();fd.append('action','delete_jobs');fd.append('ids',ids.join(','));
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert(d.message);loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}
function openComplete(btn)`;

// Only replace in admin context (the second occurrence)
// Find all occurrences and replace just the LAST one (admin page)
let count = 0;
let result = unix;
while (result.includes(funcOld)) {
    result = result.replace(funcOld, funcNew);
    count++;
}
console.log('5. Delete functions replaced:', count);

// Restore line endings
data = result.replace(/\n/g, '\r\n');
fs.writeFileSync(path, data, 'utf8');
console.log('DONE!');
