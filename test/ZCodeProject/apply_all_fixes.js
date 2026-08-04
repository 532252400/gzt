const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// ============================================================
// 1. SERVER: Add cancel_job handler (after complete_job, before set_priority)
// ============================================================
const oldServerBlock = `                self._json({'status':'error','message':'\u274c \u65e0\u6548ID'})
                return
            if action == 'set_priority':`;

const cancelJobHandler = `                self._json({'status':'error','message':'\u274c \u65e0\u6548ID'})
                return
            if action == 'cancel_job':
                item_id = int(batch_name) if batch_name.isdigit() else 0
                if item_id:
                    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                    c.execute("UPDATE job_items SET status='pending', worker='', started_at=NULL, completed_at=NULL, completed_qty=NULL WHERE id=? AND status='processing'", (item_id,))
                    conn.commit(); conn.close()
                    self._json({'status':'ok','message':'\u2705 \u5df2\u53d6\u6d88\u52a0\u5de5'})
                    return
                self._json({'status':'error','message':'\u274c \u65e0\u6548ID'})
                return
            if action == 'set_priority':`;

data = data.replace(oldServerBlock, cancelJobHandler);
console.log('1. Server cancel_job handler: OK');

// ============================================================
// 2. CSS: Add .btn-x and .people-bar
// ============================================================
const oldCSS = `.btn-s{background:#3182ce;color:#fff}.btn-c{background:#38a169;color:#fff}.btn-d{background:#e2e8f0;color:#718096;cursor:not-allowed}`;
const newCSS = `.btn-s{background:#3182ce;color:#fff}.btn-c{background:#38a169;color:#fff}.btn-x{background:#e53e3e;color:#fff}.btn-d{background:#e2e8f0;color:#718096;cursor:not-allowed}
.people-bar{display:flex;align-items:center;gap:4px;margin:-2px 12px 10px;font-size:12px;color:#4a5568}
.people-bar input{width:56px;padding:3px 6px;border:1px solid #cbd5e0;border-radius:4px;font-size:12px;text-align:center;background:#fff}
.people-bar input:focus{outline:none;border-color:#3182ce;box-shadow:0 0 0 2px rgba(49,130,206,.15)}`;

let count = data.split(oldCSS).length - 1;
data = data.split(oldCSS).join(newCSS);
console.log('2. CSS added (' + count + ' templates): OK');

// ============================================================
// 3. WORKSHOP HTML: Add people bar after filter bar
// ============================================================
const filterBar = `<div class="filter-bar"><button id="f_all" class="on" onclick="setFilter('all')">\ud83d\udccb \u5168\u90e8</button><button id="f_priority" onclick="setFilter('priority')">\u2b50 \u4f18\u5148</button></div>`;
const filterBarWithPeople = filterBar + `\n<div class="people-bar"><span>\ud83d\udc65</span><input id="peopleInput" type="number" min="1" placeholder="\u4eba\u6570" onchange="savePeople()"><span style="font-size:11px;color:#999" id="peopleLabel">\u4e0a\u6b21: -</span></div>`;

data = data.replace(filterBar, filterBarWithPeople);
console.log('3. WORKSHOP_PAGE people bar: OK');

// ============================================================
// 4. WORKSHOP_ADMIN HTML: Add people bar after admin filter bar
// ============================================================
const adminFilterBar = `<div class="filter-bar"><button id="af_all" class="on" onclick="setFilter('all')">\ud83d\udccb \u5168\u90e8</button><button id="af_priority" onclick="setFilter('priority')">\u2b50 \u4f18\u5148</button></div>`;
const adminFilterWithPeople = adminFilterBar + `\n<div class="people-bar"><span>\ud83d\udc65</span><input id="peopleInput" type="number" min="1" placeholder="\u4eba\u6570" onchange="savePeople()"><span style="font-size:11px;color:#999" id="peopleLabel">\u4e0a\u6b21: -</span></div>`;

data = data.replace(adminFilterBar, adminFilterWithPeople);
console.log('4. WORKSHOP_ADMIN people bar: OK');

// ============================================================
// 5. Fix openComplete button: data-id/data-sku + cancel button
// ============================================================
const oldBtn = `onclick="openComplete('+i.id+',\\''+i.sku+'\\')">\u2714 \u5b8c\u6210\u52a0\u5de5</button>';}`;
const newBtn = `onclick="openComplete(this)">\u2714 \u5b8c\u6210\u52a0\u5de5</button><button class="btn-x" onclick="cancelJob('+i.id+')">\u2716 \u53d6\u6d88</button>';}`;

count = data.split(oldBtn).length - 1;
data = data.split(oldBtn).join(newBtn);
console.log('5. Cancel button added (' + count + ' templates): OK');

// ============================================================
// 6. Fix openComplete function signature
// ============================================================
const oldFunc = `function openComplete(id,sku){curItemId=id;document.getElementById('mSku').textContent=sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}`;
const newFunc = `function openComplete(btn){curItemId=btn.dataset.id;document.getElementById('mSku').textContent=btn.dataset.sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}`;

count = data.split(oldFunc).length - 1;
data = data.split(oldFunc).join(newFunc);
console.log('6. openComplete function (' + count + ' templates): OK');

// ============================================================
// 7. Replace startJob and add savePeople + cancelJob
// ============================================================
// The old startJob pattern (using node escape for Chinese chars)
const oldStartJob = `async function startJob(id){
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker','\u8f66\u95f4\u5de5\u4eba');
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u2705');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}`;

const newStartJobSaveCancel = `async function startJob(id){
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
    if(v&&parseInt(v)>=1){localStorage.setItem('default_people',v);document.getElementById('peopleLabel').textContent='\u4e0a\u6b21: '+v+'\u4eba';}
}
async function cancelJob(id){
    if(!confirm('\u786e\u5b9a\u53d6\u6d88\u6b64\u52a0\u5de5\u4efb\u52a1\uff1f'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('\u2705');loadJobs();}else alert('\u274c '+d.message);}
    catch(e){alert('\u274c '+e.message);}
}`;

count = data.split(oldStartJob).length - 1;
data = data.split(oldStartJob).join(newStartJobSaveCancel);
console.log('7. startJob/savePeople/cancelJob (' + count + ' templates): OK');

// ============================================================
// 8. Add people init code before fetch('/job_batches')
// ============================================================
const oldFetch = `fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;
const newFetch = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='\u4e0a\u6b21: '+savedPpl+'\u4eba';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;

count = data.split(oldFetch).length - 1;
data = data.split(oldFetch).join(newFetch);
console.log('8. People init (' + count + ' templates): OK');

// ============================================================
// Write back
// ============================================================
fs.writeFileSync(path, data, 'utf8');
console.log('\n=== ALL FIXES APPLIED SUCCESSFULLY ===');
