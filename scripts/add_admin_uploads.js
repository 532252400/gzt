const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// ============================================================
// 1. ADMIN PAGE: Add batch list section between people-bar and jobList
// ============================================================
const adminTarget = `<div class="people-bar"><span>👥</span><input id="peopleInput" type="number" min="1" placeholder="人数" onchange="savePeople()"><span style="font-size:11px;color:#999" id="peopleLabel">上次: -</span></div>
<ul class="li" id="jobList"></ul>`;

const adminNew = `<div class="people-bar"><span>👥</span><input id="peopleInput" type="number" min="1" placeholder="人数" onchange="savePeople()"><span style="font-size:11px;color:#999" id="peopleLabel">上次: -</span></div>
<div style="margin:6px 12px;padding:8px;background:#fff;border-radius:8px;font-size:11px;box-shadow:0 1px 3px rgba(0,0,0,.06)">
<div style="font-weight:600;font-size:12px;margin-bottom:6px">📦 已上传批次</div>
<div id="batchList" style="max-height:120px;overflow-y:auto"></div>
</div>
<ul class="li" id="jobList"></ul>`;

data = data.replace(adminTarget, adminNew);
console.log('1. Admin batch list section: OK');

// ============================================================
// 2. ADMIN JS: Add batch list loading + auto-refresh timer
// ============================================================
// Find the end of admin script (before </script>)
// Add batch list loading after the existing fetch chain
const adminEndJS = `});\n</script></body></html>'''\nWORKSHOP_PAGE`;

const adminBatchListJS = `});
// Load batch list
function loadBatchList(){
    fetch('/job_batches').then(r=>r.json()).then(bs=>{
        var div=document.getElementById('batchList');
        if(!div)return;
        div.innerHTML=bs.map(function(b){return '<div style="padding:2px 0;border-bottom:1px solid #eee;display:flex;justify-content:space-between"><span>'+b.name+'</span><span style="color:#999">'+b.count+'项</span></div>';}).join('');
    }).catch(function(){});
}
loadBatchList();
</script></body></html>'''
WORKSHOP_PAGE`;

data = data.replace(adminEndJS, adminBatchListJS);
console.log('2. Admin batch list JS: OK');

// ============================================================
// 3. WORKSHOP PAGE: Add auto-refresh (poll every 15s)
// ============================================================
// Add a refresh timer at the end of workshop script (before </script>)
const wsEndJS = `});\n</script></body></html>'''\nWORKSHOP_ADMIN`;

const wsRefreshJS = `});
// Auto-refresh: check for new batches every 15 seconds
setInterval(function(){
    var sel=document.getElementById('batchSel');
    if(!sel)return;
    var curVal=sel.value;
    fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
        var oldLen=sel.options.length;
        // Check if new batches appeared
        if(bs.length>oldLen){
            // Rebuild select
            sel.innerHTML='';
            bs.forEach(function(b){var o=document.createElement('option');o.value=b.id;o.textContent=b.name+' · '+b.count+'项';sel.appendChild(o);});
            if(bs.length){document.getElementById('batchInfo').textContent='当前: '+bs[0].name;sel.value=bs[0].id;loadJobs();}
        }
    }).catch(function(){});
}, 15000);
</script></body></html>'''
WORKSHOP_ADMIN`;

data = data.replace(wsEndJS, wsRefreshJS);
console.log('3. Workshop auto-refresh: OK');

fs.writeFileSync(path, data, 'utf8');
console.log('DONE!');
