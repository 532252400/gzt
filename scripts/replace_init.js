const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Normalize line endings
let unix = data.replace(/\r\n/g, '\n');

// Wait for people init / load / auto-refresh
// Match from "var savedPpl" to the end of the last setInterval
// Then replace with simplified version

// The pattern in the workshop page (after confirmComplete ends)
const oldEnd = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){
    var sel=document.getElementById('batchSel');
    bs.forEach(function(b){var o=document.createElement('option');o.value=b.id;o.textContent=b.name+' · '+b.count+'项';sel.appendChild(o);});
    if(bs.length){document.getElementById('batchInfo').textContent='当前: '+bs[0].name;document.getElementById('batchSel').value=bs[0].id;loadJobs();}
    else document.getElementById('batchInfo').textContent='暂无加工单';
});
// Auto-refresh: check new batches every 15s
setInterval(function(){
    var sel=document.getElementById('batchSel');
    if(!sel)return;
    fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
        if(bs.length>sel.options.length){
            sel.innerHTML='';
            bs.forEach(function(b){var o=document.createElement('option');o.value=b.id;o.textContent=b.name+' · '+b.count+'项';sel.appendChild(o);});
            if(bs.length){document.getElementById('batchInfo').textContent='当前: '+bs[0].name;sel.value=bs[0].id;loadJobs();}
        }
    }).catch(function(){});`;

const newEnd = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
document.getElementById('batchInfo').textContent='全部加工单';
loadJobs();
// Auto-refresh every 15s
setInterval(function(){loadJobs();}, 15000);`;

let count = 0;
let result = unix;
while (result.includes(oldEnd)) {
    result = result.replace(oldEnd, newEnd);
    count++;
}
console.log('Workshop page replaced:', count);

// For admin page - it uses a different pattern (loadBatchList etc.)
// The admin might have a different ending 
const oldAdminEnd = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){
    var sel=document.getElementById('batchSel');
    bs.forEach(function(b){var o=document.createElement('option');o.value=b.id;o.textContent=b.name+' · '+b.count+'项';sel.appendChild(o);});
    if(bs.length){document.getElementById('batchInfo').textContent='当前: '+bs[0].name;document.getElementById('batchSel').value=bs[0].id;loadJobs();}
    else document.getElementById('batchInfo').textContent='暂无加工单';
});
// Batch list
function loadBatchList(){
    fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
        var div=document.getElementById('batchList');
        if(!div)return;
        var h='';
        bs.forEach(function(b){h+='<div style=\"padding:2px 0;border-bottom:1px solid #eee;display:flex;justify-content:space-between\"><span>'+b.name+'</span><span style=\"color:#999\">'+b.count+'项</span></div>';});
        div.innerHTML=h||'<span style=\"color:#999\">暂无批次</span>';
    }).catch(function(){});
}
loadBatchList();`;

const newAdminEnd = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
document.getElementById('batchInfo').textContent='全部加工单';
loadJobs();
// Batch list
function loadBatchList(){
    fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
        var div=document.getElementById('batchList');
        if(!div)return;
        var h='';
        bs.forEach(function(b){h+='<div style=\"padding:2px 0;border-bottom:1px solid #eee;display:flex;justify-content:space-between\"><span>'+b.name+'</span><span style=\"color:#999\">'+b.count+'项</span></div>';});
        div.innerHTML=h||'<span style=\"color:#999\">暂无批次</span>';
    }).catch(function(){});
}
loadBatchList();`;

while (result.includes(oldAdminEnd)) {
    result = result.replace(oldAdminEnd, newAdminEnd);
    count++;
}
console.log('Admin page replaced:', count);

// Restore Windows line endings
data = result.replace(/\n/g, '\r\n');
fs.writeFileSync(path, data, 'utf8');
console.log('Total replaced:', count);
