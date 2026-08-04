const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Normalize line endings temporarily for matching
const unix = data.replace(/\r\n/g, '\n');

// ============================================================
// 7. Replace startJob and add savePeople + cancelJob
// ============================================================
const oldStartJob = 
`async function startJob(id){
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker','车间工人');
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}`;

const newBlock = 
`async function startJob(id){
    var ppl=document.getElementById('peopleInput')?document.getElementById('peopleInput').value:'';
    if(!ppl||parseInt(ppl)<1)ppl='1';
    localStorage.setItem('default_people',ppl);
    var wn='车间工人 x'+ppl;
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker',wn);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}
function savePeople(){
    var v=document.getElementById('peopleInput').value;
    if(v&&parseInt(v)>=1){localStorage.setItem('default_people',v);document.getElementById('peopleLabel').textContent='上次: '+v+'人';}
}
async function cancelJob(id){
    if(!confirm('确定取消此加工任务？'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}`;

let count = 0;
let result = unix;
while (result.includes(oldStartJob)) {
    result = result.replace(oldStartJob, newBlock);
    count++;
}
console.log('7. startJob/savePeople/cancelJob replaced:', count);

// ============================================================
// 8. Add people init code before fetch
// ============================================================
const oldFetch = `fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;
const newFetch = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;

let count2 = 0;
while (result.includes(oldFetch)) {
    result = result.replace(oldFetch, newFetch);
    count2++;
}
console.log('8. People init replaced:', count2);

// Restore Windows line endings
data = result.replace(/\n/g, '\r\n');

fs.writeFileSync(path, data, 'utf8');
console.log('Part 2 done!');
