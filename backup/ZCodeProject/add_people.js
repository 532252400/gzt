const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Replace startJob in both templates
const oldStartJob = `async function startJob(id){
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker','车间工人');
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}`;

const newStartJob = `async function startJob(id){
    var ppl=document.getElementById('peopleInput')?document.getElementById('peopleInput').value:'';
    if(!ppl||parseInt(ppl)<1)ppl='1';
    localStorage.setItem('default_people',ppl);
    var wn='车间工人 x'+ppl;
    var fd=new FormData();fd.append('action','start_job');fd.append('batch_name',id);fd.append('worker',wn);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}
    catch(e){alert('❌ '+e.message);}
}`;

let count = 0;
let result = data;
while (result.includes(oldStartJob)) {
    result = result.replace(oldStartJob, newStartJob);
    count++;
}
console.log('startJob replaced:', count);

// Add savePeople function before cancelJob (replace cancelJob marker)
// Find cancelJob function
const cancelFunc = `async function cancelJob(id){`;
// Add savePeople before each cancelJob
const withPeople = `function savePeople(){
    var v=document.getElementById('peopleInput').value;
    if(v&&parseInt(v)>=1){localStorage.setItem('default_people',v);document.getElementById('peopleLabel').textContent='上次: '+v+'人';}
}
async function cancelJob(id){`;

count = 0;
result = result.split(cancelFunc).join('__MARKER__');
result = result.split('__MARKER__').join(withPeople);
// The first occurrence would have been before the first cancelJob
// But split/join replaces ALL occurrences
console.log('savePeople inserted');

// Initial load of people count at the bottom of script
// Before the fetch, add people loading code
const oldInit = `fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;
const newInit = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;

count = 0;
result = result.split(oldInit).join(newInit);
console.log('Init updated');

fs.writeFileSync(path, result, 'utf8');
console.log('Done!');
