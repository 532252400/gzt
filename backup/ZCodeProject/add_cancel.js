const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// The cancelJob function (no indent at top level, same as startJob)
const cancelFunc = 
`}
async function cancelJob(id){
    if(!confirm('确定取消此加工任务？'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('OK');loadJobs();}else alert('FAIL');}
    catch(e){alert('FAIL');}
}
function openComplete(btn)`;

let count = 0;
while (data.includes('function openComplete(btn)')) {
    data = data.replace('function openComplete(btn)', cancelFunc);
    count++;
}

console.log('Inserted', count, 'cancelJob functions');
fs.writeFileSync(path, data, 'utf8');
console.log('Saved!');
