const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Replace each occurrence of 'function openComplete(btn)' with cancelJob + it
// Using a different marker to avoid infinite loop
const marker = '___OPENCOMPLETE_MARKER___';
data = data.split('function openComplete(btn)').join(marker);

const cancelFunc = 
`}
async function cancelJob(id){
    if(!confirm('确定取消此加工任务？'))return;
    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('OK');loadJobs();}else alert('FAIL');}
    catch(e){alert('FAIL');}
}
function openComplete(btn)`;

data = data.split(marker).join(cancelFunc);

fs.writeFileSync(path, data, 'utf8');
console.log('Done!');
