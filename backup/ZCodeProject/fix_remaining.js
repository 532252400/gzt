const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Windows line endings are \r\n
const old1 = "}\r\nfunction openComplete(btn)";
const new1 = "}\r\n\tasync function cancelJob(id){\r\n\t    if(!confirm('确定取消此加工任务？'))return;\r\n\t    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);\r\n\t    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}\r\n\t    catch(e){alert('❌ '+e.message);}\r\n\t}\r\nfunction openComplete(btn)";

let count = 0;
let result = data;
while (result.includes(old1)) {
    result = result.replace(old1, new1);
    count++;
}
console.log('Replacements:', count);

if (count === 0) {
    console.log('Pattern not found with \\r\\n, trying \\n...');
    const old2 = "}\nfunction openComplete(btn)";
    const new2 = "}\n\tasync function cancelJob(id){\n\t    if(!confirm('确定取消此加工任务？'))return;\n\t    var fd=new FormData();fd.append('action','cancel_job');fd.append('batch_name',id);\n\t    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok'){alert('✅');loadJobs();}else alert('❌ '+d.message);}\n\t    catch(e){alert('❌ '+e.message);}\n\t}\nfunction openComplete(btn)";
    count = 0;
    result = data;
    while (result.includes(old2)) {
        result = result.replace(old2, new2);
        count++;
    }
    console.log('Replacements with \\n:', count);
}

fs.writeFileSync(path, result, 'utf8');
console.log('Done. Count:', count);
