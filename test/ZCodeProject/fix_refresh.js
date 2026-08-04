const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Simple string replacement by finding unique parts
// Workshop: find the auto-refresh line after loadJobs();
const wsFind = 'setInterval(function(){loadJobs();}, 3000);';
const wsReplace = 'function autoRefresh(){loadJobs();setTimeout(autoRefresh,2000);}\nautoRefresh();';

let idx = data.indexOf(wsFind);
if (idx >= 0) {
    // Find the start of the line to remove the comment too
    const lineStart = data.lastIndexOf('\n', idx) + 1;
    const beforeLine = data.substring(lineStart, idx);
    // Remove the entire line with comment + setInterval line
    const endOfLine = data.indexOf('\n', idx + wsFind.length);
    const toRemove = data.substring(lineStart, endOfLine + 1);
    data = data.replace(toRemove, '\n' + wsReplace + '\n');
    console.log('Workshop fixed');
} else {
    console.log('Workshop pattern not found');
}

// Admin: find the setInterval with loadBatchList
const admFind = 'setInterval(function(){loadJobs();loadBatchList();}, 5000);';
const admReplace = 'function autoRefresh(){loadJobs();loadBatchList();setTimeout(autoRefresh,3000);}\nautoRefresh();';

idx = data.indexOf(admFind);
if (idx >= 0) {
    const lineStart = data.lastIndexOf('\n', idx) + 1;
    const beforeLine = data.substring(lineStart, idx);
    const endOfLine = data.indexOf('\n', idx + admFind.length);
    const toRemove = data.substring(lineStart, endOfLine + 1);
    data = data.replace(toRemove, '\n' + admReplace + '\n');
    console.log('Admin fixed');
} else {
    console.log('Admin pattern not found');
}

fs.writeFileSync(path, data, 'utf8');
console.log('DONE!');
