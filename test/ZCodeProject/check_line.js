const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
const data = fs.readFileSync(path, 'utf8');
const idx = data.indexOf('function openComplete(btn)');
console.log('Position:', idx);
console.log('Previous 80 chars:', JSON.stringify(data.substring(Math.max(0,idx-80), idx)));
// Count occurrences
let count = 0, pos = -1;
while ((pos = data.indexOf('function openComplete(btn)', pos+1)) !== -1) count++;
console.log('Found', count, 'occurrences');
// Show line endings around the area
const before = data.substring(idx-100, idx);
console.log('Line endings:', JSON.stringify(before.match(/\r?\n/g)));
