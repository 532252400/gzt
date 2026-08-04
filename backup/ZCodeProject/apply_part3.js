const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Normalize to Unix line endings
let unix = data.replace(/\r\n/g, '\n');

// ============================================================
// 8. Add people init + error handling to fetch
// The clean base has: fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){
// We need to add people init BEFORE it, and error handling TO it
// ============================================================
const oldFetch = `fetch('/job_batches').then(function(r){return r.json()}).then(function(bs){`;
const newFetch = `var savedPpl=localStorage.getItem('default_people');if(savedPpl){document.getElementById('peopleInput').value=savedPpl;document.getElementById('peopleLabel').textContent='上次: '+savedPpl+'人';}
fetch('/job_batches').then(function(r){if(!r.ok)throw new Error('HTTP '+r.status); return r.json()}).then(function(bs){`;

let count = 0;
while (unix.includes(oldFetch)) {
    unix = unix.replace(oldFetch, newFetch);
    count++;
}
console.log('8. People init + error handling:', count);

// Restore Windows line endings
data = unix.replace(/\n/g, '\r\n');
fs.writeFileSync(path, data, 'utf8');
console.log('Part 3 done!');
