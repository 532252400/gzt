const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Show context around 'function savePeople'  
const idx = data.indexOf('function savePeople');
if (idx >= 0) {
    console.log('Before savePeople:', JSON.stringify(data.substring(idx-40, idx)));
    // Check what comes before
    const before = data.substring(idx-10, idx);
    const chars = [];
    for (let i = 0; i < before.length; i++) chars.push(before.charCodeAt(i));
    console.log('Char codes:', chars.join(','));
}

// Also check cancelJob
const idx2 = data.indexOf('async function cancelJob');
if (idx2 >= 0) {
    console.log('Before cancelJob:', JSON.stringify(data.substring(idx2-20, idx2)));
}

// Now do the fix properly
// Remove extra brace before cancelJob  
let re1 = /}\s*}\s*function cancelJob/g;
let m1 = data.match(re1);
console.log('Extra brace before cancelJob:', m1 ? m1.length : 0);

let re2 = /}\s*}\s*function savePeople/g;
let m2 = data.match(re2);
console.log('Extra brace before savePeople:', m2 ? m2.length : 0);

// Fix them
data = data.replace(/}\s*}\s*function cancelJob/g, '}\nfunction cancelJob');
data = data.replace(/}\s*}\s*function savePeople/g, '}\nfunction savePeople');

// Also remove extra brace before openComplete
data = data.replace(/}\s*}\s*function openComplete/g, '}\nfunction openComplete');

fs.writeFileSync(path, data, 'utf8');
console.log('Fixed!');
