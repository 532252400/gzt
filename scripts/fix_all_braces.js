const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// More aggressive: replace any sequence of braces+whitespace before 
// function declarations with a single "}\n"
const re = /(?:}\s*){2,}(?=function\s+(?:cancelJob|savePeople|openComplete))/g;
data = data.replace(re, '}\n');

// Count braces before each function
['cancelJob', 'savePeople', 'openComplete'].forEach(fn => {
    const re2 = new RegExp('function\\s+' + fn, 'g');
    let m;
    while ((m = re2.exec(data)) !== null) {
        const before = data.substring(m.index - 20, m.index);
        const braces = (before.match(/}/g) || []).length;
        console.log(fn + ' at ' + m.index + ': ' + braces + ' braces before');
    }
});

fs.writeFileSync(path, data, 'utf8');
console.log('Done!');
