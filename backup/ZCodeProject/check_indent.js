const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
const d = fs.readFileSync(path, 'utf8');
const lines = d.split(/\r?\n/);
for (let i = 1085; i <= 1095 && i <= lines.length; i++) {
    const line = lines[i-1] || '(empty)';
    const spaces = (line.match(/^ */) || [''])[0].length;
    const tabs = (line.match(/^\t*/) || [''])[0].length;
    console.log(i + ': spaces=' + spaces + ' tabs=' + tabs + ' len=' + line.length + ' |' + line.substring(0, 60) + '|');
}
