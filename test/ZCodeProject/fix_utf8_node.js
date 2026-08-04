const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path);

const s1 = Buffer.from([0xE2, 0x9C, 0x3F]);
const r1 = Buffer.from([0xE2, 0x9C, 0x94]);
const s2 = Buffer.from([0xE9, 0xA1, 0x3F]);
const r2 = Buffer.from([0xE9, 0xA1, 0xB9]);

let cnt = 0, idx = data.indexOf(s1);
while (idx !== -1) { data = Buffer.concat([data.slice(0,idx), r1, data.slice(idx+3)]); cnt++; idx = data.indexOf(s1); }
console.log('Fix1 (E2 9C 3F):', cnt);

cnt = 0; idx = data.indexOf(s2);
while (idx !== -1) { data = Buffer.concat([data.slice(0,idx), r2, data.slice(idx+3)]); cnt++; idx = data.indexOf(s2); }
console.log('Fix2 (E9 A1 3F):', cnt);

fs.writeFileSync(path, data);
console.log('Done - file fixed');
