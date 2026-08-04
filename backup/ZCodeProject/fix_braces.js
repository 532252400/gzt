const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');

// Remove extra double braces before function declarations
// Pattern: "}\n}\nfunction" -> "}\nfunction"  
const old1 = "}\n}\nfunction cancelJob";
const new1 = "}\nfunction cancelJob";

let count = 0;
while (data.includes(old1)) {
    data = data.replace(old1, new1);
    count++;
}
console.log('Fixed cancelJob braces:', count);

const old2 = "}\n}\nfunction savePeople";
const new2 = "}\nfunction savePeople";

count = 0;
while (data.includes(old2)) {
    data = data.replace(old2, new2);
    count++;
}
console.log('Fixed savePeople braces:', count);

// Also check for any other double braces
const old3 = "}\n}\nfunction openComplete";
const new3 = "}\nfunction openComplete";
count = 0;
while (data.includes(old3)) {
    data = data.replace(old3, new3);
    count++;
}
console.log('Fixed openComplete braces:', count);

fs.writeFileSync(path, data, 'utf8');
console.log('Done!');
