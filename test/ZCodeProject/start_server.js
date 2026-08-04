const { spawn } = require('child_process');
const p = spawn('C:/Users/pc/AppData/Local/Python/bin/python3.exe', 
    ['-X', 'utf8', 'D:/桌面/工厂工作台.py'], 
    { stdio: 'ignore', detached: true });
p.unref();
console.log('Started PID:', p.pid);
// Keep running for a bit to let it start
setTimeout(() => process.exit(0), 2000);
