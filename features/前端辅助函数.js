/* 前端通用辅助函数 - 所有功能共用 */
function f(id) { document.getElementById(id).click() }

function fs(fid, tid, bid) {
    document.getElementById(tid).textContent = '📄 ' + document.getElementById(fid).files[0].name;
    document.getElementById(bid).disabled = false;
}

function s(id, c, m) {
    var r = document.getElementById(id);
    r.className = 'rs ' + c;
    r.style.display = 'block';
    r.textContent = m;
}

async function go(a) {
    var m = {
        lbl100: { f: 'f1', b: 'b1', r: 'r1' },
        lbl30:  { f: 'f2', b: 'b2', r: 'r2' },
        us:     { f: 'f3', b: 'b3', r: 'r3' },
        ca:     { f: 'f4', b: 'b4', r: 'r4' },
        rc:     { f: 'f5', b: 'b5', r: 'r5' }
    }[a];
    var fi = document.getElementById(m.f);
    var btn = document.getElementById(m.b);
    var res = document.getElementById(m.r);
    
    if (!fi || !btn || !res) return;
    if (!fi.files || !fi.files[0]) return;
    
    btn.disabled = true;
    btn.innerHTML = '<span class="ld"></span>...';
    s(res, 'wa', '⏳...');
    
    var fd = new FormData();
    fd.append('file', fi.files[0]);
    fd.append('action', a);
    
    try {
        var r = await fetch('/run', { method: 'POST', body: fd });
        var d = await r.json();
        s(res, d.status === 'ok' ? 'ok' : 'er',
          d.status === 'ok' ? d.message : '❌ ' + d.message);
    } catch (e) {
        s(res, 'er', '❌ ' + e.message);
    }
    
    btn.disabled = false;
    btn.textContent = a === 'rc' ? '开始核对' :
                      (a === 'us' || a === 'ca' ? '生成汇总' : '生成标签');
}
