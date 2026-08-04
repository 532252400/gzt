with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

# Find and fix the broken togglePri function
old_toggle = """async function togglePri(id,pri){
    var fd=new FormData();fd.append('action','set_priority');fd.append('batch_name',id);fd.append('priority',pri);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok')loadBoard();else alert(d.message)"""

new_toggle = """async function togglePri(id,pri){
    var fd=new FormData();fd.append('action','set_priority');fd.append('batch_name',id);fd.append('priority',pri);
    try{var r=await fetch('/run',{method:'POST',body:fd});var d=await r.json();if(d.status==='ok')loadBoard();else alert(d.message);
    }catch(e){alert('❌ '+e.message);}
}"""

if old_toggle in c:
    c = c.replace(old_toggle, new_toggle)
    print("Fixed togglePri function")
    with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "w", encoding="utf-8") as f:
        f.write(c)
    print("Saved, size:", len(c))
else:
    print("NOT FOUND - may already be fixed or different format")
    # Search for togglePri
    idx = c.find("togglePri")
    if idx >= 0:
        print("togglePri at", idx, c[idx:idx+350])
