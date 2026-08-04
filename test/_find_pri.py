with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

# Find the startJob function in the board and add setPriority after it
board_idx = c.find("WORKSHOP_BOARD = '''") + 20
end_marker = c.find("\n\n'''", board_idx)
board = c[board_idx:end_marker]

# Find startJob function
sj = board.find("function startJob(id){")
if sj >= 0:
    # Find the end of startJob (next function or closing brace pattern)
    # Look for the pattern: } \nfunction or } \n\n
    after_startjob = board[sj:]
    # Find the end of this function
    func_end = after_startjob.find("\n}\n\n")
    if func_end < 0:
        func_end = after_startjob.find("\n}\n")
    if func_end < 0:
        func_end = 400  # fallback
    
    startjob_func = after_startjob[:func_end+3]
    print("startJob function:")
    print(startjob_func[:300])

# Look for setPriority or priority-related code
for term in ["setPriority", "set_priority", "priority", "设优"]:
    idx = board.find(term)
    if idx >= 0:
        print(f"\n{term} found at {idx}:")
        print(board[max(0,idx-20):idx+100])
    else:
        print(f"\n{term}: NOT FOUND")
