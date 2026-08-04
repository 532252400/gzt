with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

board_idx = c.find("WORKSHOP_BOARD = '''") + 20
end_marker = c.find("\n\n'''", board_idx)
board = c[board_idx:end_marker]

# Find set_priority and get surrounding function
sp = board.find("set_priority")
if sp >= 0:
    # Get 600 chars before to find function name
    before = board[max(0,sp-600):sp+200]
    print(before)
