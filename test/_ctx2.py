import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

board_idx = c.find("WORKSHOP_BOARD = '''") + 20
end_marker = c.find("\n\n'''", board_idx)
board = c[board_idx:end_marker]

sp = board.find("set_priority")
before = board[max(0,sp-600):sp+200]
print(before)
