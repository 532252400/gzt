with open(r"C:\Users\pc\ZCodeProject\server\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

board_idx = c.find("WORKSHOP_BOARD = '''") + 20
end_marker = c.find("\n\n'''", board_idx)
if end_marker < 0:
    end_marker = c.find("\n'''", board_idx)
board = c[board_idx:end_marker]

print(f"Stable BOARD: {len(board)} chars")
print("setPriority:", "setPriority" in board)
print("startJob:", "startJob" in board)
print("peopleInput:", "peopleInput" in board)
print("Title:", board[board.find("<title>")+7:board.find("</title>")])
