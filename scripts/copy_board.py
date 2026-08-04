"""Copy WORKSHOP_BOARD from test to stable"""
stable_file = 'C:/Users/pc/ZCodeProject/server/工厂工作台.py'
test_file = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'

# Read both
with open(stable_file, 'r', encoding='utf-8') as f:
    stable = f.read()
with open(test_file, 'r', encoding='utf-8') as f:
    test = f.read()

# Extract WORKSHOP_BOARD from test
test_start = test.find("WORKSHOP_BOARD = '''")
test_end = test.find("\nDIAG_PAGE", test_start)
test_board = test[test_start:test_end+1]

# Replace WORKSHOP_BOARD in stable
stable_start = stable.find("WORKSHOP_BOARD = '''")
stable_end = stable.find("\nDIAG_PAGE", stable_start)
stable_board = stable[stable_start:stable_end+1]

new_stable = stable[:stable_start] + test_board + stable[stable_end+1:]
with open(stable_file, 'w', encoding='utf-8') as f:
    f.write(new_stable)

print("WORKSHOP_BOARD copied from test to stable!")
