"""Directly test run_rc from the test module"""
import sys
sys.path.insert(0, 'C:/Users/pc/ZCodeProject/test')

# Import the function directly from the test server file
import importlib.util
spec = importlib.util.spec_from_file_location("factory_test", "C:/Users/pc/ZCodeProject/test/工厂工作台.py")

# We need to set up the DB_PATH first
import os
os.environ['DB_PATH'] = 'C:/Users/pc/ZCodeProject/test/_scan_data.db'

mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

fp = r'D:\桌面\7-8Receiving-Record-20260708-933381695180312576.xlsx'
result = mod.run_rc(fp)
print(result)
