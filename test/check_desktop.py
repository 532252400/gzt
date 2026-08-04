import os, subprocess

# Check junction/reparse info
result = subprocess.run('cmd /c dir "C:\\Users\\pc\\Desktop" /al', capture_output=True, text=True, shell=True)
print('Junction/reparse info for C:\\Users\\pc\\Desktop:')
print(result.stdout)
if result.stderr:
    print('STDERR:', result.stderr)

print()

# Check D:\桌面
result2 = subprocess.run('cmd /c dir "D:\\桌面" /al', capture_output=True, text=True, shell=True)
print('Junction/reparse info for D:\\桌面:')
print(result2.stdout)
if result2.stderr:
    print('STDERR:', result2.stderr)

print()

# Check what registry value Shell Folders (non-User) has
import winreg
try:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as k:
        val = winreg.QueryValueEx(k, 'Desktop')[0]
        print('Shell Folders\\Desktop:', repr(val), '->', os.path.expandvars(val))
except Exception as e:
    print('Shell Folders error:', e)

# Check alternative known folder API
try:
    import ctypes
    from ctypes import wintypes
    # SHGetKnownFolderPath
    FOLDERID_Desktop = '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}'
    # Use PowerShell to get the actual Desktop path
    ps = subprocess.run(['powershell', '-Command', '[Environment]::GetFolderPath("Desktop")'], capture_output=True, text=True)
    print('PowerShell GetFolderPath(Desktop):', ps.stdout.strip())
except Exception as e:
    print('Error:', e)

print()
print('Does D:\\桌面 have desktop.ini?', os.path.exists('D:\\桌面\\desktop.ini'))
print('Does C:\\Users\\pc\\Desktop have desktop.ini?', os.path.exists('C:\\Users\\pc\\Desktop\\desktop.ini'))
