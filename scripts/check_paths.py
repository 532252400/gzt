# Check server paths
import os, sys
sys.path.insert(0, 'C:/Users/pc/ZCodeProject/server')

# Check what the server sees as DESKTOP
import winreg
with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as k:
    desktop = os.path.expandvars(winreg.QueryValueEx(k, 'Desktop')[0])
print(f"Registry Desktop: {desktop}")
print(f"Path exists: {os.path.exists(desktop)}")

# Check upload dir
upload_dir = os.path.join(desktop, '_工作台上传')
print(f"Upload dir: {upload_dir}")
print(f"Upload dir exists: {os.path.exists(upload_dir)}")

# Try to create it
try:
    os.makedirs(upload_dir, exist_ok=True)
    print(f"Created: {upload_dir}")
    print(f"Now exists: {os.path.exists(upload_dir)}")
except Exception as e:
    print(f"Error creating: {e}")

# Check permissions
print(f"\nDesktop writable: {os.access(desktop, os.W_OK)}")
print(f"Desktop readable: {os.access(desktop, os.R_OK)}")
