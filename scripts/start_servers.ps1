# Start both factory workbench servers silently
$python = "C:\Users\pc\AppData\Local\Python\bin\python3.exe"
$serverPy = "C:\Users\pc\ZCodeProject\server\工厂工作台.py"
$testPy = "C:\Users\pc\ZCodeProject\test\工厂工作台.py"

# Start stable version (8932)
Start-Process -FilePath $python -ArgumentList "-X utf8", "`"$serverPy`"" -WindowStyle Hidden

# Start test version (8933)
Start-Process -FilePath $python -ArgumentList "-X utf8", "`"$testPy`"" -WindowStyle Hidden
