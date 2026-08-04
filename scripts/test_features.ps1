# Test the label/summary/receiving features
$baseUrl = "http://127.0.0.1:8933"
$testFile = "C:\Users\pc\ZCodeProject\data\test_features.xlsx"
$python = "C:\Users\pc\AppData\Local\Python\bin\python3.exe"
$script = "C:\Users\pc\ZCodeProject\scripts\test_all_features.py"

# First check if the server is up
try {
    $r = Invoke-WebRequest "$baseUrl/health" -UseBasicParsing -TimeoutSec 5
    Write-Host "Server: $($r.Content)"
} catch {
    Write-Host "Server not running!"
    exit 1
}

# Check upload dir exists
$testDir = "C:\Users\pc\ZCodeProject\test"
$uploads = Join-Path $testDir "uploads"
$outputs = Join-Path $testDir "outputs"
Write-Host "Upload dir exists: $(Test-Path $uploads)"
Write-Host "Output dir exists: $(Test-Path $outputs)"

# Try to call the go() function endpoints via direct POST
# Create a simple test file
$wb = New-Object -ComObject Excel.Application
$wb.Visible = $false
$workbook = $wb.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$ws.Cells.Item(1,1) = "货件单号"
$ws.Cells.Item(1,2) = "物流中心编码"
$ws.Cells.Item(1,3) = "物流商"
$ws.Cells.Item(1,4) = "物流渠道"
$ws.Cells.Item(1,5) = "物流商单号"
$ws.Cells.Item(1,7) = "国家"
$ws.Cells.Item(1,8) = "货件单号"
$ws.Cells.Item(2,1) = "TEST001"
$ws.Cells.Item(2,2) = "东北部中心1"
$ws.Cells.Item(2,3) = "FedEx"
$ws.Cells.Item(2,4) = "Ground"
$ws.Cells.Item(2,5) = "FDX123"
$ws.Cells.Item(2,7) = "US"
$ws.Cells.Item(2,8) = "SHIP001"
$ws5 = $workbook.Worksheets.Add()
$ws5.Name = "Sheet5"
$ws5.Cells.Item(1,1) = "货件单号"
$ws5.Cells.Item(1,2) = "箱数"
$ws5.Cells.Item(1,3) = "重量"
$ws5.Cells.Item(1,4) = "体积"
$ws5.Cells.Item(1,5) = "类型"
$ws5.Cells.Item(2,1) = "TEST001"
$ws5.Cells.Item(2,2) = 10
$ws5.Cells.Item(2,3) = 100.5
$ws5.Cells.Item(2,4) = 2.5
$ws5.Cells.Item(2,5) = "A"
$workbook.SaveAs($testFile)
$workbook.Close()
$wb.Quit()
Write-Host "Test file created: $testFile"

# Now try to POST to /run with action=us
$uri = "$baseUrl/run"
$boundary = "----TestBoundary" + [DateTime]::Now.Ticks
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes("--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"test.xlsx`"`r`nContent-Type: application/octet-stream`r`n`r`n")
$bodyBytes += [System.IO.File]::ReadAllBytes($testFile)
$bodyBytes += [System.Text.Encoding]::UTF8.GetBytes("`r`n--$boundary`r`nContent-Disposition: form-data; name=`"action`"`r`n`r`nus`r`n--$boundary--`r`n")

try {
    $r = Invoke-WebRequest -Uri $uri -Method POST -Body $bodyBytes -ContentType "multipart/form-data; boundary=$boundary" -UseBasicParsing -TimeoutSec 15
    Write-Host "Response: $($r.Content)"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        Write-Host "Details: $($reader.ReadToEnd())"
    }
}
