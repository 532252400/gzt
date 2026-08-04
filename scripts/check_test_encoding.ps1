$url = "http://127.0.0.1:8933/workshop"
try {
    $r = Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 5
    $c = $r.Content
    if ($c -match '<title>(.*?)</title>') {
        Write-Host ("Title: " + $matches[1])
    }
    # Check title content for garbled chars
    $titleLine = ($c -split "`n" | Select-String "<title>")[0]
    Write-Host ("Title line: " + $titleLine.Substring(0, [Math]::Min(80, $titleLine.Length)))
    
    # Check first few Chinese characters  
    $hasChinese = $c -match '[\u4e00-\u9fff]'
    Write-Host ("Has proper Chinese: " + $hasChinese)
    
    # Check for common mojibake patterns
    $hasGarbled = $c -match '宸ュ巶|鎵撳嵃|鏍囩|鏂扮増|鍔犲伐'
    Write-Host ("Has garbled Chinese pattern: " + $hasGarbled)
} catch {
    Write-Host ("Error: " + $_.Exception.Message)
}

# Also check the main page
Write-Host ""
$url2 = "http://127.0.0.1:8933/"
try {
    $r2 = Invoke-WebRequest $url2 -UseBasicParsing -TimeoutSec 5
    $c2 = $r2.Content
    if ($c2 -match '<title>(.*?)</title>') {
        Write-Host ("Main page title: " + $matches[1])
    }
    $hasGarbled2 = $c2 -match '宸ュ巶|鎵撳嵃|鏍囩|鍔犲伐'
    Write-Host ("Main page garbled: " + $hasGarbled2)
} catch {
    Write-Host ("Error: " + $_.Exception.Message)
}
