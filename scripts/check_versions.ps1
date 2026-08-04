$r1 = Invoke-WebRequest "http://127.0.0.1:8932/" -UseBasicParsing -TimeoutSec 5
$c1 = $r1.Content
if ($c1 -match "<title>(.*?)</title>") { Write-Host ("稳定版 8932: " + $matches[1]) }
if ($c1 -match "v[\d.]+") { Write-Host ("  版本: " + $matches[0]) }
$hasTest = $c1 -match "测试版"
Write-Host ("  含测试版标记: " + $hasTest)
if ($c1 -match '<div class="tb"><div class="logo">([^<]+)') { Write-Host ("  标题栏: " + $matches[1]) }

Write-Host ""
$r2 = Invoke-WebRequest "http://127.0.0.1:8933/" -UseBasicParsing -TimeoutSec 5
$c2 = $r2.Content
if ($c2 -match "<title>(.*?)</title>") { Write-Host ("测试版 8933: " + $matches[1]) }
if ($c2 -match "v[\d.]+") { Write-Host ("  版本: " + $matches[0]) }
$hasTest2 = $c2 -match "测试版"
Write-Host ("  含测试版标记: " + $hasTest2)
if ($c2 -match '<div class="tb"><div class="logo">([^<]+)') { Write-Host ("  标题栏: " + $matches[1]) }
