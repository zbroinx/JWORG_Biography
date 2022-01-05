$website = 'https://wol.jw.org/pt/wol/d/r5/lp-t/1200273453'


Clear-Host

Write-Host "Check if file exist..." -ForegroundColor Yellow
# If file exists, delete it
$FileName = "c:\temp\WatchtowerUrls.txt"
if (Test-Path $FileName) {
  Remove-Item $FileName
  Write-Host "File deleted." -ForegroundColor Red
}


Write-Host "Getting watchtower URLs..." -ForegroundColor Yellow
#Get Watchtower URLs
$watchtower = (Invoke-WebRequest -Uri $website).Links | Where-Object {$_.innerText -like "*w*/*"}| Select-Object href

foreach ($wt in $watchtower) {

    $wURL = "https://wol.jw.org/" + $wt.href

    Add-Content c:\temp\WatchtowerUrls.txt $wURL
}

Write-Host "Getting awake URLs..." -ForegroundColor Yellow
#Get Watchtower URLs

$awake = (Invoke-WebRequest -Uri $website).Links | Where-Object {$_.innerText -like "*g*/*"}| Select-Object href

foreach ($ak in $awake) {

    $aURL = "https://wol.jw.org/" + $ak.href

    Add-Content c:\temp\WatchtowerUrls.txt $aURL
}

Write-Host "Done!!!" -ForegroundColor Green