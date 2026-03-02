# Office 2021 Pro Volume - Install & Activate
$ErrorActionPreference = "Stop"
$d = "C:\ODT"; $ospp = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"

New-Item $d -ItemType Directory -Force | Out-Null
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue; Start-Sleep 2
$tmp = "$env:TEMP\odt_setup.exe"
Invoke-WebRequest "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $tmp
Copy-Item $tmp "$d\setup.exe" -Force
'<Configuration><Add OfficeClientEdition="64" Channel="PerpetualVL2021"><Product ID="Pro2021Volume"><Language ID="en-us"/></Product></Add><Display Level="None" AcceptEULA="TRUE"/></Configuration>' | Out-File "$d\configuration.xml" -Encoding UTF8

Set-Location $d
"Downloading","Installing" | ForEach-Object { Write-Host "$_ Office 2021..."; Start-Process .\setup.exe -ArgumentList "$(if($_ -eq 'Downloading'){'/download'}else{'/configure'}) configuration.xml" -Wait }

if (Test-Path $ospp) { $k = Read-Host "Enter product key"; cscript $ospp /inpkey:$k; cscript $ospp /act }
else { Write-Host "Office not found. Activation skipped." }
Write-Host "Done."
