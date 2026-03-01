# Logon/Logoff Extractor – 90 Days (4624+4634)
# Self-elevate
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`"" -Wait; exit
}
# Setup
if (-not (Get-Module -ListAvailable ImportExcel)) { Install-Module ImportExcel -Force -EA Stop }; Import-Module ImportExcel -EA Stop
$Now=Get-Date; $s=$Now.AddDays(-90); $si=$s.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z"); $ei=$Now.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")
$o=Join-Path([Environment]::GetFolderPath('Desktop')) "EvtExport_$($Now.ToString('yyyyMMdd_HHmmss'))"; mkdir $o -Force|Out-Null; $xl=Join-Path $o "LogonLogoff.xlsx"
Write-Host "`n=== Logon/Logoff Extractor | $($s.ToString('yyyy-MM-dd')) to $($Now.ToString('yyyy-MM-dd')) ===" -ForegroundColor Green
# Fetch + parse once
try { $raw=@(Get-WinEvent -LogName Security -FilterXPath "*[System[(EventID=4624 or EventID=4634) and TimeCreated[@SystemTime>='$si' and @SystemTime<='$ei']]]" -EA Stop) } catch { $raw=@() }
if (!$raw.Count) { Write-Host "No events found." -ForegroundColor DarkGray; Read-Host "`nEnter to close"; exit }
$All=foreach($e in $raw){[xml]$x=$e.ToXml();$ns=New-Object System.Xml.XmlNamespaceManager($x.NameTable);$ns.AddNamespace("e","http://schemas.microsoft.com/win/2004/08/events/event")
    $r=[ordered]@{TimeCreated=$e.TimeCreated;EventID=$e.Id;Computer=$e.MachineName};foreach($n in $x.SelectNodes("//e:EventData/e:Data",$ns)){if($n.Name){$r[$n.Name]=$n.'#text'}};[PSCustomObject]$r}
# Excel tabs
foreach($eid in 4624,4634){$lbl=@{4624="Logon (4624)";4634="Logoff (4634)"}[$eid];$sub=@($All|Where-Object EventID -eq $eid)
    if($sub.Count){$sub|Export-Excel -Path $xl -WorksheetName $lbl -AutoSize -AutoFilter -FreezeTopRow -Append;Write-Host "$lbl : $($sub.Count)" -ForegroundColor Green}}
# Sessions (interactive types 2,10,11)
$logons=$All|Where-Object{$_.EventID -eq 4624 -and $_.LogonType -in '2','10','11' -and $_.TargetUserName -and $_.TargetUserName -ne '-' -and $_.TargetUserName -notmatch '\$$'}
$loMap=@{};$All|Where-Object{$_.EventID -eq 4634 -and $_.TargetLogonId}|ForEach-Object{$loMap[$_.TargetLogonId]=$_.TimeCreated}
$sess=@(foreach($li in $logons){$lo=$loMap[$li.TargetLogonId];if($lo -and $lo -gt $li.TimeCreated){$d=$lo-$li.TimeCreated
    [PSCustomObject]@{Date=$li.TimeCreated.ToString('yyyy-MM-dd');User="$($li.TargetDomainName)\$($li.TargetUserName)";LogonTime=$li.TimeCreated;LogoffTime=$lo
        Duration=("{0:d2}:{1:d2}:{2:d2}" -f [int]$d.TotalHours,$d.Minutes,$d.Seconds);Mins=[math]::Round($d.TotalMinutes,2);LogonType=$li.LogonType}}})
if ($sess.Count) {
    $sess|Sort-Object LogonTime -Desc|Export-Excel -Path $xl -WorksheetName "Session Details" -AutoSize -AutoFilter -FreezeTopRow -Append
    $sess|Group-Object Date,User|ForEach-Object{$p=$_.Name -split ', ';$m=($_.Group|Measure-Object Mins -Sum).Sum;$t=[timespan]::FromMinutes($m)
        [PSCustomObject]@{Date=$p[0];User=$p[1];Sessions=$_.Count;FirstLogon=($_.Group|Sort-Object LogonTime|Select-Object -First 1).LogonTime
            LastLogoff=($_.Group|Sort-Object LogoffTime -Desc|Select-Object -First 1).LogoffTime
            TotalDuration=("{0:d2}:{1:d2}:{2:d2}" -f [int]$t.TotalHours,$t.Minutes,$t.Seconds);TotalMinutes=[math]::Round($m,2)}
    }|Sort-Object Date,User|Export-Excel -Path $xl -WorksheetName "Daily Summary" -AutoSize -AutoFilter -FreezeTopRow -Append
    Write-Host "Sessions: $($sess.Count) matched" -ForegroundColor Green}
# CSV + Done
$All|Sort-Object TimeCreated -Desc|Export-Csv (Join-Path $o "Combined.csv") -NoTypeInformation -Encoding UTF8
Write-Host "=== Done === Output: $o" -ForegroundColor Green; Read-Host "`nEnter to close"