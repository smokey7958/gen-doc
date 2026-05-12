$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('C:\Gen-Doc\Gen Doc.lnk')
$sc.TargetPath = 'C:\Gen-Doc\release\win-unpacked\Gen Doc.exe'
$sc.WorkingDirectory = 'C:\Gen-Doc\release\win-unpacked'
$sc.IconLocation = 'C:\Gen-Doc\release\win-unpacked\Gen Doc.exe,0'
$sc.Description = 'Gen Doc — 統一筆記 / 文書 / 表格 / 簡報，內建 AI 編輯助手'
$sc.Save()
Write-Host "Shortcut created at C:\Gen-Doc\Gen Doc.lnk"
