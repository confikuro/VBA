Option Explicit
 
' 構文　引数は0以外であれば、引数として設定
'Call createshorcut(WorkingDirectory, proguram, ショートカット名, アイコン画像パス, 引数)
 
Call createshorcut("C:\Program Files (x86)\Google\Chrome\Application", "chrome.exe", "C:\Users\Confi\Documents\cre\　 test.lnk", "C:\Users\Confi\Documents\obake.png")
 
Function createshorcut(path, program, name, icon)
    Dim WSH,sc
    Set WSH=CreateObject("WScript.Shell")
    Set sc = WSH.CreateShortcut(name)
    WScript.Echo sc
    sc.TargetPath = path & "\" & program
    WScript.Echo sc.TargetPath
    sc.IconLocation = icon
    'If arg <> "0" Then
    '    sc.Arguments = "引数"
    'End If
    sc.WorkingDirectory = path
    sc.save
    Set sc = Nothing
    Set WSH = nothing
End Function