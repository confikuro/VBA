Option Explicit
 
' �\���@������0�ȊO�ł���΁A�����Ƃ��Đݒ�
'Call createshorcut(WorkingDirectory, proguram, �V���[�g�J�b�g��, �A�C�R���摜�p�X, ����)
 
Call createshorcut("C:\Program Files (x86)\Google\Chrome\Application", "chrome.exe", "C:\Users\Confi\Documents\cre\�@ test.lnk", "C:\Users\Confi\Documents\obake.png")
 
Function createshorcut(path, program, name, icon)
    Dim WSH,sc
    Set WSH=CreateObject("WScript.Shell")
    Set sc = WSH.CreateShortcut(name)
    WScript.Echo sc
    sc.TargetPath = path & "\" & program
    WScript.Echo sc.TargetPath
    sc.IconLocation = icon
    'If arg <> "0" Then
    '    sc.Arguments = "����"
    'End If
    sc.WorkingDirectory = path
    sc.save
    Set sc = Nothing
    Set WSH = nothing
End Function