'======================================================================
'ファイル一覧、容量取得
'======================================================================
Option Explicit

Sub setFileList(searchPath)
    Dim startCell As Range
    Dim maxRow As Long
    Dim maxCol As Long

    '出力開始するセル
    Set startCell = Cells(6, 2)
    startCell.Select
    
    'シート情報を一旦クリアする
    maxRow = startCell.SpecialCells(xlLastCell).Row
    maxCol = startCell.SpecialCells(xlLastCell).Column
    Range(startCell, Cells(maxRow, maxCol)).ClearContents
    
    Call getFileList(searchPath)
    startCell.Select
End Sub

Sub getFileList(searchPath)

    Dim FSO As New FileSystemObject
    Dim objFiles As File
    Dim objFolders As Folder
    Dim separateNum As Long

    'サブフォルダ取得
    For Each objFolders In FSO.GetFolder(searchPath).SubFolders
        Call getFileList(objFolders.Path)
    Next
    
    'ファイル名の取得
    For Each objFiles In FSO.GetFolder(searchPath).Files
        separateNum = InStrRev(objFiles.Path, "\")
        
        'セルにパスとファイル名、容量を書き込む
        ActiveCell.Value = Left(objFiles.Path, separateNum - 1)
        ActiveCell.Offset(0, 1).Value = Right(objFiles.Path, Len(objFiles.Path) - separateNum)
        ActiveCell.Offset(0, 2).Value = Format(FileLen(objFiles), "#.0")
        'ActiveCell.Offset(0, 2).Value = Format((FileLen(objFiles) / 1024), "#.0") 'MB変換時
        'ActiveCell.Offset(0, 3).Value = FileDateTime(objFiles) 'タイムスタンプ
        ActiveCell.Offset(1, 0).Select
    Next
     
End Sub
