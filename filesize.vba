'======================================================================
'�t�@�C���ꗗ�A�e�ʎ擾
'======================================================================
Option Explicit

Sub setFileList(searchPath)
    Dim startCell As Range
    Dim maxRow As Long
    Dim maxCol As Long

    '�o�͊J�n����Z��
    Set startCell = Cells(6, 2)
    startCell.Select
    
    '�V�[�g������U�N���A����
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

    '�T�u�t�H���_�擾
    For Each objFolders In FSO.GetFolder(searchPath).SubFolders
        Call getFileList(objFolders.Path)
    Next
    
    '�t�@�C�����̎擾
    For Each objFiles In FSO.GetFolder(searchPath).Files
        separateNum = InStrRev(objFiles.Path, "\")
        
        '�Z���Ƀp�X�ƃt�@�C�����A�e�ʂ���������
        ActiveCell.Value = Left(objFiles.Path, separateNum - 1)
        ActiveCell.Offset(0, 1).Value = Right(objFiles.Path, Len(objFiles.Path) - separateNum)
        ActiveCell.Offset(0, 2).Value = Format(FileLen(objFiles), "#.0")
        'ActiveCell.Offset(0, 2).Value = Format((FileLen(objFiles) / 1024), "#.0") 'MB�ϊ���
        'ActiveCell.Offset(0, 3).Value = FileDateTime(objFiles) '�^�C���X�^���v
        ActiveCell.Offset(1, 0).Select
    Next
     
End Sub
