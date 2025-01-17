Attribute VB_Name = "Utlities"

Public Sub DeleteReportGraphSheets()
    Dim ws As Worksheet
    Dim i As Long
    
    ' ��납��O�Ƀ��[�v���ăV�[�g���폜
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        
        ' �V�[�g���Ɂu���|�[�g�O���t�v���܂܂�Ă���V�[�g���폜
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then
            Application.DisplayAlerts = False  ' �폜�m�F���b�Z�[�W��\�����Ȃ�
            ws.Delete
            Application.DisplayAlerts = True   ' �폜�m�F���b�Z�[�W�̕\�������ɖ߂�
        End If
    Next i
End Sub

' "���|�[�g�{��"�V�[�g��L��� "Insert" �ƈ󂪂��Ă���s���폜����
Public Sub DeleteInsertedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' "���|�[�g�{��"�V�[�g���擾
    Set ws = ThisWorkbook.Sheets("���|�[�g�{��")
    
    ' I��̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).row
    
    ' �Ō�̍s����1�s����Ɍ������č폜���m�F
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "L").value, 6) = "Insert" Then
            ws.Rows(currentRow).Delete
        End If
    Next currentRow
End Sub

' �V�[�g����"Impact"�Ƃ��Ă���V�[�g���폜����B
Sub DeleteImpactSheets()
    Dim ws As Worksheet
    Dim sheetNamesToDelete As Collection
    Dim sheetName As String
    Dim i As Long
    
    ' �폜�Ώۂ̃V�[�g�����ꎞ�I�ɕێ�����R���N�V�������쐬
    Set sheetNamesToDelete = New Collection
    
    ' ���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g����"Impact"���܂܂�Ă��邩�`�F�b�N
        If InStr(ws.Name, "Impact") > 0 Then
            ' �폜�Ώۂ̃V�[�g�����R���N�V�����ɒǉ�
            sheetNamesToDelete.Add ws.Name
        End If
    Next ws
    
    ' �R���N�V�������̃V�[�g���폜
    For i = sheetNamesToDelete.Count To 1 Step -1
        ThisWorkbook.Sheets(sheetNamesToDelete(i)).Delete
    Next i
End Sub
