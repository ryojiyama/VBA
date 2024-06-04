Attribute VB_Name = "Main"

Public Sub ShowForm()
    Call Form_Helmet.Show
End Sub



Sub MultiplyValues()
    ' 1�s�ڂ�ms�̒l�ɕϊ�����B(1000�������邾��)
    ' ���[�N�V�[�g��ݒ�
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' �J�n��ƏI�����ݒ�
    Dim startColumn As Integer: startColumn = 22 ' V��
    Dim endColumn As Integer: endColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column ' �ŏI��

    ' �񂲂ƂɃ��[�v
    Dim i As Integer
    For i = startColumn To endColumn
        ' �Z���̒l��1000���|����
        ws.Cells(1, i).Value = ws.Cells(1, i).Value * 1000
    Next i
End Sub



' �e��ɏ����ݒ������
Sub FormatCells()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range
    
    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)
        
        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.Value, "�ő�l(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.00 ""kN"""
            ElseIf InStr(1, cell.Value, "�ő�l(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0 ""G"""
            ElseIf InStr(1, cell.Value, "����") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""ms"""
            ElseIf InStr(1, cell.Value, "���x") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""��"""
            ElseIf InStr(1, cell.Value, "�d��") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""g"""
            ElseIf InStr(1, cell.Value, "���b�g") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "@"
            End If
        Next cell
    Next sheet
End Sub



Sub ClickIconAttheTop()
    ' �E�̃V�[�g�Ɉړ�����
    On Error Resume Next
    ActiveSheet.Next.Select
    If Err.number <> 0 Then
        MsgBox "This is the last sheet."
    End If
    On Error GoTo 0
End Sub

Sub ClickUSBIcon()
    'USB�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClickGraphIcon()
    '�O���t�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClickPhotoIcon()
    '�摜�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClicIconAttheBottom()
    ' ���̃V�[�g�Ɉړ�����
    On Error Resume Next
    ActiveSheet.Previous.Select
    If Err.number <> 0 Then
        MsgBox "This is the first sheet."
    End If
    On Error GoTo 0
End Sub





