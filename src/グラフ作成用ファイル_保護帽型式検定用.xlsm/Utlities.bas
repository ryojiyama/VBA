Attribute VB_Name = "Utlities"
' DeleteAllChartsAndSheets_�V�[�g���̃O���t�Ɨ]�v�ȃV�[�g���폜����
Sub DeleteAllChartsAndSheets()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String
    Dim proceed As Integer

    ' �V�[�g�̃��X�g
    Dim sheetList() As Variant
    sheetList = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")

    Application.DisplayAlerts = False

    ' �e�V�[�g�ɑ΂��ď��������s
    For Each sheet In ThisWorkbook.Sheets
        sheetName = sheet.name
        ' �O���t�̍폜�ƃf�[�^�̌x���\��
        If IsInArray(sheetName, sheetList) Then
            For Each chart In sheet.ChartObjects
                chart.Delete
            Next chart
            ' B2�Z������ZZ15�܂ł̃f�[�^�̗L�����`�F�b�N���A�L��Όx����\��
            If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
                Application.DisplayAlerts = True
                proceed = MsgBox("Sheet '" & sheetName & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
                Application.DisplayAlerts = False
                If proceed = vbNo Then Exit Sub
            End If
        ' �V�[�g�̍폜
        ElseIf sheetName <> "Setting" And sheetName <> "Hel_SpecSheet" Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True


End Sub

' DeleteAllChartsAndSheets_�z����ɓ���̒l�����݂��邩�`�F�b�N����֐�
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub PrintImpactSheet()
    Dim ws As Worksheet
    
    ' ����1: ����̃V�[�g�����
    Dim sheetNames1 As Variant
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    For Each ws In ThisWorkbook.Sheets
        If foundSheetName(ws.name, sheetNames1) Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    
    ' ����2: "Impact_Side"�𖼑O�Ɋ܂ރV�[�g�����
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.name, "Impact_Side") > 0 Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Function foundSheetName(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            foundSheetName = True
            Exit Function
        End If
    Next i
    foundSheetName = False
End Function
' Impact���܂ރV�[�g���̒���
Sub DeleteRowsBelowHeader()
    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim sheetName As String

    ' ���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g����"Impact"���܂܂�Ă��邩�`�F�b�N
        If InStr(ws.name, "Impact") > 0 Then
            ' �w�b�_�[�̉��̍s����ŏI�s�܂ł��폜
            ws.Rows("15:" & ws.Rows.Count).Delete
        End If
    Next ws
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
