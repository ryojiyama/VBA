Attribute VB_Name = "Utliteis"
'CopiedSheetNames�ŋL����Ă���V�[�g���폜����B
Sub DeleteCopiedSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "CopiedSheetNames�V�[�g��������܂���B"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    Dim i As Long
    Application.DisplayAlerts = False
    For i = 1 To lastRow
        Dim sheetName As String
        sheetName = ws.Cells(i, 1).value
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    ClearCopiedSheetNames
End Sub
'CopiedSheetNames���N���A����B
Sub ClearCopiedSheetNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Cells.ClearContents
    End If
End Sub
' "LOG_Helmet��̃O���t���폜����
Public Sub DeleteAllChartsOnLOG_Helmet()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    ' "LOG_Helmet"�V�[�g���擾
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' �V�[�g��̂��ׂẴO���t�I�u�W�F�N�g�����[�v
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Sub PrintMatchingSheetsFirstPage()
    Dim ws As Worksheet
    Dim copiedSheetNames As Worksheet
    Dim sheetName As String
    Dim cell As Range
    Dim foundSheet As Worksheet
    
    ' CopiedSheetNames�V�[�g��ݒ�
    Set copiedSheetNames = ThisWorkbook.Sheets("CopiedSheetNames")
    
    ' A��̒l�����[�v
    For Each cell In copiedSheetNames.Range("A1:A" & copiedSheetNames.Cells(copiedSheetNames.Rows.count, "A").End(xlUp).row)
        sheetName = cell.value
        
        ' ��v����V�[�g������
        On Error Resume Next
        Set foundSheet = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' �V�[�g�����݂���ꍇ�A1�y�[�W�ڂ����
        If Not foundSheet Is Nothing Then
            With foundSheet
                ' ����̈��ݒ�
                .PageSetup.PrintArea = ""
                ' �V�[�g��1�y�[�W�ڂ݈̂��
                .PrintOut Preview:=False
            End With
            ' foundSheet���N���A
            Set foundSheet = Nothing
        End If
    Next cell
End Sub




'Sub DeleteAllChartsAndSheets()
'    ' �V�[�g���̃O���t�Ɨ]�v�ȃV�[�g���폜
'    Dim sheet As Worksheet
'    Dim chart As ChartObject
'    Dim sheetName As String
'    Dim proceed As Integer
'
'    ' �V�[�g�̃��X�g
'    Dim sheetList() As Variant
'    sheetList = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
'
'    Application.DisplayAlerts = False
'
'    ' �e�V�[�g�ɑ΂��ď��������s
'    For Each sheet In ThisWorkbook.Sheets
'        sheetName = sheet.name
'        ' �O���t�̍폜�ƃf�[�^�̌x���\��
'        If IsInArray(sheetName, sheetList) Then
'            For Each chart In sheet.ChartObjects
'                chart.Delete
'            Next chart
'            ' B2�Z������ZZ15�܂ł̃f�[�^�̗L�����`�F�b�N���A�L��Όx����\��
'            If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
'                Application.DisplayAlerts = True
'                proceed = MsgBox("Sheet '" & sheetName & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
'                Application.DisplayAlerts = False
'                If proceed = vbNo Then Exit Sub
'            End If
'        ' �V�[�g�̍폜
'        ElseIf sheetName <> "Setting" And sheetName <> "Hel_SpecSheet" And sheetName <> "InspectionSheet" Then
'            sheet.Delete
'        End If
'    Next sheet
'
'    Application.DisplayAlerts = True
'
'    ' �u�b�N��ۑ�
'    ThisWorkbook.Save
'
'End Sub
