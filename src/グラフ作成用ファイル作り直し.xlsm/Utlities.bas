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
        sheetName = sheet.Name
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

Sub UniformizeLineGraphAxes()

    On Error GoTo ErrorHandler

    ' Display input dialog to set the maximum value for the axes
    Dim MaxValue As Variant
    MaxValue = InputBox("Y���̍ő�l����͂��Ă��������B(����)", "�ő�l�����")

    ' Check if the user pressed Cancel
    If MaxValue = False Then
        MsgBox "���삪�L�����Z������܂����B", vbInformation
        Exit Sub
    End If

    ' Validate the input
    If Not IsNumeric(MaxValue) Or MaxValue <= 0 Then
        MsgBox "�L���Ȑ��l����͂��Ă��������B", vbExclamation
        Exit Sub
    End If

    MaxValue = CDbl(MaxValue)

    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim ChartObj As ChartObject
            For Each ChartObj In ws.ChartObjects
                With ChartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                End With
            Next ChartObj
        End If
    Next ws

    MsgBox "���ׂẴV�[�g�̃O���t��Y���̍ő�l�� " & MaxValue & " �ɐݒ肵�܂����B", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical

End Sub


Sub HighlightDuplicateValues()
    ' �ΏۃV�[�g���̃��X�g
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    Dim sheetName As Variant

    ' �V�[�g���Ƃɏ���
    For Each sheetName In sheetNames
        ' �V�[�g�I�u�W�F�N�g��ݒ�
        Set ws = ThisWorkbook.Sheets(sheetName)

        ' �ŏI�s���擾
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

        ' �F�̃C���f�b�N�X��������
        colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�

        For i = 2 To lastRow
            ' ���݂̃Z���̒l���擾
            valueToFind = ws.Cells(i, "H").value

            ' �����l�����Z�������ɐF�t������Ă��Ȃ����`�F�b�N
            If ws.Cells(i, "H").Interior.colorIndex = xlNone Then
                For j = i + 1 To lastRow
                    If ws.Cells(j, "H").value = valueToFind And ws.Cells(j, "H").Interior.colorIndex = xlNone Then
                        ' �����l�����Z�����������ꍇ�A�F��h��
                        ws.Cells(i, "H").Interior.colorIndex = colorIndex
                        ws.Cells(j, "H").Interior.colorIndex = colorIndex
                    End If
                Next j

                ' �F�C���f�b�N�X���X�V���Ď��̐F�ɕύX
                colorIndex = colorIndex + 1
                ' Excel�̐F�C���f�b�N�X�̍ő�l�𒴂��Ȃ��悤�Ƀ`�F�b�N
                If colorIndex > 56 Then colorIndex = 3 ' �F�C���f�b�N�X�����Z�b�g
            End If
        Next i
    Next sheetName
End Sub

Public Sub FillBlanksWithHyphenInMultipleSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim sheetName As Variant

    ' �ΏۃV�[�g�̖��O��z��ɐݒ�
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' �e�V�[�g�ɂ��ď������s��
    For Each sheetName In sheetNames
        On Error Resume Next
        ' �ΏۃV�[�g��ݒ�
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = Nothing ' ws�ϐ����N���A
            GoTo NextSheet ' ���̃V�[�g�ɐi��
        End If

        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        lastCol = ws.Cells(1, "Z").Column ' Z��̗�ԍ���ݒ�

        ' 2�s�ڂ���ŏI�s�܂Ń��[�v�i1�s�ڂ̓w�b�_�[�Ɖ���j
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").Column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    ws.Cells(i, j).value = "-"
                End If
            Next j
        Next i

        ' �V�[�g�����̏I�����x��
NextSheet:
        ' ���̃V�[�g�̏����Ɉڂ�O�ɕϐ����N���A
        Set ws = Nothing
    Next sheetName
End Sub
