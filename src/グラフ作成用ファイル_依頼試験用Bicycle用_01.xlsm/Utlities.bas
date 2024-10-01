Attribute VB_Name = "Utlities"
Sub ShowFormInspectionType()
    ' ���[�U�[�t�H�[�� "Form_InspectionType" ��\��
    Form_InspectionType.Show
End Sub
Sub ShowFormTenki()
    ' ���[�U�[�t�H�[�� "Form_Tenki" ��\��
    Form_Tenki.Show
End Sub

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
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                With chartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                End With
            Next chartObj
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

Sub AdjustingDuplicateValues()
    ' �ΏۃV�[�g���̃��X�g
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Double
    Dim sheetName As Variant
    Dim newValue As Double
    Dim randomDigit As Integer
    Dim roundedValue As Double
    Dim maxCol As Long

    ' �V�[�g���Ƃɏ���
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row
        
        ' "�ő�l"���܂ރw�b�_�[������������
        maxCol = 0
        For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            If InStr(ws.Cells(1, i).value, "�ő�l") > 0 Then
                maxCol = i
                Exit For
            End If
        Next i
        
        ' "�ő�l"�񂪌�����Ȃ���Ύ��̃V�[�g��
        If maxCol = 0 Then
            MsgBox "�V�[�g " & sheetName & " �ɂ� '�ő�l' ���܂ޗ񂪌�����܂���B"
            GoTo NextSheet
        End If

        For i = 2 To lastRow
            ' �Z���̒l�����l���ǂ����m�F
            If IsNumeric(ws.Cells(i, maxCol).value) Then
                ' ���l�Ƃ��Ď擾���A�����_�ȉ�2���Ŋۂ߂�
                roundedValue = Round(ws.Cells(i, maxCol).value, 2)

                If ws.Cells(i, maxCol).Interior.colorIndex = xlNone Then
                    For j = i + 1 To lastRow
                        ' �d���l���`�F�b�N�i���l�`�F�b�N��ǉ��j
                        If IsNumeric(ws.Cells(j, maxCol).value) And Round(ws.Cells(j, maxCol).value, 2) = roundedValue And ws.Cells(j, maxCol).Interior.colorIndex = xlNone Then
                            Debug.Print "Duplicate Row Number: " & j
                            Do
                                ' 1����9�̃����_���Ȑ��𐶐�
                                randomDigit = Int((9 - 1 + 1) * Rnd + 1)
                                ' ���̒l�Ƀ����_���Ȓl�������_�ȉ�4���Ƃ��Ēǉ��i�����_�ȉ�2���͈ێ��j
                                newValue = roundedValue + randomDigit / 10000
                                Debug.Print "New Value: " & newValue
                            Loop While WorksheetFunction.CountIf(ws.Range(ws.Cells(2, maxCol), ws.Cells(lastRow, maxCol)), newValue) > 0
                            
                            ' �V�����l���Z���ɐݒ�
                            ws.Cells(j, maxCol).value = newValue
                        End If
                    Next j
                End If
            End If
        Next i
NextSheet:
    Next sheetName
End Sub


Sub ListChartNames()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    ' Loop through all sheets in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            For Each chartObj In ws.ChartObjects
                ' Display the chart name
                MsgBox "�V�[�g��: " & ws.Name & vbCrLf & "�O���t��: " & chartObj.Name, vbInformation
            Next chartObj
        End If
    Next ws
End Sub


' �e��ɏ����ݒ������
Public Sub CustomizeSheetFormats()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "ID") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "����ID") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�i��") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�������e") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "������") > 0 Then ' Date
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToDate(rng)
            ElseIf InStr(1, cell.value, "���x") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "�ő�l(kN)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericFourDecimals(rng)
            ElseIf InStr(1, cell.value, "�ő�l�̎���(ms)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "4.9kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "7.3kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "�O����") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�d��") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "�V��������") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "���i���b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�X�̃��b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�������b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�\������") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�ϊђʌ���") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�����敪") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            End If
        Next cell
    Next sheet
End Sub

Sub ConvertToNumeric(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericTwoDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.00"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericFourDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0000"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToString(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "@"
    For Each cell In rng
        cell.value = CStr(cell.value)
    Next cell
End Sub

Sub ConvertToDate(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "yyyy/mm/dd"  ' ���t�̕\���`����ݒ�
    For Each cell In rng
        If IsDate(cell.value) Then
            cell.value = CDate(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub
' �󔒃Z����"-"��}��
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
        lastCol = ws.Cells(1, "Z").column ' Z��̗�ԍ���ݒ�

        ' 2�s�ڂ���ŏI�s�܂Ń��[�v�i1�s�ڂ̓w�b�_�[�Ɖ���j
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    'Debug.Print "EmptyCell:" & "Cells&("; i; "," & j; ")"
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

Sub DeleteAllChartsAndDataFromActiveSheet()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim proceed As Integer

    ' �A�N�e�B�u�ȃV�[�g���擾
    Set sheet = ThisWorkbook.ActiveSheet

    Application.DisplayAlerts = False

    ' �O���t�̍폜
    For Each chart In sheet.ChartObjects
        chart.Delete
    Next chart

    ' B2�Z������ZZ15�܂ł̃f�[�^�̗L�����`�F�b�N���A�L��Όx����\��
    If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
        Application.DisplayAlerts = True
        proceed = MsgBox("Sheet '" & sheet.Name & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
        Application.DisplayAlerts = False
        If proceed = vbNo Then Exit Sub
    End If

    Application.DisplayAlerts = True
End Sub

