Attribute VB_Name = "SpecSheet"
Sub SetupTestSamples()
    Call CreateInspectionSheetIDs
    Call InsertXLookupAndUpdateKColumn
End Sub


Sub SyncSpecSheetToLogHel()
    ' �A�C�R���ɕR�Â��BSpecSheet�ɓ]�L����v���V�[�W���̂܂Ƃ�
    ' ���l�����������ꍇ�̓G���[���b�Z�[�W��\�����ď����𒆒f
    If HighlightDuplicateValues Then
        MsgBox "�Ռ��l�œ��l��������܂����B�����_���񌅂ɉe�����o�Ȃ��͈͂ŏC�����Ă��������B", vbCritical
        Exit Sub
    End If
    
    Dim errMsg As String
    errMsg = LocateEmptySpaces()
    
    If errMsg <> "" Then
        ' �G���[���b�Z�[�W������ꍇ�A�����\��
        MsgBox "�ȉ��̖�肪����܂��B�܂��͂������������Ă��������F" & vbNewLine & errMsg, vbCritical
        Exit Sub
    Else
    End If
    
    Call CopyDataBasedOnCondition
    Call CustomizeSheetFormats
    MsgBox "�]�L���I�����܂����B"
End Sub


Sub CreateInspectionSheetIDs_0410Before()
    ' SpecSheet��B��Ɏ���ID���쐬����B����͓]�L����Ƃ��̃L�[�Ƃ��Ďg�p����B
    
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' D��̍ŏI�s���擾
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row

    Dim i As Long
    For i = 2 To lastRow
        ' D��ɒl������s�̏ꍇ�̂ݏ���
        If ws.Cells(i, "D").value <> "" Then
            ' S��Ɏ���ݒ�
            ws.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""����"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""�ቷ"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""�Z����"", ""Wet"", """")))"

            ' ID���쐬
            Dim id As String
            id = ws.Cells(i, "D").value & "-" & ws.Cells(i, "S").value & "-" & Left(ws.Cells(i, "E").value, 1)

            ' D��̒l��"F"���܂܂�Ă���ꍇ�AID�̐擪��"F"��ǉ�
            If InStr(ws.Cells(i, "D").value, "F") > 0 Then
                id = "F" & id
            End If

            ' �쐬����ID��B��ɐݒ�
            ws.Cells(i, "B").value = id
            ws.Cells(i, "Q").value = "���i"
            ws.Cells(i, "R").value = "���i"
        End If
    Next i
End Sub

Sub CreateInspectionSheetIDs()
    Dim wsSpecSheet As Worksheet
    Set wsSpecSheet = ThisWorkbook.Sheets("Hel_SpecSheet")

    Dim wsSetting As Worksheet
    Set wsSetting = ThisWorkbook.Sheets("Setting")

    Dim lastRow As Long
    lastRow = wsSpecSheet.Cells(wsSpecSheet.Rows.count, "D").End(xlUp).row

    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    For i = 2 To lastRow
        If wsSpecSheet.Cells(i, "D").value <> "" Then
            wsSpecSheet.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""����"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""�ቷ"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""�Z����"", ""Wet"", """")))"
            Dim id As String
            id = wsSpecSheet.Cells(i, "D").value & "-" & wsSpecSheet.Cells(i, "S").value & "-" & Left(wsSpecSheet.Cells(i, "E").value, 1)

            foundMatch = False
            For j = 2 To wsSetting.Cells(wsSetting.Rows.count, "H").End(xlUp).row
                If wsSpecSheet.Cells(i, "D").value = wsSetting.Cells(j, "H").value Then
                    foundMatch = True
                    If InStr(wsSetting.Cells(j, "J").value, "x") > 0 Then
                        id = "F" & id
                    End If
                    Exit For
                End If
            Next j

            If Not foundMatch Then
                MsgBox "�G���[: D��̒l��Setting�V�[�g��H��ƈ�v���鍀�ڂ�����܂���B�����𒆎~���܂��B"
                Exit Sub
            End If

            wsSpecSheet.Cells(i, "B").value = id
            wsSpecSheet.Cells(i, "Q").value = "���i"
            wsSpecSheet.Cells(i, "R").value = "���i"
        End If
    Next i
End Sub

Sub InsertXLookupAndUpdateKColumn()
    ' "Hel_SpecSheet"�̓V�����Ԃ𒲐�����
    ' ���������V�����Ԃ̍s��"Changed"�����Ă킩��₷�����Ă���B
    Dim wsHelSpecSheet As Worksheet
    Dim wsSetting As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim formulaResult As Variant
    Dim kValue As Variant
    
    ' �V�[�g�̐ݒ�
    Set wsHelSpecSheet = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' D��̍ŏI�s��T��
    lastRow = wsHelSpecSheet.Cells(wsHelSpecSheet.Rows.count, "D").End(xlUp).row
    
    ' D���T�����A�l������e�s�ɑ΂��ď��������s
    For i = 2 To lastRow
        If wsHelSpecSheet.Cells(i, "D").value <> "" Then
            ' T���XLOOKUP�֐�����
            wsHelSpecSheet.Cells(i, "T").Formula = "=XLOOKUP(TEXT(Hel_SpecSheet!D" & i & ", ""0""), " & _
                "TEXT(Setting!$H$2:$H$49, ""0""), " & _
                "Setting!$I$2:$I$49, """")"

            ' XLOOKUP�֐��̌��ʂ��擾
            formulaResult = wsHelSpecSheet.Cells(i, "T").value
            
            ' K��̒l���擾
            kValue = wsHelSpecSheet.Cells(i, "K").value
            
            ' K��̒l����T��̒l�������āA���ʂ�K��ɑ��
            wsHelSpecSheet.Cells(i, "K").value = kValue - formulaResult
            
            ' U���'Changed'����
            wsHelSpecSheet.Cells(i, "U").value = "Changed"
        End If
    Next i
End Sub


Function HighlightDuplicateValues() As Boolean
    ' �V�[�g����ϐ��Œ�`
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"
    
    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundDuplicate As Boolean
    foundDuplicate = False ' ���l�������������ǂ����̃t���O��������
    
    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row
    
    ' �F�̃C���f�b�N�X��������
    Dim colorIndex As Integer
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�
    
    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            If ws.Cells(i, "H").value = ws.Cells(j, "H").value And ws.Cells(i, "H").value <> "" Then
                ' ���l�����Z�������������ꍇ�A�t���O��True�ɐݒ肵�A�Z���ɐF��h��
                foundDuplicate = True
                ws.Cells(i, "H").Interior.colorIndex = colorIndex
                ws.Cells(j, "H").Interior.colorIndex = colorIndex
                ws.Cells(i, "H").Interior.colorIndex = colorIndex ' ���l�����������Z���ɐF��h��
            End If
        Next j
        ' ���l�����������ꍇ�A���̐F�ɕύX
        If foundDuplicate And ws.Cells(i, "H").Interior.colorIndex <> xlNone Then
            colorIndex = colorIndex + 1
            ' �F�C���f�b�N�X�̍ő�l�𒴂��Ȃ��悤�Ƀ`�F�b�N
            If colorIndex > 56 Then colorIndex = 3 ' �F�C���f�b�N�X�����Z�b�g
        End If
    Next i
    
    ' ���l�����������Ȃ������ꍇ�AH��̃Z���̐F�𔒂ɐݒ�
    If Not foundDuplicate Then
        For i = 2 To lastRow
            ws.Cells(i, "H").Interior.Color = xlNone
        Next i
    End If
    
    ' ���l�������������ǂ����Ɋ�Â��Č��ʂ�Ԃ�
    HighlightDuplicateValues = foundDuplicate
End Function


Function LocateEmptySpaces() As String
    ' "Hel_SpecSheet"�ɋ󗓂܂��̓f�[�^�^�̌�肪�Ȃ������`�F�b�N
    ' �ϐ��錾�Ə�����
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    Dim lastCol As Long
    lastCol = ws.Columns("S").Column
    Dim errorMsg As String
    errorMsg = ""
    
    ' �w��͈͂����[�v���ăG���[�`�F�b�N
    For i = 2 To lastRow
        For j = 2 To lastCol
            Dim cell As Range
            Set cell = ws.Cells(i, j)
            ' �󔒂̃`�F�b�N
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "�󔒃Z��: " & cell.Address(False, False) & vbNewLine
            End If
            ' ���l�`�F�b�N
            If (j = 7 Or j = 8 Or j = 10 Or j = 11) And Not IsNumeric(cell.value) Then
                errorMsg = errorMsg & "���l�łȂ��Z��: " & cell.Address(False, False) & vbNewLine
            End If
            ' ������`�F�b�N
'            If (j = 14 Or j = 15 Or j = 16) And Not VarType(cell.Value) = vbString Then
'                errorMsg = errorMsg & "������łȂ��Z��: " & cell.Address(False, False) & vbNewLine
'            End If
        Next j
    Next i
    
    ' �G���[���b�Z�[�W������΂����Ԃ��A�Ȃ���΋�̕������Ԃ�
    LocateEmptySpaces = errorMsg
End Function



Sub CopyDataBasedOnCondition()
    'SpecSheet�̓��e��Log�V�[�g�ɓ]�L����
    Dim logSheet As Worksheet
    Dim helSpec As Worksheet
    Dim lastRowLog As Long
    Dim lastRowSpec As Long
    Dim i As Long, j As Long
    Dim matchCount As Long

    ' ���[�N�V�[�g���Z�b�g
    Set logSheet = ThisWorkbook.Worksheets("LOG_Helmet")
    Set helSpec = ThisWorkbook.Worksheets("Hel_SpecSheet")

    ' LOG_Helmet�̍ŏI�s���擾
    lastRowLog = logSheet.Cells(logSheet.Rows.count, "H").End(xlUp).row
    ' Hel_SpecSheet�̍ŏI�s���擾
    lastRowSpec = helSpec.Cells(helSpec.Rows.count, "H").End(xlUp).row

    ' LOG_Helmet��H��̒l�𐮂���
'    For i = 2 To lastRowLog
'        logSheet.Cells(i, "H").Value = Application.Round(logSheet.Cells(i, "H").Value, 2)
'    Next i

    ' �l���r���ē]�L
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").value = helSpec.Cells(j, "H").value Then
                ' H��̒l����v�����ꍇ�A�e��̓��e��]�L
                matchCount = matchCount + 1
                logSheet.Cells(i, "B").value = helSpec.Cells(j, "B").value
                logSheet.Cells(i, "C").value = helSpec.Cells(j, "B").value
                logSheet.Cells(i, "D").value = helSpec.Cells(j, "D").value
                logSheet.Cells(i, "E").value = helSpec.Cells(j, "E").value
                logSheet.Cells(i, "F").value = helSpec.Cells(j, "F").value
                logSheet.Cells(i, "G").value = helSpec.Cells(j, "G").value
                logSheet.Cells(i, "L").value = helSpec.Cells(j, "I").value
                logSheet.Cells(i, "M").value = helSpec.Cells(j, "J").value
                logSheet.Cells(i, "N").value = helSpec.Cells(j, "K").value '�V��������
                logSheet.Cells(i, "O").value = helSpec.Cells(j, "L").value
                logSheet.Cells(i, "U").value = helSpec.Cells(j, "M").value '�������e
                logSheet.Cells(i, "P").value = helSpec.Cells(j, "N").value '�������b�g
                logSheet.Cells(i, "Q").value = helSpec.Cells(j, "O").value
                logSheet.Cells(i, "R").value = helSpec.Cells(j, "P").value
                logSheet.Cells(i, "S").value = helSpec.Cells(j, "Q").value '�\������
                logSheet.Cells(i, "T").value = helSpec.Cells(j, "R").value
                'logSheet.Cells(i, "U").Value = helSpec.Cells(j, "S").Value
                'logSheet.Cells(i, "U").Value = helSpec.Cells(j, "U").Value
                
            End If
        Next j
        
        ' ��v�����l���������݂���ꍇ�A�����𑾎��ɂ���
        If matchCount > 1 Then
            logSheet.Cells(i, "C").Font.Bold = True
            logSheet.Cells(i, "D").Font.Bold = True
            logSheet.Cells(i, "E").Font.Bold = True
            logSheet.Cells(i, "F").Font.Bold = True
            logSheet.Cells(i, "G").Font.Bold = True
            logSheet.Cells(i, "L").Font.Bold = True
            logSheet.Cells(i, "M").Font.Bold = True
            logSheet.Cells(i, "N").Font.Bold = True
            logSheet.Cells(i, "O").Font.Bold = True
        End If
    Next i
End Sub


Sub CustomizeSheetFormats()
' �e��ɏ����ݒ������
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "�ő�l(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.00 ""kN"""
            ElseIf InStr(1, cell.value, "�ő�l(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0 ""G"""
            ElseIf InStr(1, cell.value, "����") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""ms"""
            ElseIf InStr(1, cell.value, "���x") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""��"""
            ElseIf InStr(1, cell.value, "�d��") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""g"""
            ElseIf InStr(1, cell.value, "���b�g") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "@"
            ElseIf InStr(1, cell.value, "�V��������") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""mm"""
            End If

        Next cell
    Next sheet
End Sub

Sub UniformizeLineGraphAxes()

    ' Display input dialog to set the maximum value for the axes
    Dim MaxValue As Double
    MaxValue = InputBox("Y���̍ő�l����͂��Ă��������B(����)", "�ő�l�����")
    
    ' Loop through all the charts in the active sheet
    Dim chartObj As ChartObject
    For Each chartObj In ActiveSheet.ChartObjects
        With chartObj.chart.Axes(xlValue)
            ' Set the Y-axis maximum value
            .MaximumScale = MaxValue
        End With
    Next chartObj

End Sub

