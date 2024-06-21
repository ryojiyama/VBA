Attribute VB_Name = "SpecSheet"
Sub CreateIDplusDrawBorders()
    '�]�L��Ƃƃt�H�[�}�b�g�����A�A�C�R���ɕR�Â�
    Call CreateID
    Call DrawBordersWithHairline
End Sub

Sub CreateID()
    '�i�ԁA�����ӏ��Ȃǂɉ�����ID���쐬����
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim ID As String
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' �Ō�̍s���擾
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' �e�s�ɑ΂���ID�𐶐�
    For i = 2 To lastRow ' 1�s�ڂ̓w�b�_�Ɖ���
        
        ' C��: 2���ȉ��̐���
        If Len(ws.Cells(i, 3).Value) <= 2 Then
            ID = Right("00" & ws.Cells(i, 3).Value, 2)
        Else
            ID = "??"
        End If
        
        ' C���D��̊Ԃ�"-"
        ID = ID & "-"
        
        ' D��: ������4�����ڂ���6�����ڂ̕�����
        ID = ID & Mid(ws.Cells(i, 4).Value, 4, 3)
        
        ' E��̏���
        Select Case ws.Cells(i, 5).Value
            Case "�V��"
                ID = ID & "T"
            Case "�O����"
                ID = ID & "F"
            Case "�㓪��"
                ID = ID & "R"
            Case Else
                ID = ID & "?"
        End Select
        
        ' I��̏���
        Select Case ws.Cells(i, 9).Value
            Case "����"
                ID = ID & "H"
            Case "�ቷ"
                ID = ID & "L"
            Case "�Z����"
                ID = ID & "W"
            Case Else
                ID = ID & "?"
        End Select
        
        ' I���L��̊Ԃ�"-"
        ID = ID & "-"
        
        ' L��̏���
        If ws.Cells(i, 12).Value = "��" Then
            ID = ID & "W"
        Else
            ID = ID & "O"
        End If
        
        ' B���ID���Z�b�g
        ws.Cells(i, 2).Value = ID
    Next i

End Sub

Sub DrawBordersWithHairline()
    ' �V�[�g�uHel_SpecSheet�v��I��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' �V�[�g�̊����̌r����S�ď���
    ws.Cells.Borders.LineStyle = xlNone

    ' C��̍ŏI�s��T��
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' �͈� Cells(2, "B"):Cells(lastRow, "M") �ɐV���Ɍr���������i1�s�ڂ͏��O�j
    With ws.Range(ws.Cells(2, "B"), ws.Cells(lastRow, "M"))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlHairline
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlHairline

        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlHairline
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
End Sub



Sub SyncSpecSheetToLogHel()
    ' �A�C�R���ɕR�Â��BSpecSheet�ɓ]�L����v���V�[�W���̂܂Ƃ�
    ' ���l�����������ꍇ�̓G���[���b�Z�[�W��\�����ď����𒆒f
    If HighlightDuplicateValues Then
        MsgBox "�Ռ��l�œ��l��������܂����B�����_���񌅂ɉe�����o�Ȃ��͈͂ŏC�����Ă��������B", vbCritical
        Exit Sub
    End If

    ' �\�ɋ󗓂�����ꍇ�ɃG���[���b�Z�[�W���o���Ē��f
    If Not LocateEmptySpaces Then
        MsgBox "�󗓂�����܂��B�܂��͂���𖄂߂Ă��������B", vbCritical
        Exit Sub
    End If

    Call CopyDataBasedOnCondition
    Call CustomizeSheetFormats
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
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' �F�̃C���f�b�N�X��������
    Dim colorIndex As Integer
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�

    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            If ws.Cells(i, "H").Value = ws.Cells(j, "H").Value And ws.Cells(i, "H").Value <> "" Then
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
            ws.Cells(i, "H").Interior.color = xlNone
        Next i
    End If

    ' ���l�������������ǂ����Ɋ�Â��Č��ʂ�Ԃ�
    HighlightDuplicateValues = foundDuplicate
End Function

Function LocateEmptySpaces() As Boolean
    ' "Hel_SpecSheet"�ɋ󗓂��Ȃ������`�F�b�N
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim cell As Range
    Dim errorMsg As String

    ' �G���[���b�Z�[�W�p�̕������������
    errorMsg = ""

    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' �ŏI���"M"(�����敪)�ɌŒ�
    Dim lastCol As Long
    lastCol = ws.Columns("M").column

    ' �w��͈͂����[�v
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' �󔒂̃`�F�b�N
            If IsEmpty(cell.Value) Then
                errorMsg = errorMsg & "�󔒃Z��: " & cell.Address(False, False) & vbNewLine
            End If

            ' ��G�AH�AJ�AK�Ő��l�̊m�F
            If j = Columns("G").column Or j = Columns("H").column Or j = Columns("J").column Or j = Columns("K").column Then
                If Not IsNumeric(cell.Value) Then
                    errorMsg = errorMsg & "���l�łȂ��Z��: " & cell.Address(False, False) & vbNewLine
                End If
            End If

            ' ��N�AO�AP�ŕ�����̊m�F
            If j = Columns("N").column Or j = Columns("O").column Or j = Columns("P").column Then
                If Not VarType(cell.Value) = vbString Then
                    errorMsg = errorMsg & "������łȂ��Z��: " & cell.Address(False, False) & vbNewLine
                End If
            End If
        Next j
    Next i

    ' �G���[���b�Z�[�W������Ε\�����AFalse��Ԃ�
    If Len(errorMsg) > 0 Then
        LocateEmptySpaces = False
        Exit Function
    Else
        LocateEmptySpaces = True
    End If
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
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).row
    ' Hel_SpecSheet�̍ŏI�s���擾
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).row

    ' LOG_Helmet��H��̒l�𐮂���
'    For i = 2 To lastRowLog
'        logSheet.Cells(i, "H").Value = Application.Round(logSheet.Cells(i, "H").Value, 2)
'    Next i

    ' �l���r���ē]�L
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").Value = helSpec.Cells(j, "H").Value Then
                ' H��̒l����v�����ꍇ�A�e��̓��e��]�L
                matchCount = matchCount + 1
                logSheet.Cells(i, "C").Value = helSpec.Cells(j, "B").Value
                logSheet.Cells(i, "D").Value = helSpec.Cells(j, "D").Value
                logSheet.Cells(i, "E").Value = helSpec.Cells(j, "E").Value
                logSheet.Cells(i, "F").Value = helSpec.Cells(j, "F").Value
                logSheet.Cells(i, "G").Value = helSpec.Cells(j, "G").Value
                logSheet.Cells(i, "L").Value = helSpec.Cells(j, "I").Value
                logSheet.Cells(i, "M").Value = helSpec.Cells(j, "J").Value
                logSheet.Cells(i, "N").Value = helSpec.Cells(j, "K").Value
                logSheet.Cells(i, "O").Value = helSpec.Cells(j, "L").Value
                logSheet.Cells(i, "U").Value = helSpec.Cells(j, "M").Value
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
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

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
            ElseIf InStr(1, cell.Value, "�V��������") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
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
    Dim ChartObj As ChartObject
    For Each ChartObj In ActiveSheet.ChartObjects
        With ChartObj.chart.Axes(xlValue)
            ' Set the Y-axis maximum value
            .MaximumScale = MaxValue
        End With
    Next ChartObj

End Sub

