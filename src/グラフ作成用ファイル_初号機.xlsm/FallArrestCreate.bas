Attribute VB_Name = "FallArrestCreate"
Public Sub Create_FallArrestGraph()
    Call CreateGraphFallArrest
    Call FallArrest_2kN_DurationTime
End Sub


Function ColNumToLetter(colNum As Integer) As String
    ' CreateGraphFallArrest�ɂĎg�p����֐�
    Dim d As Integer
    Dim m As Integer
    Dim name As String
    d = colNum
    name = ""
    While (d > 0)
        m = (d - 1) Mod 26
        name = Chr(65 + m) & name
        d = Int((d - m) / 26)
    Wend
    ColNumToLetter = name
End Function

Sub CreateGraphFallArrest()

    ' ���[�N�V�[�g��錾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' �ŏI�s�ƍŏI�������
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    Dim i As Long
    Dim maxVal As Double
    Dim colStart As String
    Dim colEnd As String

    ' P��(16�Ԗ�) + 100�񂩂�n�܂��
    colStart = ColNumToLetter(16 + 100)

    ' P��(16�Ԗ�) + 800�񂩂�n�܂��
    colEnd = ColNumToLetter(16 + 1200)

    ' �����̃`���[�g�̈ʒu
    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' �ő�l�����߂�
        maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

        ' K��ɍő�l��\��
        ws.Cells(i, "G").Value = maxVal

        ' �`���[�g���쐬
        Dim ChartObj As ChartObject
        Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
        Dim chart As chart
        Set chart = ChartObj.chart

        ' �܂���O���t��ݒ�
        chart.ChartType = xlLine

        ' �O���t�̃f�[�^�͈͂�ݒ�
        chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))

        ' �O���t�̃^�C�g����ݒ�
        chart.HasTitle = True
        chart.ChartTitle.Text = ws.Cells(i, "B").Value

        ' �n��̕\�����I�t
        chart.SetElement msoElementLegendNone

        ' ���̑�����ݒ�
        chart.SeriesCollection(1).Format.Line.Weight = 0.75

        ' Y���̐ݒ�
        Dim yAxis As Axis
        Set yAxis = chart.Axes(xlValue, xlPrimary)

        ' Y���̍ő�l��ݒ�
        If maxVal <= 4.95 Then
            yAxis.MaximumScale = 5
        ElseIf maxVal > 4.95 And maxVal <= 9.81 Then
            yAxis.MaximumScale = 10
        Else
            yAxis.MaximumScale = Int(maxVal) + 1
        End If

        yAxis.MinimumScale = -1 ' Y���̍Œ�l��-1�ɐݒ肵�܂��B

        ' Y���� TickLabels ��ݒ�
        With yAxis.TickLabels
            .NumberFormatLocal = "0.0""kN"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With


        ' X���̐ݒ�
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 200
        xAxis.TickMarkSpacing = 50


        ' X���� TickLabels ��ݒ�
        With xAxis.TickLabels
            .NumberFormatLocal = "0""ms"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With
        
        ' �`���[�g�̈ʒu�����ɍX�V
        chartLeft = chartLeft + 10

    Next i

End Sub

Sub FallArrest_2kN_DurationTime()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim lastRow As Long
    Dim cellVal As Double
    Dim startRange As Long
    Dim endRange As Long
    Dim maxRange As Double
    Dim sumRange As Double
    Dim countRange As Long
    Dim rangeCollection As Collection
    Dim item As Variant
    Dim currentColor As Long

    ' "LOG"�V�[�g���w�肵�܂��B
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' �ŏI�s���擾���܂��B
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' �S�s�ɂ��ď������s���܂��B
    For i = 2 To lastRow
        ' ���݂̍s�͈̔͂�ݒ肵�܂��B
        Set rng = ws.Range(ws.Cells(i, "P"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' �͈͂̏��������s���܂��B
        startRange = 0
        endRange = 0
        maxRange = 0
        sumRange = 0
        countRange = 0
        ' �R���N�V�����̏��������s���܂��B
        Set rangeCollection = New Collection
        ' �e�s�ōő�l�������A"F"��ɋL�����܂��B
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' ���݂̍s���̊e�Z�����`�F�b�N���܂��B
        For Each cell In rng
            cellVal = cell.Value
            ' �l��2.2�ȏ�Ȃ�͈͂��X�V���A�F���p�����܂��B
            If cellVal >= 2.2 Then
                If startRange = 0 Then
                    startRange = cell.column
                    currentColor = RGB(Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0))
                End If
                endRange = cell.column
                cell.Interior.color = currentColor
                sumRange = sumRange + cellVal
                countRange = countRange + 1
            ' ����ȊO�Ȃ�͈͂��R���N�V�����ɕۑ����A�͈͂����Z�b�g���܂��B
            Else
                If startRange > 0 And endRange > 0 Then
                    rangeCollection.Add Array(startRange, endRange)
                    startRange = 0
                    endRange = 0
                End If
            End If
        Next cell

        ' �c�����͈͂��R���N�V�����ɒǉ����܂��B
        If startRange > 0 And endRange > 0 Then rangeCollection.Add Array(startRange, endRange)

        ' 2.2�ȏ�̒l���擾���A���̍��v�̕��ϒl��(k,i)�ɕ\�����܂��B
        ' 2.2�ȉ��̏ꍇ�́A(k,i)��0��\�����܂��B8.0kN�𒴂���ꍇ�͂��̍ő�l��(k,i)�ɕ\�����܂��B
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If
    Next i
    
    ' ��̃Z���͑S��"-"�Ŗ��߂�
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    Dim cellRng As Range
    For Each cellRng In ws.Range("F2:O" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng
End Sub

Sub FallArrest_2kN_DurationTime_0915()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim lastRow As Long
    Dim cellVal As Double
    Dim startPoint As Long
    Dim endPoint As Long
    Dim maxRange As Double
    Dim sumRange As Double
    Dim countRange As Long

    ' "LOG"�V�[�g���w�肵�܂��B
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' �ŏI�s���擾���܂��B
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' �S�s�ɂ��ď������s���܂��B
    For i = 2 To lastRow
        ' ���݂̍s�͈̔͂�ݒ肵�܂��B
        Set rng = ws.Range(ws.Cells(i, "Q"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' �͈͂̏��������s���܂��B
        startPoint = 0
        endPoint = 0
        maxRange = 0
        sumRange = 0
        countRange = 0

        ' �e�s�ōő�l�������A"F"��ɋL�����܂��B
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' ���݂̍s���̊e�Z�����`�F�b�N���܂��B
        For Each cell In rng
            cellVal = cell.Value
            ' �l��2.2�ȏ�Ȃ�͈͂��X�V���܂��B
            If cellVal >= 2.2 Then
                If startPoint = 0 Then
                    startPoint = cell.column
                End If
                endPoint = cell.column
                sumRange = sumRange + cellVal
                countRange = countRange + 1
            End If
        Next cell

        ' startPoint��endPoint�̊Ԃ̗�̊Y���s�̒l�����v���A���̌��ʂ�L��ɓ��͂��܂��B
        If startPoint > 0 And endPoint > 0 Then
            Dim sumAbs As Double: sumAbs = 0
            Dim j As Long
            
            For j = startPoint To endPoint
                sumAbs = sumAbs + Abs(ws.Cells(i, j).Value)
            Next j
            
            ws.Cells(i, "L").Value = sumAbs / (endPoint - startPoint + 1)
        End If

        ' 2.2�ȏ�̒l���擾���A���̍��v�̕��ϒl��(k,i)�ɕ\�����܂��B
        ' 2.2�ȉ��̏ꍇ�́A(k,i)��0��\�����܂��B8.0kN�𒴂���ꍇ�͂��̍ő�l��(k,i)�ɕ\�����܂��B
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If

    Next i
    
    ' ��̃Z���͑S��"-"�Ŗ��߂�
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    Dim cellRng As Range
    For Each cellRng In ws.Range("F2:O" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng

End Sub

Sub FallArrest_2kN_DurationTime_1135()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim lastRow As Long
    Dim cellVal As Double
    Dim startPoint As Long
    Dim endPoint As Long
    Dim maxRange As Double
    Dim sumRange As Double
    Dim countRange As Long
    Dim currentColor As Long

    ' "LOG"�V�[�g���w�肵�܂��B
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' �ŏI�s���擾���܂��B
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' �S�s�ɂ��ď������s���܂��B
    For i = 2 To lastRow
        ' ���݂̍s�͈̔͂�ݒ肵�܂��B
        Set rng = ws.Range(ws.Cells(i, "Q"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' �͈͂̏��������s���܂��B
        startPoint = 0
        endPoint = 0
        maxRange = 0
        sumRange = 0
        countRange = 0

        ' �e�s�ōő�l�������A"F"��ɋL�����܂��B
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' ���݂̍s���̊e�Z�����`�F�b�N���܂��B
        For Each cell In rng
            cellVal = cell.Value
            ' �l��2.2�ȏ�Ȃ�͈͂��X�V���A�F���p�����܂��B
            If cellVal >= 2.2 Then
                If startPoint = 0 Then
                    startPoint = cell.column
                    currentColor = RGB(Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0))
                End If
                endPoint = cell.column
                cell.Interior.color = currentColor
                sumRange = sumRange + cellVal
                countRange = countRange + 1
            End If
        Next cell

        ' startPoint��endPoint�̊Ԃ̗�̊Y���s�̒l�����v���A���̌��ʂ�L��ɓ��͂��܂��B
        If startPoint > 0 And endPoint > 0 Then
            Dim sumAbs As Double: sumAbs = 0
            Dim j As Long
            
            For j = startPoint To endPoint
                sumAbs = sumAbs + Abs(ws.Cells(i, j).Value)
            Next j
            
            ws.Cells(i, "L").Value = sumAbs / (endPoint - startPoint + 1)
        End If

        ' 2.2�ȏ�̒l���擾���A���̍��v�̕��ϒl��(k,i)�ɕ\�����܂��B
        ' 2.2�ȉ��̏ꍇ�́A(k,i)��0��\�����܂��B8.0kN�𒴂���ꍇ�͂��̍ő�l��(k,i)�ɕ\�����܂��B
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If

    Next i
    
    ' ��̃Z���͑S��"-"�Ŗ��߂�
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    Dim cellRng As Range
    For Each cellRng In ws.Range("F2:O" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng

End Sub




' �A�N�e�B�u�V�[�g���̃O���t���폜
Sub DeleteAllChartsInActiveSheet()
    Dim chart As ChartObject
    
    For Each chart In ActiveSheet.ChartObjects
        chart.Delete
    Next chart
End Sub

