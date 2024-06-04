Attribute VB_Name = "HelmetCreate"
Public Sub Create_HelmetGraph()
    Call CreateGraphHelmet
    Call InspectHelmetDurationTime
    Call HighlightDuplicateValues
End Sub
Sub VisualizeSelectedData_HelmetGraph()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    Dim chartColors As Variant
    chartColors = Array(RGB(47, 85, 151), RGB(241, 88, 84), RGB(111, 178, 85), _
                    RGB(250, 194, 58), RGB(158, 82, 143), RGB(255, 127, 80), _
                    RGB(250, 159, 137), RGB(72, 61, 139))
                        
    Dim colorIndex As Integer
    colorIndex = 0

    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100

    Dim colStart As String
    Dim colEnd As String
    colStart = ColNumToLetter(16 + 52) 'Excel�̃��x����BP
    colEnd = ColNumToLetter(16 + 850) 'Excel�̃��x����AGH

    Dim ChartObj As ChartObject
    Dim chart As chart
    Dim maxVal As Double

    For i = 2 To lastRow
        If ws.Cells(i, "B").Interior.color = RGB(252, 228, 214) Then
            maxVal = Application.WorksheetFunction.Max(ws.Range(colStart & i & ":" & colEnd & i))

            If Not ChartObj Is Nothing Then
                Dim series As series
                Set series = chart.SeriesCollection.NewSeries
                series.Values = ws.Range(colStart & i & ":" & colEnd & i)
                series.XValues = ws.Range(colStart & "1:" & colEnd & "1")
                series.Format.Line.ForeColor.RGB = chartColors(colorIndex)
                series.name = ws.Cells(i, "D").Value & " - " & ws.Cells(i, "L").Value
            Else
                Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
                Set chart = ChartObj.chart
                
                chart.ChartType = xlLine
                chart.SetSourceData Source:=ws.Range(colStart & i & ":" & colEnd & i)
                chart.SeriesCollection(1).Format.Line.ForeColor.RGB = chartColors(colorIndex)
                chart.SeriesCollection(1).XValues = ws.Range(colStart & "1:" & colEnd & "1")
                chart.SeriesCollection(1).name = ws.Cells(i, "D").Value & " - " & ws.Cells(i, "L").Value
            End If
            
        ' ���̑�����ݒ�
        chart.SeriesCollection(1).Format.Line.Weight = 1#

        ' Y���̐ݒ�
        Dim yAxis As Axis
        Set yAxis = chart.Axes(xlValue, xlPrimary)

        yAxis.MinimumScale = -1 ' Y���̍Œ�l��0�ɐݒ肵�܂��B

        ' Y���� TickLabels ��ݒ�
        With yAxis.TickLabels
            .NumberFormatLocal = "0.0""kN"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With

        ' X���̐ݒ�
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 100
        xAxis.TickMarkSpacing = 25


        ' X���� TickLabels ��ݒ�
        With xAxis.TickLabels
            .NumberFormatLocal = "0.00""ms"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With
            
            colorIndex = (colorIndex + 1) Mod UBound(chartColors)

        End If
    Next i

End Sub
'----------------------------------------------------------------------------------
Function ColNumToLetter(colNum As Integer) As String
    Dim d As Integer, m As Integer, name As String
    d = colNum
    name = ""
    While d > 0
        m = (d - 1) Mod 26
        name = Chr(65 + m) & name
        d = Int((d - m) / 26)
    Wend
    ColNumToLetter = name
End Function

Sub CreateGraphHelmet()

    Const START_OFFSET As Long = 16
    Const START_EXTENSION As Long = 52
    Const END_EXTENSION As Long = 850

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    Dim colStart As String
    Dim colEnd As String

    colStart = ColNumToLetter(START_OFFSET + START_EXTENSION)
    colEnd = ColNumToLetter(START_OFFSET + END_EXTENSION)

    Dim chartLeft As Long
    Dim chartTop As Long
    chartTop = ws.Rows(lastRow).Height - 20
    chartLeft = 250

    For i = 2 To lastRow
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd
        chartLeft = chartLeft + 10
    Next i

End Sub

Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String)
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

    ws.Cells(i, "H").Value = maxVal

    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
    Dim chart As chart
    Set chart = ChartObj.chart

    ConfigureChart chart, ws, i, colStart, colEnd, maxVal

End Sub

Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.ChartTitle.Text = ws.Cells(i, "B").Value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75

    SetYAxis chart, maxVal
    SetXAxis chart

End Sub

Sub SetYAxis(ByRef chart As chart, ByVal maxVal As Double)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    If maxVal <= 4.95 Then
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1#
    ElseIf maxVal <= 9.81 Then
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2#
    Else
        yAxis.MaximumScale = Int(maxVal) + 1
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub

Sub SetXAxis(ByRef chart As chart)
    Dim xAxis As Axis
    Set xAxis = chart.Axes(xlCategory, xlPrimary)
    xAxis.TickLabelSpacing = 100
    xAxis.TickMarkSpacing = 25

    With xAxis.TickLabels
        .NumberFormatLocal = "0.00""ms"""
        .Font.color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub
'----------------------------------------------------------------------------------

Function ColNumToLetter_Old(colNum As Integer) As String
    ' CreateGraphHelmet�ɂĎg�p����֐�
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


Sub CreateGraphHelmet_Old()

    ' ���[�N�V�[�g��錾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' �ŏI�s�ƍŏI�������
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    Dim i As Long
    Dim maxVal As Double
    Dim colStart As String
    Dim colEnd As String

    ' P��(16�Ԗ�) + 52�񂩂�n�܂��
    colStart = ColNumToLetter_Old(16 + 52)

    ' P��(16�Ԗ�) + 800�񂩂�n�܂��
    colEnd = ColNumToLetter_Old(16 + 850)

    ' �����̃`���[�g�̈ʒu
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim finalRowHeight As Long
    finalRowHeight = ws.Rows(lastRow).Height

    chartLeft = 250
    chartTop = finalRowHeight - 20

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' �ő�l�����߂�
        maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

        ' K��ɍő�l��\��
        ws.Cells(i, "H").Value = maxVal

        ' �`���[�g���쐬
        Dim ChartObj As ChartObject
        Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
        Dim chart As chart
        Set chart = ChartObj.chart

        ' �܂���O���t��ݒ�
        chart.ChartType = xlLine

        ' �O���t�̃f�[�^�͈͂�ݒ�
        chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
        
        ' X���̃f�[�^�͈͂�ݒ�
        chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))

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

        yAxis.MinimumScale = -1 ' Y���̍Œ�l��0�ɐݒ肵�܂��B

        ' Y���� TickLabels ��ݒ�
        With yAxis.TickLabels
            .NumberFormatLocal = "0.0""kN"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With


        ' X���̐ݒ�
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 100
        xAxis.TickMarkSpacing = 25


        ' X���� TickLabels ��ݒ�
        With xAxis.TickLabels
            .NumberFormatLocal = "0.00""ms"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With

        ' �`���[�g�̈ʒu�����ɍX�V
        chartLeft = chartLeft + 10

    Next i

End Sub
'----------------------------------------------------------------------------------


Sub InspectHelmetDurationTime()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' "LOG_Helmet" �V�[�g���w�肷��B
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' �ŏI�s���擾����B
    lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).row

    ' �e�s����������B
    For i = 2 To lastRow
        UpdateMaxValueInRow ws, i             ' �s���̍ő�l���X�V����B
        UpdatePartOfHelmet ws, i              ' �w�����b�g�̕������X�V����B
        UpdateRangeForThresholds ws, i, 4.9, "J"  ' 臒l�͈̔͂��X�V����B
        UpdateRangeForThresholds ws, i, 7.35, "K"
    Next i

    ' ��̃Z���𖄂߂�B
    FillEmptyCells ws, lastRow
End Sub

Sub UpdateMaxValueInRow(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�s���̍ő�l���X�V����
    Dim rng As Range
    Dim MaxValue As Double
    Dim maxValueColumn As Long

    ' �s���͈̔͂��Z�b�g����B
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))
    ' �ő�l���擾����B
    MaxValue = Application.WorksheetFunction.Max(rng)
    ws.Cells(row, "H").Value = MaxValue

    ' �ő�l�̈ʒu��������B
    For j = 1 To rng.Columns.Count
        If rng(1, j).Value = MaxValue Then
            maxValueColumn = j
            rng(1, j).Interior.color = RGB(250, 150, 0)  ' �F��ݒ肷��B
            Exit For
        End If
    Next j

    If maxValueColumn <> 0 Then
        ws.Cells(row, "I").Value = ws.Cells(1, "V").Offset(0, maxValueColumn - 1).Value
    End If
End Sub

Sub UpdatePartOfHelmet(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�w�����b�g�̎����ӏ����X�V����
    '�ŏ��ɎQ�Ƃ���̂�B��̒l
    Dim cellValue As String
    cellValue = ws.Cells(row, "B").Value
    
    ' �����̒l���擾����B
    Dim existingValue As String
    existingValue = ws.Cells(row, "E").Value

    ' �w�����b�g�̕����Ɋ�Â��Ēl��ݒ肷��B�������A"�V��"��"����"�����łɊ܂܂�Ă���ꍇ�͕ύX���Ȃ��B
    ' �����߂ł͍ŏ���E��̒l���`�F�b�N����B
    If InStr(existingValue, "�V��") > 0 Or InStr(existingValue, "����") > 0 Then
        ' �ύX���Ȃ�
    ElseIf InStr(cellValue, "HEL_TOP") > 0 Then
        ws.Cells(row, "E").Value = "�V��"
    ElseIf InStr(cellValue, "HEL_ZENGO") > 0 Then
        ws.Cells(row, "E").Value = "�O�㓪��"
    End If
End Sub


Sub UpdateRangeForThresholds(ByRef ws As Worksheet, ByVal row As Long, ByVal threshold As Double, ByVal columnToWrite As String)
    'InspectHelmetDurationTime()����4.9�A7.35�͈̔͒l�̐F�t���ƏՌ����Ԃ��L������B
    Dim rng As Range, cell As Range
    Dim startRange As Long, endRange As Long, maxRange As Long
    Dim rangeCollection As New Collection
    Dim timeDifference As Double

    ' �s�͈̔͂��Z�b�g����B
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))

    ' 臒l�𒴂���Z���͈̔͂�������B
    For Each cell In rng
        If cell.Value >= threshold Then
            If startRange = 0 Then startRange = cell.column
            endRange = cell.column
            cell.Interior.color = IIf(threshold = 4.9, RGB(255, 111, 56), RGB(234, 67, 53))
        Else
            If startRange > 0 And endRange > 0 Then
                rangeCollection.Add Array(startRange, endRange)
                startRange = 0
                endRange = 0
            End If
        End If
    Next cell

    If startRange > 0 And endRange > 0 Then rangeCollection.Add Array(startRange, endRange)

    For Each item In rangeCollection
        If (item(1) - item(0) + 1) > maxRange Then
            maxRange = item(1) - item(0) + 1
            startRange = item(0)
            endRange = item(1)
        End If
    Next item

    If startRange > 0 And endRange > 0 Then
        timeDifference = ws.Cells(1, endRange).Value - ws.Cells(1, startRange).Value
        ws.Cells(row, columnToWrite).Value = timeDifference
    Else
        ws.Cells(row, columnToWrite).Value = "-"
    End If
End Sub

Sub FillEmptyCells(ByRef ws As Worksheet, ByVal lastRow As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�󗓂�-������B
    Dim cellRng As Range

    ' B��̍Ō�̍s���擾����B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' ��̃Z����"-"�Ŗ��߂�B
    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng
End Sub

Sub HighlightDuplicateValues()
    ' �Ռ��l�����ꍇ������̂ł�����`�F�b�N����v���V�[�W��
    Dim sheetName As String
    sheetName = "LOG_Helmet"
    
    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    
    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row
    
    ' �F�̃C���f�b�N�X��������
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�
    
    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' ���݂̃Z���̒l���擾
        valueToFind = ws.Cells(i, "H").Value
        
        ' �����l�����Z�������ɐF�t������Ă��Ȃ����`�F�b�N
        If ws.Cells(i, "H").Interior.colorIndex = xlNone Then
            ' ���݂̃Z���̒l�Ɠ����l�����Z����T��
            For j = i + 1 To lastRow
                If ws.Cells(j, "H").Value = valueToFind And ws.Cells(j, "H").Interior.colorIndex = xlNone Then
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
End Sub


' �A�N�e�B�u�V�[�g���̃O���t���폜
Sub DeleteAllChartsInActiveSheet()
    Dim chart As ChartObject
    
    For Each chart In ActiveSheet.ChartObjects
        chart.Delete
    Next chart
End Sub



Sub UpdatePartOfHelmet_231013SyuuseiMae(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�w�����b�g�̎����ӏ����X�V����
    Dim cellValue As String
    cellValue = ws.Cells(row, "B").Value

    ' �w�����b�g�̕����Ɋ�Â��Ēl��ݒ肷��B
    If InStr(cellValue, "TOP") > 0 Then
        ws.Cells(row, "E").Value = "�V��"
    ElseIf InStr(cellValue, "MAE") > 0 Then
        ws.Cells(row, "E").Value = "�O����"
    ElseIf InStr(cellValue, "USHIRO") > 0 Then
        ws.Cells(row, "E").Value = "�㓪��"
    Else
        ws.Cells(row, "E").Value = "�O�㓪��"
    End If
End Sub
