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
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row

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

    Dim chartObj As ChartObject
    Dim chart As chart
    Dim maxVal As Double

    For i = 2 To lastRow
        If ws.Cells(i, "B").Interior.Color = RGB(252, 228, 214) Then
            maxVal = Application.WorksheetFunction.Max(ws.Range(colStart & i & ":" & colEnd & i))

            If Not chartObj Is Nothing Then
                Dim series As series
                Set series = chart.SeriesCollection.NewSeries
                series.Values = ws.Range(colStart & i & ":" & colEnd & i)
                series.XValues = ws.Range(colStart & "1:" & colEnd & "1")
                series.Format.Line.ForeColor.RGB = chartColors(colorIndex)
                series.name = ws.Cells(i, "D").value & " - " & ws.Cells(i, "L").value
            Else
                Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=425, Top:=chartTop, Height:=225)
                Set chart = chartObj.chart
                
                chart.ChartType = xlLine
                chart.SetSourceData Source:=ws.Range(colStart & i & ":" & colEnd & i)
                chart.SeriesCollection(1).Format.Line.ForeColor.RGB = chartColors(colorIndex)
                chart.SeriesCollection(1).XValues = ws.Range(colStart & "1:" & colEnd & "1")
                chart.SeriesCollection(1).name = ws.Cells(i, "D").value & " - " & ws.Cells(i, "L").value
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
            .Font.Color = RGB(89, 89, 89)
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
            .Font.Color = RGB(89, 89, 89)
            .Font.Size = 8
        End With
            
            colorIndex = (colorIndex + 1) Mod UBound(chartColors)

        End If
    Next i

End Sub
'----------------------------------------------------------------------------------
Function ColNumToLetter(colNum As Integer) As String
    'CreateGraphHelmet�Ɏg�p����֐��B
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
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

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
    ' CreateGraphHelmet�̃T�u�v���V�[�W��
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

    ws.Cells(i, "H").value = maxVal

    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=445, Top:=chartTop, Height:=225) '�^�������p�ɃT�C�Y�ύX
    Dim chart As chart
    Set chart = chartObj.chart

    ConfigureChart chart, ws, i, colStart, colEnd, maxVal

End Sub

Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    'CreateIndividualChart�̃T�u�v���V�[�W��
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.ChartTitle.text = ws.Cells(i, "B").value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75

    SetYAxis chart, ws, i
    SetXAxis chart

End Sub

Sub SetYAxis(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long)
    'ConfigureChart�Ŏg�p����֐�
    '�ی�X��������ɍ��킹�ēV���̎����������łȂ����Ŕ��f����悤�ɂ����B
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    Dim eValue As String
    eValue = ws.Cells(i, "E").value ' E��̒l���擾
    'Debug.Print eValue

    If eValue = "�V��" Then
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1# '1.0����
    Else
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2# '2.0����
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.Color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub


'Sub SetYAxis(ByRef chart As chart, ByVal maxVal As Double)
'    Dim yAxis As Axis
'    Set yAxis = chart.Axes(xlValue, xlPrimary)
'
'    If maxVal <= 4.95 Then
'        yAxis.MaximumScale = 5
'    ElseIf maxVal <= 9.81 Then
'        yAxis.MaximumScale = 10
'    Else
'        yAxis.MaximumScale = Int(maxVal) + 1
'    End If
'
'    yAxis.MinimumScale = -1
'
'    With yAxis.TickLabels
'        .NumberFormatLocal = "0.0""kN"""
'        .Font.color = RGB(89, 89, 89)
'        .Font.Size = 8
'    End With
'End Sub

Sub SetXAxis(ByRef chart As chart)
    'ConfigureChart�Ŏg�p����֐�
    Dim xAxis As Axis
    Set xAxis = chart.Axes(xlCategory, xlPrimary)
    xAxis.TickLabelSpacing = 100
    xAxis.TickMarkSpacing = 25

    With xAxis.TickLabels
        .NumberFormatLocal = "0.00""ms"""
        .Font.Color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub
'----------------------------------------------------------------------------------





Sub InspectHelmetDurationTime()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' "LOG_Helmet" �V�[�g���w�肷��B
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' �ŏI�s���擾����B
    lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).row

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
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.count).End(xlToLeft))
    ' �ő�l���擾����B
    MaxValue = Application.WorksheetFunction.Max(rng)
    ws.Cells(row, "H").value = MaxValue

    ' �ő�l�̈ʒu��������B
    For j = 1 To rng.Columns.count
        If rng(1, j).value = MaxValue Then
            maxValueColumn = j
            rng(1, j).Interior.Color = RGB(250, 150, 0)  ' �F��ݒ肷��B
            Exit For
        End If
    Next j

    If maxValueColumn <> 0 Then
        ws.Cells(row, "I").value = ws.Cells(1, "V").Offset(0, maxValueColumn - 1).value
    End If
End Sub

Sub UpdatePartOfHelmet(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�w�����b�g�̎����ӏ����X�V����
    '�ŏ��ɎQ�Ƃ���̂�B��̒l
    Dim cellValue As String
    cellValue = ws.Cells(row, "B").value
    
    ' �����̒l���擾����B
    Dim existingValue As String
    existingValue = ws.Cells(row, "E").value

    ' �w�����b�g�̕����Ɋ�Â��Ēl��ݒ肷��B�������A"�V��"��"����"�����łɊ܂܂�Ă���ꍇ�͕ύX���Ȃ��B
    ' �����߂ł͍ŏ���E��̒l���`�F�b�N����B
    If InStr(existingValue, "�V��") > 0 Or InStr(existingValue, "����") > 0 Then
        ' �ύX���Ȃ�
    ElseIf InStr(cellValue, "HEL_TOP") > 0 Then
        ws.Cells(row, "E").value = "�V��"
    ElseIf InStr(cellValue, "HEL_ZENGO") > 0 Then
        ws.Cells(row, "E").value = "�O�㓪��"
    ElseIf InStr(cellValue, "HEL_SIDE") > 0 Then
        ws.Cells(row, "E").value = "������"
    End If
End Sub


Sub UpdateRangeForThresholds(ByRef ws As Worksheet, ByVal row As Long, ByVal threshold As Double, ByVal columnToWrite As String)
    'InspectHelmetDurationTime()����4.9�A7.35�͈̔͒l�̐F�t���ƏՌ����Ԃ��L������B
    Dim rng As Range, cell As Range
    Dim startRange As Long, endRange As Long, maxRange As Long
    Dim rangeCollection As New Collection
    Dim timeDifference As Double

    ' �s�͈̔͂��Z�b�g����B
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.count).End(xlToLeft))

    ' 臒l�𒴂���Z���͈̔͂�������B
    For Each cell In rng
        If cell.value >= threshold Then
            If startRange = 0 Then startRange = cell.Column
            endRange = cell.Column
            cell.Interior.Color = IIf(threshold = 4.9, RGB(255, 111, 56), RGB(234, 67, 53))
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
        timeDifference = ws.Cells(1, endRange).value - ws.Cells(1, startRange).value
        ws.Cells(row, columnToWrite).value = timeDifference
    Else
        ws.Cells(row, columnToWrite).value = "-"
    End If
End Sub

Sub FillEmptyCells(ByRef ws As Worksheet, ByVal lastRow As Long)
    'InspectHelmetDurationTime()���̃v���V�[�W��_�󗓂�-������B
    Dim cellRng As Range

    ' B��̍Ō�̍s���擾����B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row

    ' ��̃Z����"-"�Ŗ��߂�B
    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.value = "-"
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
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row
    
    ' �F�̃C���f�b�N�X��������
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�
    
    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' ���݂̃Z���̒l���擾
        valueToFind = ws.Cells(i, "H").value
        
        ' �����l�����Z�������ɐF�t������Ă��Ȃ����`�F�b�N
        If ws.Cells(i, "H").Interior.colorIndex = xlNone Then
            ' ���݂̃Z���̒l�Ɠ����l�����Z����T��
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
    cellValue = ws.Cells(row, "B").value

    ' �w�����b�g�̕����Ɋ�Â��Ēl��ݒ肷��B
    If InStr(cellValue, "TOP") > 0 Then
        ws.Cells(row, "E").value = "�V��"
    ElseIf InStr(cellValue, "MAE") > 0 Then
        ws.Cells(row, "E").value = "�O����"
    ElseIf InStr(cellValue, "USHIRO") > 0 Then
        ws.Cells(row, "E").value = "�㓪��"
    Else
        ws.Cells(row, "E").value = "�O�㓪��"
    End If
End Sub
