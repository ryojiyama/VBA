Attribute VB_Name = "ChartCreate_BicycleHelmet"
' �O���t�̃T�C�Y�����肷��֐�
Function GetChartSize(ByVal helmetType As String) As Variant
    Dim size(1) As Long
    Select Case helmetType
        Case "HEL_TOP", "HEL_ZENGO"
            size(0) = 250  ' Width
            size(1) = 300  ' Height
        Case "HEL_SIDE"
            size(0) = 270  ' Width
            size(1) = 300  ' Height
        Case Else
            size(0) = 350  ' Width
            size(1) = 300  ' Height
    End Select
    GetChartSize = size
End Function
Sub CreateGraphBicycle()
    Call CreateGraphBicycleMain
    Call Bicycle_150G_DurationTime
End Sub


' �O���t���쐬���郁�C���̃T�u�v���V�[�W��
Sub CreateGraphBicycleMain()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim colStart As String
    Dim colEnd As String
    Dim chartSize As Variant
    Dim userInput As String

    'userInput = InputBox("�O���t�̎�ނ���͂��Ă��������i����V���A����O��A�^���V���A�^���O��j")

    colStart = "BO"  ' �J�n���'-2'
    chartTop = ws.Rows(lastRow + 1).Top + 10
    chartLeft = 250

    For i = 2 To lastRow
        colEnd = "ARW" '�I�����'9'
        chartSize = GetChartSize(ws.Cells(i, "B").value)
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd, chartSize
        chartLeft = chartLeft + 10
    Next i

End Sub


' �ʂ̃O���t��ݒ�E�ǉ�����T�u�v���V�[�W��
Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String, ByVal chartSize As Variant)
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))
    ws.Cells(i, "H").value = maxVal
    
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartSize(0), Top:=chartTop, Height:=chartSize(1))
    Dim chart As chart
    Set chart = chartObj.chart
    
    With chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))  ' X���͈̔͂�1�s�ڂ���ݒ�
        .SeriesCollection(1).Values = ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))  ' Y���̃f�[�^�͈͂�ݒ�
        .SeriesCollection(1).Name = "Data Series " & i
    End With
    
    ConfigureChart chart, ws, i, colStart, colEnd, maxVal
End Sub

Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    '���̃v���V�[�W����X����Y���̖ڐ�����ǉ�����B�������Ȃ��Ƃ��܂������Ȃ��B
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.chartTitle.text = ws.Cells(i, "B").value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75  '0.75

    SetYAxis chart, ws, i, maxVal
    SetXAxis chart

    ' Y���ڐ�����ǉ�
    With chart.Axes(xlValue, xlPrimary)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.Weight = 0.25
        .MajorGridlines.Format.Line.DashStyle = msoLineDashDot
    End With

    ' X���ڐ�����ǉ�
    With chart.Axes(xlCategory, xlPrimary)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.Weight = 0.25
        .MajorGridlines.Format.Line.DashStyle = msoLineDashDot
    End With
End Sub

Sub SetYAxis(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal maxVal As Double)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    Dim eValue As String
    eValue = ws.Cells(i, "E").value ' E��̒l���擾

    ' Y���̍ő�l�� maxVal �̒l��10�̈ʂ�50�P�ʂŌJ��グ
    Dim roundedMax As Double
    roundedMax = WorksheetFunction.RoundUp(maxVal / 50, 0) * 50

    ' Y���̐ݒ���s��
    yAxis.MaximumScale = roundedMax
    yAxis.MajorUnit = WorksheetFunction.RoundUp((roundedMax / 5), 0) ' �ڐ���P�ʂ��K�؂ɐݒ�
    yAxis.MinimumScale = 0
    yAxis.MajorUnit = 50
    

    With yAxis.TickLabels
        .NumberFormatLocal = "0""G""" ' ���x���̐��l�`����ݒ�
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With
End Sub



Sub SetXAxis(ByRef chart As chart)
    Dim xAxis As Axis
    Set xAxis = chart.Axes(xlCategory, xlPrimary)

    xAxis.TickLabelSpacing = 100
    xAxis.TickMarkSpacing = 50

    With xAxis.TickLabels
        .NumberFormatLocal = "0.0""ms"""
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With

End Sub

Sub Bicycle_150G_DurationTime()
    '���]�ԖX�����̃f�[�^���������郁�C���̃T�u���[�`��
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = GetLastRow(ws, "B")

    '�e�s�̍ő�l��F�t�����A�ő�l�̎��Ԃ��L�^���܂�
    ColorAndRecordMaxVal ws, lastRow, 150

    '150G�ȏ���L�^�����͈͂�F�t�����A���͈̔͂̎��ԍ����L�^���܂�
    ColorAndRecordTimeDifference ws, lastRow, 150

    '��̃Z����"-"�Ŗ��߂܂�
    FillEmptyCells ws, GetLastRow(ws, "B")
End Sub

'Bicycle150GDuration_����̃J�����̍ŏI�s���擾����֐��ł��B
Function GetLastRow(ws As Worksheet, column As String) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
End Function

'Bicycle150GDuration_�e�s�̍ő�l�̃Z����F�t�����A�ő�l�̎��Ԃ��L�^����T�u���[�`���ł��B
Sub ColorAndRecordMaxVal(ws As Worksheet, lastRow As Long, threshold As Double)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range

    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "AA"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))

        Dim MaxValue As Double
        MaxValue = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "H").value = MaxValue

        For Each cell In rng
            If cell.value = MaxValue Then
                cell.Interior.Color = RGB(255, 111, 56)
                ws.Cells(i, "I").value = ws.Cells(1, cell.column).value ' �Ή����鎞�Ԃ�I��ɋL�^
                Exit For ' �ŏ��̍ő�l�����������烋�[�v�𔲂���
            End If
        Next cell
    Next i
End Sub

'Bicycle150GDuration_150G�ȏ���L�^�����͈͂�F�t�����A���͈̔͂̎��ԍ����L�^����T�u���[�`���ł��B
Sub ColorAndRecordTimeDifference(ws As Worksheet, lastRow As Long, threshold As Double)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range
    Dim startRange150 As Long
    Dim endRange150 As Long
    Dim maxRange150 As Long
    Dim maxStart150 As Long
    Dim maxEnd150 As Long
    Dim rangeCollection150 As Collection

    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "AA"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        startRange150 = 0
        endRange150 = 0
        maxRange150 = 0
        Set rangeCollection150 = New Collection

        For Each cell In rng
            If cell.value >= threshold Then
                If startRange150 = 0 Then startRange150 = cell.column
                endRange150 = cell.column
                cell.Interior.Color = RGB(0, 138, 211)
            Else
                If startRange150 > 0 And endRange150 > 0 Then
                    rangeCollection150.Add Array(startRange150, endRange150)
                    If (endRange150 - startRange150 + 1) > maxRange150 Then
                        maxRange150 = endRange150 - startRange150 + 1
                        maxStart150 = startRange150
                        maxEnd150 = endRange150
                    End If
                    startRange150 = 0
                    endRange150 = 0
                End If
            End If
        Next cell

        If startRange150 > 0 And endRange150 > 0 Then
            rangeCollection150.Add Array(startRange150, endRange150)
            If (endRange150 - startRange150 + 1) > maxRange150 Then
                maxRange150 = endRange150 - startRange150 + 1
                maxStart150 = startRange150
                maxEnd150 = endRange150
            End If
        End If

        If maxRange150 > 0 Then
            Dim timeDifference150 As Double
            timeDifference150 = ws.Cells(1, maxEnd150).value - ws.Cells(1, maxStart150).value
            ws.Cells(i, "K").value = timeDifference150
        Else
            ws.Cells(i, "K").value = "-"
        End If
    Next i
End Sub

'Bicycle150GDuration_��̃Z����"-"�Ŗ��߂�T�u���[�`���ł��B
Sub FillEmptyCells(ws As Worksheet, lastRow As Long)
    Dim cellRng As Range

    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.value = "-"
        End If
    Next cellRng
End Sub
