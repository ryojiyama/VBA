Attribute VB_Name = "ChartCreate_BaseBall"



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


' �O���t���쐬���郁�C���̃T�u�v���V�[�W��
Sub CreateGraphBaseBall()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("LOG_BaseBall")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim colStart As String
    Dim colEnd As String
    Dim chartSize As Variant
    Dim userInput As String

    'userInput = InputBox("�O���t�̎�ނ���͂��Ă��������i����V���A����O��A�^���V���A�^���O��j")

    colStart = "HI"  ' �J�n���'-0.5'
    chartTop = ws.Rows(lastRow + 1).Top + 10
    chartLeft = 250

    For i = 2 To lastRow
        colEnd = "SC" '�I�����'2.3'
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
    chart.SeriesCollection(1).Format.Line.Weight = 1.5 '0.75

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
    yAxis.MinimumScale = -50
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





