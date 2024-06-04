Attribute VB_Name = "HelmetCreate_Katashiki"
Sub VisualizeSelectedData_HelmetGraph()
    ' ���[�N�V�[�g�̐ݒ�
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' �ŏI�s���擾
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' �O���t�̐F�z��
    Dim chartColors As Variant
    chartColors = Array(RGB(47, 85, 151), RGB(241, 88, 84), RGB(111, 178, 85), _
                        RGB(250, 194, 58), RGB(158, 82, 143), RGB(255, 127, 80), _
                        RGB(250, 159, 137), RGB(72, 61, 139))

    ' �O���t�ݒu�ʒu
    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100

    ' ��̊J�n�ƏI��
    Dim colStart As String
    Dim colEnd As String
    colStart = "HT"   '-1:JA, -2:HT
    colEnd = "SA"

    ' �O���t�I�u�W�F�N�g�̏����ݒ�
    Dim ChartObj As ChartObject
    Dim chart As chart
    Dim colorIndex As Integer
    colorIndex = 0

    ' �f�[�^�s���Ƃ̏���
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, "B").Interior.color = RGB(252, 228, 214) Then ' ����̔w�i�F�̏ꍇ�ɏ���
            ' �V�����O���t�̍쐬�܂��͊����O���t�ւ̒ǉ�
            If Not ChartObj Is Nothing Then
                Dim series As series
                Set series = chart.SeriesCollection.NewSeries
                series.Values = ws.Range(colStart & i & ":" & colEnd & i)
                series.XValues = ws.Range(colStart & "1:" & colEnd & "1")
                series.Format.Line.ForeColor.RGB = chartColors(colorIndex)
                series.name = ws.Cells(i, "D").Value & " - " & ws.Cells(i, "L").Value
            Else
                Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=425, Top:=chartTop, Height:=225)
                Set chart = ChartObj.chart
                
                chart.ChartType = xlLine
                chart.SetSourceData Source:=ws.Range(colStart & i & ":" & colEnd & i)
                chart.SeriesCollection(1).Format.Line.ForeColor.RGB = chartColors(colorIndex)
                chart.SeriesCollection(1).XValues = ws.Range(colStart & "1:" & colEnd & "1")
                chart.SeriesCollection(1).name = ws.Cells(i, "D").Value & " - " & ws.Cells(i, "L").Value
            End If

            ' �O���t�̐ݒ蒲��
            With chart
                .SeriesCollection(1).Format.Line.Weight = 1#

                With .Axes(xlValue, xlPrimary)
                    .MinimumScale = -1 ' Y���̍Œ�l�ݒ�
                    With .TickLabels
                        .NumberFormatLocal = "0.0""kN"""
                        .Font.color = RGB(89, 89, 89)
                        .Font.Size = 8
                    End With
                End With

                With .Axes(xlCategory, xlPrimary)
                    .TickLabelSpacing = 100
                    .TickMarkSpacing = 25
                    With .TickLabels
                        .NumberFormatLocal = "0.00""ms"""
                        .Font.color = RGB(89, 89, 89)
                        .Font.Size = 8
                    End With
                End With
            End With

            ' ���̐F��
            colorIndex = (colorIndex + 1) Mod UBound(chartColors)
        End If
    Next i
End Sub

Sub AddVerticalGridlinesToAllCharts_Test01()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim ChartObj As ChartObject
    Dim yAxis As Axis

    ' �V�[�g��̑S�Ẵ`���[�g�I�u�W�F�N�g�����[�v
    For Each ChartObj In ws.ChartObjects
        ' �c���̖ڐ������ݒ�
        Set yAxis = ChartObj.chart.Axes(xlValue, xlPrimary)
        
        On Error Resume Next ' ���̎����ڐ�������T�|�[�g���Ă��Ȃ��ꍇ�̃G���[�𖳎�
        With yAxis.MajorGridlines
            .Format.Line.Weight = 0.5
            .Format.Line.DashStyle = msoLineDashDot ' �_��
            .Visible = True
        End With
        If Err.number <> 0 Then
            Debug.Print "Error applying gridlines to chart: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0 ' �G���[�n���h�����O���f�t�H���g�ɖ߂�
    Next ChartObj
End Sub

Sub AddVerticalGridlinesToAllCharts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim ChartObj As ChartObject
    Dim yAxis As Axis

    ' �V�[�g��̑S�Ẵ`���[�g�I�u�W�F�N�g�����[�v
    For Each ChartObj In ws.ChartObjects
        ' �`���[�g�^�C�v���ڐ�������T�|�[�g���Ă��邩�ǂ����`�F�b�N
        If ChartSupportsGridlines(ChartObj.chart) Then
            ' �c���̖ڐ������ݒ�
            Set yAxis = ChartObj.chart.Axes(xlValue, xlPrimary)
            If Not yAxis.HasMajorGridlines Then
                yAxis.HasMajorGridlines = True ' MajorGridlines��L���ɂ���
            End If
            With yAxis.MajorGridlines
                .Visible = True
                .Format.Line.Weight = 0.5
                .Format.Line.DashStyle = msoLineDashDot ' �_��

            End With
        Else
            Debug.Print "Chart does not support gridlines: " & ChartObj.name
        End If
    Next ChartObj
End Sub

Function ChartSupportsGridlines(chart As chart) As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = chart.Axes(xlValue, xlPrimary).MajorGridlines
    ChartSupportsGridlines = Err.number = 0
    On Error GoTo 0
End Function

Sub AddYAxisGridlinesToAllCharts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim ChartObj As ChartObject

    ' �V�[�g��̑S�Ẵ`���[�g�I�u�W�F�N�g�����[�v
    For Each ChartObj In ws.ChartObjects
        With ChartObj.chart
            ' Y���i�l���j�����݂��邩�ǂ����m�F
            If .HasAxis(xlValue, xlPrimary) Then
                Dim yAxis As Axis
                Set yAxis = .Axes(xlValue, xlPrimary)
                yAxis.HasMajorGridlines = True ' Y���̎�v�ڐ������L���ɂ���
                yAxis.MajorGridlines.Format.Line.Visible = msoTrue ' �ڐ����������Ԃɐݒ�
            End If
        End With
    Next ChartObj
End Sub

Sub CreateLineChartWithGridlines()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' ��Ƃ��s���V�[�g��ݒ肵�܂��B�K�v�ɉ����ăV�[�g���Ŏw�肵�Ă��������B

    ' �O���t���쐬����͈͂��w��
    Dim chartRange As Range
    Set chartRange = ws.Range("V1:AX2")

    ' �O���t�I�u�W�F�N�g���쐬
    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=100, Width:=600, Top:=50, Height:=400)

    ' �O���t�̐ݒ�
    With ChartObj.chart
        .SetSourceData Source:=chartRange, PlotBy:=xlColumns
        .ChartType = xlLine ' �܂���O���t���w��

        ' X���̖ڐ������ǉ�
        With .Axes(xlCategory, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.Weight = 0.75 ' ���̑�����0.75pt�ɐݒ�
            .MajorGridlines.Format.Line.DashStyle = msoLineSolid ' �����ɐݒ�
        End With

        ' Y���̖ڐ������ǉ�
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.Weight = 0.75 ' ���̑�����0.75pt�ɐݒ�
            .MajorGridlines.Format.Line.DashStyle = msoLineSolid ' �����ɐݒ�
        End With
    End With
End Sub

