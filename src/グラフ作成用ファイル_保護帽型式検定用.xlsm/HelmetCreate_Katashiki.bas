Attribute VB_Name = "HelmetCreate_Katashiki"
Sub VisualizeSelectedData_HelmetGraph()
    ' ワークシートの設定
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' 最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' グラフの色配列
    Dim chartColors As Variant
    chartColors = Array(RGB(47, 85, 151), RGB(241, 88, 84), RGB(111, 178, 85), _
                        RGB(250, 194, 58), RGB(158, 82, 143), RGB(255, 127, 80), _
                        RGB(250, 159, 137), RGB(72, 61, 139))

    ' グラフ設置位置
    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100

    ' 列の開始と終了
    Dim colStart As String
    Dim colEnd As String
    colStart = "HT"   '-1:JA, -2:HT
    colEnd = "SA"

    ' グラフオブジェクトの初期設定
    Dim ChartObj As ChartObject
    Dim chart As chart
    Dim colorIndex As Integer
    colorIndex = 0

    ' データ行ごとの処理
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, "B").Interior.color = RGB(252, 228, 214) Then ' 特定の背景色の場合に処理
            ' 新しいグラフの作成または既存グラフへの追加
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

            ' グラフの設定調整
            With chart
                .SeriesCollection(1).Format.Line.Weight = 1#

                With .Axes(xlValue, xlPrimary)
                    .MinimumScale = -1 ' Y軸の最低値設定
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

            ' 次の色へ
            colorIndex = (colorIndex + 1) Mod UBound(chartColors)
        End If
    Next i
End Sub

Sub AddVerticalGridlinesToAllCharts_Test01()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim ChartObj As ChartObject
    Dim yAxis As Axis

    ' シート上の全てのチャートオブジェクトをループ
    For Each ChartObj In ws.ChartObjects
        ' 縦軸の目盛り線を設定
        Set yAxis = ChartObj.chart.Axes(xlValue, xlPrimary)
        
        On Error Resume Next ' この軸が目盛り線をサポートしていない場合のエラーを無視
        With yAxis.MajorGridlines
            .Format.Line.Weight = 0.5
            .Format.Line.DashStyle = msoLineDashDot ' 点線
            .Visible = True
        End With
        If Err.number <> 0 Then
            Debug.Print "Error applying gridlines to chart: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0 ' エラーハンドリングをデフォルトに戻す
    Next ChartObj
End Sub

Sub AddVerticalGridlinesToAllCharts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    Dim ChartObj As ChartObject
    Dim yAxis As Axis

    ' シート上の全てのチャートオブジェクトをループ
    For Each ChartObj In ws.ChartObjects
        ' チャートタイプが目盛り線をサポートしているかどうかチェック
        If ChartSupportsGridlines(ChartObj.chart) Then
            ' 縦軸の目盛り線を設定
            Set yAxis = ChartObj.chart.Axes(xlValue, xlPrimary)
            If Not yAxis.HasMajorGridlines Then
                yAxis.HasMajorGridlines = True ' MajorGridlinesを有効にする
            End If
            With yAxis.MajorGridlines
                .Visible = True
                .Format.Line.Weight = 0.5
                .Format.Line.DashStyle = msoLineDashDot ' 点線

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

    ' シート上の全てのチャートオブジェクトをループ
    For Each ChartObj In ws.ChartObjects
        With ChartObj.chart
            ' Y軸（値軸）が存在するかどうか確認
            If .HasAxis(xlValue, xlPrimary) Then
                Dim yAxis As Axis
                Set yAxis = .Axes(xlValue, xlPrimary)
                yAxis.HasMajorGridlines = True ' Y軸の主要目盛り線を有効にする
                yAxis.MajorGridlines.Format.Line.Visible = msoTrue ' 目盛り線を可視状態に設定
            End If
        End With
    Next ChartObj
End Sub

Sub CreateLineChartWithGridlines()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 作業を行うシートを設定します。必要に応じてシート名で指定してください。

    ' グラフを作成する範囲を指定
    Dim chartRange As Range
    Set chartRange = ws.Range("V1:AX2")

    ' グラフオブジェクトを作成
    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=100, Width:=600, Top:=50, Height:=400)

    ' グラフの設定
    With ChartObj.chart
        .SetSourceData Source:=chartRange, PlotBy:=xlColumns
        .ChartType = xlLine ' 折れ線グラフを指定

        ' X軸の目盛り線を追加
        With .Axes(xlCategory, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.Weight = 0.75 ' 線の太さを0.75ptに設定
            .MajorGridlines.Format.Line.DashStyle = msoLineSolid ' 実線に設定
        End With

        ' Y軸の目盛り線を追加
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.Weight = 0.75 ' 線の太さを0.75ptに設定
            .MajorGridlines.Format.Line.DashStyle = msoLineSolid ' 実線に設定
        End With
    End With
End Sub

