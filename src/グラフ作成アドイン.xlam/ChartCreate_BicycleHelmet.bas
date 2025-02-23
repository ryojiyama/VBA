Attribute VB_Name = "ChartCreate_BicycleHelmet"
' グラフのサイズを決定する関数
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


' グラフを作成するメインのサブプロシージャ
Sub CreateGraphBicycle()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim colStart As String
    Dim colEnd As String
    Dim chartSize As Variant
    Dim userInput As String

    'userInput = InputBox("グラフの種類を入力してください（定期天頂、定期前後、型式天頂、型式前後）")

    colStart = "BO"  ' 開始列は'-2'
    chartTop = ws.Rows(lastRow + 1).Top + 10
    chartLeft = 250

    For i = 2 To lastRow
        colEnd = "ARW" '終了列は'9'
        chartSize = GetChartSize(ws.Cells(i, "B").value)
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd, chartSize
        chartLeft = chartLeft + 10
    Next i

End Sub


' 個別のグラフを設定・追加するサブプロシージャ
Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String, ByVal chartSize As Variant)
    Dim maxVal As Double
    Dim recordID As String
    
    ' チャートIDの生成
    recordID = CreateChartID(ws.Cells(i, "B"))
    Debug.Print "recordID: " & recordID ' イミディエイトウィンドウにIDを出力
    
    ' 最大値を取得
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))
    ws.Cells(i, "J").value = maxVal '最大値をJ列に記載
    
    ' チャートオブジェクトを追加
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartSize(0), Top:=chartTop, Height:=chartSize(1))
    
    ' ★★★ チャートオブジェクトの名前にIDを設定 ★★★
    chartObj.Name = recordID
    
    Dim chart As chart
    Set chart = chartObj.chart
    
    ' チャートに設定を適用
    With chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
        .SeriesCollection(1).Values = ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
        .SeriesCollection(1).Name = "Data Series " & i
        .HasTitle = True
        .chartTitle.Text = "ID: " & recordID ' チャートタイトルにIDを設定（表示用）
    End With
    
    ' チャートの詳細設定
    ConfigureChart chart, ws, i, colStart, colEnd, maxVal
End Sub

Function CreateChartID(cell As Range) As String

    ' B列の値が空の場合は"00000"を返す
    If IsEmpty(cell) Or cell.value = "" Then
        CreateChartID = "00000"
    Else
        CreateChartID = cell.value
    End If

End Function

Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    'このプロシージャでX軸とY軸の目盛線を追加する。そうしないとうまくいかない。
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.chartTitle.Text = ws.Cells(i, "B").value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75  '0.75

    SetYAxis chart, ws, i, maxVal
    SetXAxis chart

    ' Y軸目盛線を追加
    With chart.Axes(xlValue, xlPrimary)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.Weight = 0.25
        .MajorGridlines.Format.Line.DashStyle = msoLineDashDot
    End With

    ' X軸目盛線を追加
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
    eValue = ws.Cells(i, "E").value ' E列の値を取得

    ' Y軸の最大値を maxVal の値を10の位で50単位で繰り上げ
    Dim roundedMax As Double
    roundedMax = WorksheetFunction.RoundUp(maxVal / 50, 0) * 50

    ' Y軸の設定を行う
    yAxis.MaximumScale = roundedMax
    yAxis.MajorUnit = WorksheetFunction.RoundUp((roundedMax / 5), 0) ' 目盛り単位も適切に設定
    yAxis.MinimumScale = 0
    yAxis.MajorUnit = 50
    

    With yAxis.TickLabels
        .NumberFormatLocal = "0""G""" ' ラベルの数値形式を設定
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

'自転車帽試験のデータを処理するメインのサブルーチン
Sub Bicycle_150G_DurationTime()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveWorkbook.Sheets("LOG_Bicycle")
    lastRow = GetLastRow(ws, "B")

    '各行の最大値を色付けし、最大値の時間を記録します
    ColorAndRecordMaxVal ws, lastRow, 150

    '150G以上を記録した範囲を色付けし、その範囲の時間差を記録します
    ColorAndRecordTimeDifference ws, lastRow, 150

    '空のセルを"-"で埋めます
    FillEmptyCells ws, GetLastRow(ws, "B")
    
    ' 確認ダイアログを表示
    If MsgBox("衝撃値を調整しますか？", vbYesNo + vbQuestion) = vbYes Then
        Call AdjustImpactValuesWithCustomFormatForAllLOGSheets
    End If
End Sub


Function GetLastRow(ws As Worksheet, column As String) As Long
'Bicycle150GDuration_特定のカラムの最終行を取得する関数
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
End Function


Sub ColorAndRecordMaxVal(ws As Worksheet, lastRow As Long, threshold As Double)
'Bicycle150GDuration_各行の最大値のセルを色付けし、最大値の時間を記録するサブルーチン
    Dim rng As Range
    Dim i As Long
    Dim cell As Range

    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "AA"), ws.Cells(i, ws.columns.Count).End(xlToLeft))

        Dim MaxValue As Double
        MaxValue = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "J").value = MaxValue '最大値をJ列に記録

        For Each cell In rng
            If cell.value = MaxValue Then
                cell.Interior.Color = RGB(255, 111, 56)
                ws.Cells(i, "K").value = ws.Cells(1, cell.column).value ' 最大値の時間をI列に記録
                Exit For ' 最初の最大値が見つかったらループを抜ける
            End If
        Next cell
    Next i
End Sub


Sub ColorAndRecordTimeDifference(ws As Worksheet, lastRow As Long, threshold As Double)
'Bicycle150GDuration_150G以上を記録した範囲を色付けし、その範囲の時間差を記録するサブルーチン
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
        Set rng = ws.Range(ws.Cells(i, "AA"), ws.Cells(i, ws.columns.Count).End(xlToLeft))
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
            ws.Cells(i, "L").value = timeDifference150 '150Gの継続時間をL列に記録
        Else
            ws.Cells(i, "L").value = "-"
        End If
    Next i
End Sub


Sub FillEmptyCells(ws As Worksheet, lastRow As Long)
'Bicycle150GDuration_空のセルを"-"で埋めるサブルーチン
    Dim cellRng As Range

    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.value = "−"
        End If
    Next cellRng
End Sub
