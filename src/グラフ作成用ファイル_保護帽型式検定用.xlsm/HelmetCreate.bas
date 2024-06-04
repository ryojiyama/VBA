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
    colStart = ColNumToLetter(16 + 52) 'ExcelのラベルはBP
    colEnd = ColNumToLetter(16 + 850) 'ExcelのラベルはAGH

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
                Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=425, Top:=chartTop, Height:=225)
                Set chart = ChartObj.chart
                
                chart.ChartType = xlLine
                chart.SetSourceData Source:=ws.Range(colStart & i & ":" & colEnd & i)
                chart.SeriesCollection(1).Format.Line.ForeColor.RGB = chartColors(colorIndex)
                chart.SeriesCollection(1).XValues = ws.Range(colStart & "1:" & colEnd & "1")
                chart.SeriesCollection(1).name = ws.Cells(i, "D").Value & " - " & ws.Cells(i, "L").Value
            End If
            
        ' 線の太さを設定
        chart.SeriesCollection(1).Format.Line.Weight = 1#

        ' Y軸の設定
        Dim yAxis As Axis
        Set yAxis = chart.Axes(xlValue, xlPrimary)

        yAxis.MinimumScale = -1 ' Y軸の最低値を-1に設定します。

        ' Y軸の TickLabels を設定
        With yAxis.TickLabels
            .NumberFormatLocal = "0.0""kN"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With

        ' X軸の設定
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 100
        xAxis.TickMarkSpacing = 25


        ' X軸の TickLabels を設定
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
Sub CreateGraphHelmet()
    ' 定数の代わりに列ラベルを直接指定
    Dim colStart As String
    Dim colEnd As String
    colStart = "HT"
    colEnd = "VA"   'Top:ZA, Front/Back:SA, Side SA

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' 最終行と最終列を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' グラフの位置を設定
    Dim chartLeft As Long
    Dim chartTop As Long
    chartTop = ws.Rows(lastRow).Height - 20
    chartLeft = 250

    ' 各行に対して個別のグラフを作成
    For i = 2 To lastRow
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd
        chartLeft = chartLeft + 10 ' 次のグラフの左位置を更新
    Next i

End Sub


Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String)
    ' CreateGraphHelmetのサブプロシージャ
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

    ws.Cells(i, "H").Value = maxVal

    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=300, Top:=chartTop, Height:=300)
    'Top:Width:=250,Height:=300, Front/Back:Width:=250,Height:=300, Side:Width:=270,Height:=300
    Dim chart As chart
    Set chart = ChartObj.chart

    ConfigureChart chart, ws, i, colStart, colEnd, maxVal

End Sub


Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    'このプロシージャでX軸とY軸の目盛線を追加する。そうしないとうまくいかない。
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.ChartTitle.text = ws.Cells(i, "B").Value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75

    SetYAxis chart, ws, i
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



Sub SetYAxis(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    Dim eValue As String
    eValue = ws.Cells(i, "E").Value ' E列の値を取得

'    If eValue = "天頂" Then
'        yAxis.MaximumScale = 5
'        yAxis.MajorUnit = 1# '1.0刻み
'    Else
'        yAxis.MaximumScale = 10
'        yAxis.MajorUnit = 2# '2.0刻み
'    End If
    yAxis.MaximumScale = 8
    yAxis.MajorUnit = 1# '1.0刻み
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
    xAxis.TickMarkSpacing = 100

    With xAxis.TickLabels
        .NumberFormatLocal = "0.0""ms"""
        .Font.color = RGB(89, 89, 89)
        .Font.Size = 8
    End With

End Sub

'----------------------------------------------------------------------------------

Function ColNumToLetter_Old(colNum As Integer) As String
    ' CreateGraphHelmetにて使用する関数
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



Sub InspectHelmetDurationTime()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' "LOG_Helmet" シートを指定する。
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' 最終行を取得する。
    lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).row

    ' 各行を処理する。
    For i = 2 To lastRow
        UpdateMaxValueInRow ws, i             ' 行内の最大値を更新する。
        UpdatePartOfHelmet ws, i              ' ヘルメットの部分を更新する。
        UpdateRangeForThresholds ws, i, 4.9, "J"  ' 閾値の範囲を更新する。
        UpdateRangeForThresholds ws, i, 7.35, "K"
    Next i

    ' 空のセルを埋める。
    FillEmptyCells ws, lastRow
End Sub

Sub UpdateMaxValueInRow(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()内のプロシージャ_行内の最大値を更新する
    Dim rng As Range
    Dim MaxValue As Double
    Dim maxValueColumn As Long

    ' 行内の範囲をセットする。
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))
    ' 最大値を取得する。
    MaxValue = Application.WorksheetFunction.Max(rng)
    ws.Cells(row, "H").Value = MaxValue

    ' 最大値の位置を見つける。
    For j = 1 To rng.Columns.Count
        If rng(1, j).Value = MaxValue Then
            maxValueColumn = j
            rng(1, j).Interior.color = RGB(250, 150, 0)  ' 色を設定する。
            Exit For
        End If
    Next j

    If maxValueColumn <> 0 Then
        ws.Cells(row, "I").Value = ws.Cells(1, "V").Offset(0, maxValueColumn - 1).Value
    End If
End Sub

Sub UpdatePartOfHelmet(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()内のプロシージャ_ヘルメットの試験箇所を更新する
    '最初に参照するのはB列の値
    Dim cellValue As String
    cellValue = ws.Cells(row, "B").Value
    
    ' 既存の値を取得する。
    Dim existingValue As String
    existingValue = ws.Cells(row, "E").Value

    ' ヘルメットの部分に基づいて値を設定する。ただし、"天頂"や"頭部"がすでに含まれている場合は変更しない。
    ' 条件節では最初にE列の値をチェックする。
    If InStr(existingValue, "天頂") > 0 Or InStr(existingValue, "頭部") > 0 Then
        ' 変更しない
    ElseIf InStr(cellValue, "HEL_TOP") > 0 Then
        ws.Cells(row, "E").Value = "天頂"
    ElseIf InStr(cellValue, "HEL_ZENGO") > 0 Then
        ws.Cells(row, "E").Value = "前後頭部"
    ElseIf InStr(cellValue, "HEL_SIDE") > 0 Then
        ws.Cells(row, "E").Value = "側頭部"
    End If
End Sub


Sub UpdateRangeForThresholds(ByRef ws As Worksheet, ByVal row As Long, ByVal threshold As Double, ByVal columnToWrite As String)
    'InspectHelmetDurationTime()から4.9、7.35の範囲値の色付けと衝撃時間を記入する。
    Dim rng As Range, cell As Range
    Dim startRange As Long, endRange As Long, maxRange As Long
    Dim rangeCollection As New Collection
    Dim timeDifference As Double

    ' 行の範囲をセットする。
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))

    ' 閾値を超えるセルの範囲を見つける。
    For Each cell In rng
        If cell.Value >= threshold Then
            If startRange = 0 Then startRange = cell.Column
            endRange = cell.Column
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

    For Each Item In rangeCollection
        If (Item(1) - Item(0) + 1) > maxRange Then
            maxRange = Item(1) - Item(0) + 1
            startRange = Item(0)
            endRange = Item(1)
        End If
    Next Item

    If startRange > 0 And endRange > 0 Then
        timeDifference = ws.Cells(1, endRange).Value - ws.Cells(1, startRange).Value
        ws.Cells(row, columnToWrite).Value = timeDifference
    Else
        ws.Cells(row, columnToWrite).Value = "-"
    End If
End Sub

Sub FillEmptyCells(ByRef ws As Worksheet, ByVal lastRow As Long)
    'InspectHelmetDurationTime()内のプロシージャ_空欄に-を入れる。
    Dim cellRng As Range

    ' B列の最後の行を取得する。
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' 空のセルを"-"で埋める。
    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng
End Sub

Sub HighlightDuplicateValues()
    ' 衝撃値が被る場合があるのでそれをチェックするプロシージャ
    Dim sheetName As String
    sheetName = "LOG_Helmet"
    
    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    
    ' シートオブジェクトを設定
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row
    
    ' 色のインデックスを初期化
    colorIndex = 3 ' Excelの色インデックスは3から始まる
    
    ' H列の2行目から最終行までループ
    For i = 2 To lastRow
        ' 現在のセルの値を取得
        valueToFind = ws.Cells(i, "H").Value
        
        ' 同じ値を持つセルが既に色付けされていないかチェック
        If ws.Cells(i, "H").Interior.colorIndex = xlNone Then
            ' 現在のセルの値と同じ値を持つセルを探索
            For j = i + 1 To lastRow
                If ws.Cells(j, "H").Value = valueToFind And ws.Cells(j, "H").Interior.colorIndex = xlNone Then
                    ' 同じ値を持つセルを見つけた場合、色を塗る
                    ws.Cells(i, "H").Interior.colorIndex = colorIndex
                    ws.Cells(j, "H").Interior.colorIndex = colorIndex
                End If
            Next j
            
            ' 色インデックスを更新して次の色に変更
            colorIndex = colorIndex + 1
            ' Excelの色インデックスの最大値を超えないようにチェック
            If colorIndex > 56 Then colorIndex = 3 ' 色インデックスをリセット
        End If
    Next i
End Sub



Sub DeleteAllChartsInLOGHelmetSheet()
    'LOG_Helmet内のグラフをすべて削除する。
    Dim chart As ChartObject
    Dim ws As Worksheet
    
    ' LOG_Helmet シートを指定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' 指定シート内の全グラフを削除
    For Each chart In ws.ChartObjects
        chart.Delete
    Next chart
End Sub
Sub MoveChartToNewPosition()
    Dim chart As ChartObject
    Dim ws As Worksheet
    Dim newX As Long
    Dim newY As Long

    ' 移動したいグラフがあるシートを設定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' 新しい位置のX座標とY座標を設定
    newX = 100  ' 例：左から100ポイント
    newY = 500  ' 例：上から100ポイント

    ' シート内の全グラフを移動
    For Each chart In ws.ChartObjects
        chart.Left = newX
        chart.Top = newY
        
        ' 各グラフを少しずつずらして表示
        newX = newX + 15  ' 次のグラフの位置を調整
        'newY = newY + 10  ' Y座標も少し下げる
    Next chart
End Sub







Sub UpdatePartOfHelmet_231013SyuuseiMae(ByRef ws As Worksheet, ByVal row As Long)
    'InspectHelmetDurationTime()内のプロシージャ_ヘルメットの試験箇所を更新する
    Dim cellValue As String
    cellValue = ws.Cells(row, "B").Value

    ' ヘルメットの部分に基づいて値を設定する。
    If InStr(cellValue, "TOP") > 0 Then
        ws.Cells(row, "E").Value = "天頂"
    ElseIf InStr(cellValue, "MAE") > 0 Then
        ws.Cells(row, "E").Value = "前頭部"
    ElseIf InStr(cellValue, "USHIRO") > 0 Then
        ws.Cells(row, "E").Value = "後頭部"
    Else
        ws.Cells(row, "E").Value = "前後頭部"
    End If
End Sub
