Attribute VB_Name = "FallArrestCreate"
Public Sub Create_FallArrestGraph()
    Call CreateGraphFallArrest
    Call FallArrest_2kN_DurationTime
End Sub


Function ColNumToLetter(colNum As Integer) As String
    ' CreateGraphFallArrestにて使用する関数
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

    ' ワークシートを宣言
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' 最終行と最終列を検索
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    Dim i As Long
    Dim maxVal As Double
    Dim colStart As String
    Dim colEnd As String

    ' P列(16番目) + 100列から始まる列
    colStart = ColNumToLetter(16 + 100)

    ' P列(16番目) + 800列から始まる列
    colEnd = ColNumToLetter(16 + 1200)

    ' 初期のチャートの位置
    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100

    ' 2行目から最終行までループ
    For i = 2 To lastRow
        ' 最大値を求める
        maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

        ' K列に最大値を表示
        ws.Cells(i, "G").Value = maxVal

        ' チャートを作成
        Dim ChartObj As ChartObject
        Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
        Dim chart As chart
        Set chart = ChartObj.chart

        ' 折れ線グラフを設定
        chart.ChartType = xlLine

        ' グラフのデータ範囲を設定
        chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))

        ' グラフのタイトルを設定
        chart.HasTitle = True
        chart.ChartTitle.Text = ws.Cells(i, "B").Value

        ' 系列の表示をオフ
        chart.SetElement msoElementLegendNone

        ' 線の太さを設定
        chart.SeriesCollection(1).Format.Line.Weight = 0.75

        ' Y軸の設定
        Dim yAxis As Axis
        Set yAxis = chart.Axes(xlValue, xlPrimary)

        ' Y軸の最大値を設定
        If maxVal <= 4.95 Then
            yAxis.MaximumScale = 5
        ElseIf maxVal > 4.95 And maxVal <= 9.81 Then
            yAxis.MaximumScale = 10
        Else
            yAxis.MaximumScale = Int(maxVal) + 1
        End If

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
        xAxis.TickLabelSpacing = 200
        xAxis.TickMarkSpacing = 50


        ' X軸の TickLabels を設定
        With xAxis.TickLabels
            .NumberFormatLocal = "0""ms"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With
        
        ' チャートの位置を次に更新
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

    ' "LOG"シートを指定します。
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' 最終行を取得します。
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' 全行について処理を行います。
    For i = 2 To lastRow
        ' 現在の行の範囲を設定します。
        Set rng = ws.Range(ws.Cells(i, "P"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' 範囲の初期化を行います。
        startRange = 0
        endRange = 0
        maxRange = 0
        sumRange = 0
        countRange = 0
        ' コレクションの初期化を行います。
        Set rangeCollection = New Collection
        ' 各行で最大値を見つけ、"F"列に記入します。
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' 現在の行内の各セルをチェックします。
        For Each cell In rng
            cellVal = cell.Value
            ' 値が2.2以上なら範囲を更新し、色を継続します。
            If cellVal >= 2.2 Then
                If startRange = 0 Then
                    startRange = cell.column
                    currentColor = RGB(Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0), Int((255 - 0 + 1) * Rnd + 0))
                End If
                endRange = cell.column
                cell.Interior.color = currentColor
                sumRange = sumRange + cellVal
                countRange = countRange + 1
            ' それ以外なら範囲をコレクションに保存し、範囲をリセットします。
            Else
                If startRange > 0 And endRange > 0 Then
                    rangeCollection.Add Array(startRange, endRange)
                    startRange = 0
                    endRange = 0
                End If
            End If
        Next cell

        ' 残った範囲をコレクションに追加します。
        If startRange > 0 And endRange > 0 Then rangeCollection.Add Array(startRange, endRange)

        ' 2.2以上の値を取得し、その合計の平均値を(k,i)に表示します。
        ' 2.2以下の場合は、(k,i)に0を表示します。8.0kNを超える場合はその最大値を(k,i)に表示します。
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If
    Next i
    
    ' 空のセルは全て"-"で埋める
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

    ' "LOG"シートを指定します。
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' 最終行を取得します。
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' 全行について処理を行います。
    For i = 2 To lastRow
        ' 現在の行の範囲を設定します。
        Set rng = ws.Range(ws.Cells(i, "Q"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' 範囲の初期化を行います。
        startPoint = 0
        endPoint = 0
        maxRange = 0
        sumRange = 0
        countRange = 0

        ' 各行で最大値を見つけ、"F"列に記入します。
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' 現在の行内の各セルをチェックします。
        For Each cell In rng
            cellVal = cell.Value
            ' 値が2.2以上なら範囲を更新します。
            If cellVal >= 2.2 Then
                If startPoint = 0 Then
                    startPoint = cell.column
                End If
                endPoint = cell.column
                sumRange = sumRange + cellVal
                countRange = countRange + 1
            End If
        Next cell

        ' startPointとendPointの間の列の該当行の値を合計し、その結果をL列に入力します。
        If startPoint > 0 And endPoint > 0 Then
            Dim sumAbs As Double: sumAbs = 0
            Dim j As Long
            
            For j = startPoint To endPoint
                sumAbs = sumAbs + Abs(ws.Cells(i, j).Value)
            Next j
            
            ws.Cells(i, "L").Value = sumAbs / (endPoint - startPoint + 1)
        End If

        ' 2.2以上の値を取得し、その合計の平均値を(k,i)に表示します。
        ' 2.2以下の場合は、(k,i)に0を表示します。8.0kNを超える場合はその最大値を(k,i)に表示します。
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If

    Next i
    
    ' 空のセルは全て"-"で埋める
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

    ' "LOG"シートを指定します。
    Set ws = ThisWorkbook.Sheets("LOG_FallArrest")
    ' 最終行を取得します。
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).row

    ' 全行について処理を行います。
    For i = 2 To lastRow
        ' 現在の行の範囲を設定します。
        Set rng = ws.Range(ws.Cells(i, "Q"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        ' 範囲の初期化を行います。
        startPoint = 0
        endPoint = 0
        maxRange = 0
        sumRange = 0
        countRange = 0

        ' 各行で最大値を見つけ、"F"列に記入します。
        maxRange = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "F").Value = maxRange

        ' 現在の行内の各セルをチェックします。
        For Each cell In rng
            cellVal = cell.Value
            ' 値が2.2以上なら範囲を更新し、色を継続します。
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

        ' startPointとendPointの間の列の該当行の値を合計し、その結果をL列に入力します。
        If startPoint > 0 And endPoint > 0 Then
            Dim sumAbs As Double: sumAbs = 0
            Dim j As Long
            
            For j = startPoint To endPoint
                sumAbs = sumAbs + Abs(ws.Cells(i, j).Value)
            Next j
            
            ws.Cells(i, "L").Value = sumAbs / (endPoint - startPoint + 1)
        End If

        ' 2.2以上の値を取得し、その合計の平均値を(k,i)に表示します。
        ' 2.2以下の場合は、(k,i)に0を表示します。8.0kNを超える場合はその最大値を(k,i)に表示します。
        If maxRange <= 2.2 Then
            ws.Cells(i, "K").Value = 0
        ElseIf maxRange > 8# Then
            ws.Cells(i, "K").Value = maxRange
        Else
            ws.Cells(i, "K").Value = sumRange / countRange
        End If

    Next i
    
    ' 空のセルは全て"-"で埋める
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    Dim cellRng As Range
    For Each cellRng In ws.Range("F2:O" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng

End Sub




' アクティブシート内のグラフを削除
Sub DeleteAllChartsInActiveSheet()
    Dim chart As ChartObject
    
    For Each chart In ActiveSheet.ChartObjects
        chart.Delete
    Next chart
End Sub

