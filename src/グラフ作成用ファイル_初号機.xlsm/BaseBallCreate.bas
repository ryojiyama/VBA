Attribute VB_Name = "BaseBallCreate"
Public Sub Create_BaseBallGraph()
    Call CreateGraphBaseBall
    Call BaseBall_5kN7kN_DurationTime
End Sub


Function ColNumToLetter(colNum As Integer) As String
    ' CreateGraphBaseBallにて使用する関数
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



Sub CreateGraphBaseBall()

    ' ワークシートを宣言
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_BaseBall")
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
        If maxVal <= 90 Then
            yAxis.MaximumScale = 100
            yAxis.MinimumScale = -10 ' Y軸の最低値を-10に設定します。
        ElseIf maxVal > 91 And maxVal <= 299 Then
            yAxis.MaximumScale = 300
            yAxis.MinimumScale = -100 ' Y軸の最低値を-100に設定します。
        Else
            yAxis.MaximumScale = Int(maxVal) + 1
            yAxis.MinimumScale = -100 ' Y軸の最低値を-100に設定します。
        End If

        ' Y軸の TickLabels を設定
        With yAxis.TickLabels
            .NumberFormatLocal = "0""G"""
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


Sub BaseBall_5kN7kN_DurationTime()
    '野球帽の試験データの処理を行うメインのサブルーチン
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("LOG_BaseBall")
    lastRow = GetLastRow(ws, "P")
    
    '各行の最大値を色付けし、最大値の時間を記録します
    ColorMaxValCells ws, lastRow
    
    '空のセルを"-"で埋めます
    FillEmptyCells ws, GetLastRow(ws, "B")
End Sub

'特定のカラムの最終行を取得する関数です。BaseBall_5kN7kN_DurationTime()で最終行を取得するために使います。
Function GetLastRow(ws As Worksheet, column As String) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
End Function

'各行の最大値のセルを色付けし、最大値の時間を記録するサブルーチンです。BaseBall_5kN7kN_DurationTime()の一部として動作します。
Sub ColorMaxValCells(ws As Worksheet, lastRow As Long)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range
    
    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "P"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        Dim MaxValue As Double
        MaxValue = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "H").Value = MaxValue
        
        For Each cell In rng
            If cell.Value = MaxValue Then
                cell.Interior.color = RGB(250, 150, 0)
                ws.Cells(i, "I").Value = ws.Cells(1, cell.column).Value ' 対応する時間をI列に記録
                Exit For ' 最初の最大値が見つかったらループを抜ける
            End If
        Next cell
    Next i
End Sub

'すべての空セルを"-"で埋めるサブルーチンです。BaseBall_5kN7kN_DurationTime()の一部として動作します。
Sub FillEmptyCells(ws As Worksheet, lastRow As Long)
    Dim cell As Range
    For Each cell In ws.Range("F2:P" & lastRow)
        If IsEmpty(cell) Then
            cell.Value = "-"
        End If
    Next cell
End Sub



' アクティブシート内のグラフを削除
Sub DeleteAllChartsInActiveSheet()
    Dim chart As ChartObject
    
    For Each chart In ActiveSheet.ChartObjects
        chart.Delete
    Next chart
End Sub
