Attribute VB_Name = "BicycleCreate"
Public Sub Create_BicycleGraph()
    Call CreateGraphBicycle
    Call Bicycle_150G_DurationTime
    ' 開いているブックの一番左のシートを選択
    ThisWorkbook.Sheets(1).Select

    ' A1セルにカーソルを移動
    Range("A1").Select
End Sub


Function ColNumToLetter(colNum As Integer) As String
    ' CreateGraphBicycleにて使用する関数
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

Sub CreateGraphBicycle()

    ' ワークシートを宣言
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
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
        
        ' X軸のデータ範囲を設定
        chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))

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
        If maxVal <= 295 Then
            yAxis.MaximumScale = 300
        Else
            yAxis.MaximumScale = Int(maxVal) + 1
        End If

        yAxis.MinimumScale = -100 ' Y軸の最低値を-10に設定します。

        ' Y軸の TickLabels を設定
        With yAxis.TickLabels
            .NumberFormatLocal = "0""G"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With


        ' X軸の設定
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 100
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



Sub Bicycle_150G_DurationTime()
    '自転車帽試験のデータを処理するメインのサブルーチン
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("LOG_BICYCLE")
    lastRow = GetLastRow(ws, "V")
    
    '各行の最大値を色付けし、最大値の時間を記録します
    ColorAndRecordMaxVal ws, lastRow, 150
    
    '150G以上を記録した範囲を色付けし、その範囲の時間差を記録します
    ColorAndRecordTimeDifference ws, lastRow, 150
    
    '空のセルを"-"で埋めます
    FillEmptyCells ws, GetLastRow(ws, "B")
End Sub

'Bicycle150GDuration_特定のカラムの最終行を取得する関数です。
Function GetLastRow(ws As Worksheet, column As String) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
End Function

'Bicycle150GDuration_各行の最大値のセルを色付けし、最大値の時間を記録するサブルーチンです。
Sub ColorAndRecordMaxVal(ws As Worksheet, lastRow As Long, threshold As Double)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range
    
    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "V"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        
        Dim MaxValue As Double
        MaxValue = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "H").Value = MaxValue
        
        For Each cell In rng
            If cell.Value = MaxValue Then
                cell.Interior.color = RGB(255, 111, 56)
                ws.Cells(i, "I").Value = ws.Cells(1, cell.column).Value ' 対応する時間をI列に記録
                Exit For ' 最初の最大値が見つかったらループを抜ける
            End If
        Next cell
    Next i
End Sub

'Bicycle150GDuration_150G以上を記録した範囲を色付けし、その範囲の時間差を記録するサブルーチンです。
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
        Set rng = ws.Range(ws.Cells(i, "V"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        startRange150 = 0
        endRange150 = 0
        maxRange150 = 0
        Set rangeCollection150 = New Collection
        
        For Each cell In rng
            If cell.Value >= threshold Then
                If startRange150 = 0 Then startRange150 = cell.column
                endRange150 = cell.column
                cell.Interior.color = RGB(0, 138, 211)
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
            timeDifference150 = ws.Cells(1, maxEnd150).Value - ws.Cells(1, maxStart150).Value
            ws.Cells(i, "K").Value = timeDifference150
        Else
            ws.Cells(i, "K").Value = "-"
        End If
    Next i
End Sub

'Bicycle150GDuration_空のセルを"-"で埋めるサブルーチンです。
Sub FillEmptyCells(ws As Worksheet, lastRow As Long)
    Dim cellRng As Range
    
    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng
End Sub


