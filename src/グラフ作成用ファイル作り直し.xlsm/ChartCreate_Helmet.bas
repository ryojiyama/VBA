Attribute VB_Name = "ChartCreate_Helmet"
Sub HelmetTestResultChartBuilder()
    'グラフ作成とヘルメット検査時間の表示、色付けなど
    Call CreateGraphHelmet
    Call InspectHelmetDurationTime
    Call Utlities.AdjustingDuplicateValues
End Sub

' 列の終わりを決定する関数
Function GetColumnEnd(ByRef ws As Worksheet, ByVal rowNumber As Long) As String
    Dim lastCol As Long
    Dim col As Long
    Dim found As Boolean
    found = False

    ' 列の最後から開始して値が1.0を超える最後の列番号を探す
    For col = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If ws.Cells(rowNumber, col).value > 1# Then
            lastCol = col
            found = True
            Exit For
        End If
    Next col

    ' 値が1.0を超える列から100列後を計算
    If found Then
        lastCol = lastCol + 100
        If lastCol > ws.Columns.Count Then lastCol = ws.Columns.Count ' 列数の最大値を超えないように調整
    Else
        ' 1.0を超える値が見つからない場合は、適当なデフォルト値を設定するか、エラー処理を行う
        lastCol = 150
    End If

    ' 列番号から列のアドレスを取得し、行番号を削除
    Dim fullAddress As String
    fullAddress = ws.Cells(1, lastCol).Address(False, False)  ' 絶対参照を避ける
    GetColumnEnd = Replace(fullAddress, "1", "")  ' 行番号を削除
End Function



' グラフのサイズを決定する関数
Function GetChartSize(ByVal graphType As String) As Variant
    Dim size(1) As Long
    
    Select Case graphType
        Case "定期試験用"
            size(0) = 250  ' Width
            size(1) = 300  ' Height
        Case "型式申請試験用"
            size(0) = 300  ' Width
            size(1) = 350  ' Height
        Case Else
            size(0) = 400  ' Default Width
            size(1) = 250  ' Default Height
    End Select
    
    GetChartSize = size
End Function

' グラフを作成するメインのサブプロシージャ
Sub CreateGraphHelmet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim colStart As String
    Dim colEnd As String
    Dim chartSize As Variant
    Dim userInput As String

    userInput = InputBox("グラフの種類を入力してください（定期試験用、型式申請試験用、その他）")

    colStart = "GY"  ' 開始列を初期設定
    chartTop = ws.Rows(lastRow + 1).Top + 10
    chartLeft = 250

    For i = 2 To lastRow
        colEnd = GetColumnEnd(ws, i)
        chartSize = GetChartSize(userInput)
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd, chartSize
        chartLeft = chartLeft + 10 ' 次のグラフの左位置を調整
    Next i

End Sub



' CreateGraphHelmet_個別のグラフを設定・追加するサブプロシージャ
Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String, ByVal chartSize As Variant)
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))
    ws.Cells(i, "H").value = maxVal
    
    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartSize(0), Top:=chartTop, Height:=chartSize(1))
    Dim chart As chart
    Set chart = ChartObj.chart
    
    With chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))  ' X軸の範囲を1行目から設定
        .SeriesCollection(1).Values = ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))  ' Y軸のデータ範囲を設定
        .SeriesCollection(1).Name = "Data Series " & i
    End With
    
    ConfigureChart chart, ws, i, colStart, colEnd, maxVal
End Sub
' CreateGraphHelmet_グラフの書式設定をするサブプロシージャ
Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    'このプロシージャでX軸とY軸の目盛線を追加する。そうしないとうまくいかない。
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.ChartTitle.Text = ws.Cells(i, "B").value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 1#

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

' CreateGraphHelmet_Y軸の書式設定
Sub SetYAxis(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal maxVal As Double)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    If maxVal >= 5# Then
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2# '2.0刻み
    Else
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1# '1.0刻み
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With

End Sub


'CreateGraphHelmet_X軸の書式設定
Sub SetXAxis(ByRef chart As chart)
    Dim xAxis As Axis
    Set xAxis = chart.Axes(xlCategory, xlPrimary)

    xAxis.TickLabelSpacing = 100
    xAxis.TickMarkSpacing = 100

    With xAxis.TickLabels
        .NumberFormatLocal = "0.0""ms"""
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With

End Sub


Sub InspectHelmetDurationTime()
    ' ヘルメット試験において最大値の更新、最大値の時間の更新、試験内容の更新、継続時間の色分けを行う
    Dim ws As Worksheet
    Dim lastRow As Long

    ' "LOG_Helmet" シートを指定する。
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' 最終行を取得する。
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).row

    ' 各行を処理する。
    For i = 2 To lastRow
        UpdateMaxValueInRow ws, i             ' 行内の最大値を更新する。
        UpdatePartOfHelmet ws, i              ' ヘルメットの部分を更新する。
        UpdateRangeForThresholds ws, i, 4.9, "J"  ' 閾値の範囲を更新する。
        UpdateRangeForThresholds ws, i, 7.35, "K"
    Next i

End Sub

'InspectHelmetDurationTime()内のプロシージャ_行内の最大値を更新し、最大値を記録した時刻のセルに色をつける。
Sub UpdateMaxValueInRow(ByRef ws As Worksheet, ByVal row As Long)
    
    Dim rng As Range
    Dim MaxValue As Double
    Dim maxValueColumn As Long

    ' 行内の範囲をセットする。
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))
    ' 最大値を取得する。
    MaxValue = Application.WorksheetFunction.Max(rng)
    ws.Cells(row, "H").value = MaxValue

    ' 最大値の位置を見つける。
    For j = 1 To rng.Columns.Count
        If rng(1, j).value = MaxValue Then
            maxValueColumn = j
            rng(1, j).Interior.Color = RGB(250, 150, 0)  ' 色を設定する。
            Exit For
        End If
    Next j

    If maxValueColumn <> 0 Then
        ws.Cells(row, "I").value = ws.Cells(1, "V").Offset(0, maxValueColumn - 1).value
    End If
End Sub
 'InspectHelmetDurationTime()内のプロシージャ_ヘルメットの試験箇所を更新する
Sub UpdatePartOfHelmet(ByRef ws As Worksheet, ByVal row As Long)

    Dim cellValue As String
    cellValue = ws.Cells(row, "B").value
    
    ' 既存の値を取得する。
    Dim existingValue As String
    existingValue = ws.Cells(row, "E").value

    ' ヘルメットの部分に基づいて値を設定する。ただし、"天頂"や"頭部"がすでに含まれている場合は変更しない。
    ' 条件節では最初にE列の値をチェックする。
    If InStr(existingValue, "天頂") > 0 Or InStr(existingValue, "頭部") > 0 Then
    ElseIf InStr(cellValue, "HEL_TOP") > 0 Then
        ws.Cells(row, "E").value = "天頂"
    ElseIf InStr(cellValue, "HEL_ZENGO") > 0 Then
        ws.Cells(row, "E").value = "前後頭部"
    ElseIf InStr(cellValue, "HEL_SIDE") > 0 Then
        ws.Cells(row, "E").value = "側頭部"
    End If
End Sub

'InspectHelmetDurationTime()から4.9、7.35の範囲値の色付けと衝撃時間を記入する。
Sub UpdateRangeForThresholds(ByRef ws As Worksheet, ByVal row As Long, ByVal threshold As Double, ByVal columnToWrite As String)

    Dim rng As Range, cell As Range
    Dim startRange As Long, endRange As Long, maxRange As Long
    Dim rangeCollection As New Collection
    Dim timeDifference As Double

    ' 行の範囲をセットする。
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))

    ' 閾値を超えるセルの範囲を見つける。
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

    For Each Item In rangeCollection
        If (Item(1) - Item(0) + 1) > maxRange Then
            maxRange = Item(1) - Item(0) + 1
            startRange = Item(0)
            endRange = Item(1)
        End If
    Next Item
    
    If startRange > 0 And endRange > 0 Then
        timeDifference = ws.Cells(1, endRange).value - ws.Cells(1, startRange).value
        If timeDifference < 0 Then
            timeDifference = 0
        End If
        ws.Cells(row, columnToWrite).value = timeDifference
    Else
        ws.Cells(row, columnToWrite).value = 0
    End If
End Sub

