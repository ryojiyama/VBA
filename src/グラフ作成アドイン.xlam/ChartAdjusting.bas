Attribute VB_Name = "ChartAdjusting"
'グラフのY軸の最大値を調整する。
Sub UniformizeLineGraphAxes()

    On Error GoTo ErrorHandler

    ' Display input dialog to set the maximum value for the axes
    Dim MaxValue As Variant
    MaxValue = InputBox("Y軸の最大値を入力してください。(整数)", "最大値を入力")

    ' Check if the user pressed Cancel
    If MaxValue = False Then
        MsgBox "操作がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' Validate the input
    If Not IsNumeric(MaxValue) Or MaxValue <= 0 Then
        MsgBox "有効な数値を入力してください。", vbExclamation
        Exit Sub
    End If

    MaxValue = CDbl(MaxValue)

    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                With chartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                    
                    ' Set the MajorUnit based on MaxValue
                    If MaxValue <= 5 Then
                        .MajorUnit = 1#
                    ElseIf MaxValue > 5 And MaxValue <= 25 Then
                        .MajorUnit = 2#
                    ElseIf MaxValue > 25 And MaxValue <= 100 Then
                        .MajorUnit = 10#
                    ElseIf MaxValue > 100 And MaxValue <= 300 Then
                        .MajorUnit = 50#
                    ElseIf MaxValue > 300 Then
                        .MajorUnit = 100#
                    End If
                End With
            Next chartObj
        End If
    Next ws

    MsgBox "すべてのシートのグラフのY軸の最大値を " & MaxValue & " に設定し、適切な目盛り間隔を設定しました。", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical

End Sub




' グラフの縦横比率を変更する。
Sub AdjustChartAspectRatio()
    Dim userChoice As Variant

    ' ユーザーにグラフの比率を選択させる
    userChoice = MsgBox("グラフの比率を、" & vbCrLf & _
                        "どのように変更しますか？" & vbCrLf & vbCrLf & _
                        "[はい] - 2列用に変更する" & vbCrLf & _
                        "[いいえ] - 3列用に変更する", vbYesNo + vbQuestion, "比率の選択")

    ' 選択に応じてプロシージャを実行
    Select Case userChoice
        Case vbYes
            Call SetChartRatio129
        Case vbNo
            Call SetChartRatio1110
        Case Else
            Exit Sub ' キャンセルされた場合は処理を終了
    End Select
End Sub

' "Impact" を含むシート内のグラフ比率を 480:360 にするプロシージャ
Sub SetChartRatio129()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    ' "Impact" を含むシートをループ処理
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' シート内のグラフをループ処理
            For Each chartObj In ws.ChartObjects
                chartObj.Width = 480
                chartObj.Height = 360
            Next chartObj
        End If
    Next ws
End Sub

' "Impact" を含むシート内のグラフ比率を 440:400 にするプロシージャ
Sub SetChartRatio1110()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    ' "Impact" を含むシートをループ処理
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' シート内のグラフをループ処理
            For Each chartObj In ws.ChartObjects
                chartObj.Width = 400
                chartObj.Height = 440
            Next chartObj
        End If
    Next ws
End Sub

