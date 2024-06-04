Attribute VB_Name = "Utlities"
' DeleteAllChartsAndSheets_シート中のグラフと余計なシートを削除する
Sub DeleteAllChartsAndSheets()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String
    Dim proceed As Integer

    ' シートのリスト
    Dim sheetList() As Variant
    sheetList = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")

    Application.DisplayAlerts = False

    ' 各シートに対して処理を実行
    For Each sheet In ThisWorkbook.Sheets
        sheetName = sheet.Name
        ' グラフの削除とデータの警告表示
        If IsInArray(sheetName, sheetList) Then
            For Each chart In sheet.ChartObjects
                chart.Delete
            Next chart
            ' B2セルからZZ15までのデータの有無をチェックし、有れば警告を表示
            If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
                Application.DisplayAlerts = True
                proceed = MsgBox("Sheet '" & sheetName & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
                Application.DisplayAlerts = False
                If proceed = vbNo Then Exit Sub
            End If
        ' シートの削除
        ElseIf sheetName <> "Setting" And sheetName <> "Hel_SpecSheet" Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True


End Sub

' DeleteAllChartsAndSheets_配列内に特定の値が存在するかチェックする関数
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

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
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim ChartObj As ChartObject
            For Each ChartObj In ws.ChartObjects
                With ChartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                End With
            Next ChartObj
        End If
    Next ws

    MsgBox "すべてのシートのグラフのY軸の最大値を " & MaxValue & " に設定しました。", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical

End Sub


Sub HighlightDuplicateValues()
    ' 対象シート名のリスト
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    Dim sheetName As Variant

    ' シートごとに処理
    For Each sheetName In sheetNames
        ' シートオブジェクトを設定
        Set ws = ThisWorkbook.Sheets(sheetName)

        ' 最終行を取得
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

        ' 色のインデックスを初期化
        colorIndex = 3 ' Excelの色インデックスは3から始まる

        For i = 2 To lastRow
            ' 現在のセルの値を取得
            valueToFind = ws.Cells(i, "H").value

            ' 同じ値を持つセルが既に色付けされていないかチェック
            If ws.Cells(i, "H").Interior.colorIndex = xlNone Then
                For j = i + 1 To lastRow
                    If ws.Cells(j, "H").value = valueToFind And ws.Cells(j, "H").Interior.colorIndex = xlNone Then
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
    Next sheetName
End Sub

Public Sub FillBlanksWithHyphenInMultipleSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim sheetName As Variant

    ' 対象シートの名前を配列に設定
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' 各シートについて処理を行う
    For Each sheetName In sheetNames
        On Error Resume Next
        ' 対象シートを設定
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = Nothing ' ws変数をクリア
            GoTo NextSheet ' 次のシートに進む
        End If

        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        lastCol = ws.Cells(1, "Z").Column ' Z列の列番号を設定

        ' 2行目から最終行までループ（1行目はヘッダーと仮定）
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").Column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    ws.Cells(i, j).value = "-"
                End If
            Next j
        Next i

        ' シート処理の終了ラベル
NextSheet:
        ' 次のシートの処理に移る前に変数をクリア
        Set ws = Nothing
    Next sheetName
End Sub
