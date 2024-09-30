Attribute VB_Name = "Utlities"
Sub ShowFormInspectionType()
    ' ユーザーフォーム "Form_InspectionType" を表示
    Form_InspectionType.Show
End Sub
Sub ShowFormTenki()
    ' ユーザーフォーム "Form_Tenki" を表示
    Form_Tenki.Show
End Sub

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
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                With chartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                End With
            Next chartObj
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

Sub AdjustingDuplicateValues()
    ' 対象シート名のリスト
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Double
    Dim sheetName As Variant
    Dim newValue As Double
    Dim randomDigit As Integer
    Dim roundedValue As Double
    Dim maxCol As Long

    ' シートごとに処理
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row
        
        ' "最大値"を含むヘッダーがある列を検索
        maxCol = 0
        For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            If InStr(ws.Cells(1, i).value, "最大値") > 0 Then
                maxCol = i
                Exit For
            End If
        Next i
        
        ' "最大値"列が見つからなければ次のシートへ
        If maxCol = 0 Then
            MsgBox "シート " & sheetName & " には '最大値' を含む列が見つかりません。"
            GoTo NextSheet
        End If

        For i = 2 To lastRow
            ' セルの値が数値かどうか確認
            If IsNumeric(ws.Cells(i, maxCol).value) Then
                ' 数値として取得し、小数点以下2桁で丸める
                roundedValue = Round(ws.Cells(i, maxCol).value, 2)

                If ws.Cells(i, maxCol).Interior.colorIndex = xlNone Then
                    For j = i + 1 To lastRow
                        ' 重複値をチェック（数値チェックを追加）
                        If IsNumeric(ws.Cells(j, maxCol).value) And Round(ws.Cells(j, maxCol).value, 2) = roundedValue And ws.Cells(j, maxCol).Interior.colorIndex = xlNone Then
                            Debug.Print "Duplicate Row Number: " & j
                            Do
                                ' 1から9のランダムな数を生成
                                randomDigit = Int((9 - 1 + 1) * Rnd + 1)
                                ' 元の値にランダムな値を小数点以下4桁として追加（小数点以下2桁は維持）
                                newValue = roundedValue + randomDigit / 10000
                                Debug.Print "New Value: " & newValue
                            Loop While WorksheetFunction.CountIf(ws.Range(ws.Cells(2, maxCol), ws.Cells(lastRow, maxCol)), newValue) > 0
                            
                            ' 新しい値をセルに設定
                            ws.Cells(j, maxCol).value = newValue
                        End If
                    Next j
                End If
            End If
        Next i
NextSheet:
    Next sheetName
End Sub


Sub ListChartNames()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    ' Loop through all sheets in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            For Each chartObj In ws.ChartObjects
                ' Display the chart name
                MsgBox "シート名: " & ws.Name & vbCrLf & "グラフ名: " & chartObj.Name, vbInformation
            Next chartObj
        End If
    Next ws
End Sub


' 各列に書式設定をする
Public Sub CustomizeSheetFormats()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "ID") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試料ID") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "品番") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試験内容") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "検査日") > 0 Then ' Date
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToDate(rng)
            ElseIf InStr(1, cell.value, "温度") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "最大値(kN)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericFourDecimals(rng)
            ElseIf InStr(1, cell.value, "最大値の時間(ms)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "4.9kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "7.3kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "前処理") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "重量") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "天頂すきま") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "製品ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "帽体ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "内装ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "構造検査") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "耐貫通検査") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試験区分") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            End If
        Next cell
    Next sheet
End Sub

Sub ConvertToNumeric(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericTwoDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.00"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericFourDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0000"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToString(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "@"
    For Each cell In rng
        cell.value = CStr(cell.value)
    Next cell
End Sub

Sub ConvertToDate(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "yyyy/mm/dd"  ' 日付の表示形式を設定
    For Each cell In rng
        If IsDate(cell.value) Then
            cell.value = CDate(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub
' 空白セルに"-"を挿入
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
        lastCol = ws.Cells(1, "Z").column ' Z列の列番号を設定

        ' 2行目から最終行までループ（1行目はヘッダーと仮定）
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    'Debug.Print "EmptyCell:" & "Cells&("; i; "," & j; ")"
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

Sub DeleteAllChartsAndDataFromActiveSheet()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim proceed As Integer

    ' アクティブなシートを取得
    Set sheet = ThisWorkbook.ActiveSheet

    Application.DisplayAlerts = False

    ' グラフの削除
    For Each chart In sheet.ChartObjects
        chart.Delete
    Next chart

    ' B2セルからZZ15までのデータの有無をチェックし、有れば警告を表示
    If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
        Application.DisplayAlerts = True
        proceed = MsgBox("Sheet '" & sheet.Name & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
        Application.DisplayAlerts = False
        If proceed = vbNo Then Exit Sub
    End If

    Application.DisplayAlerts = True
End Sub

