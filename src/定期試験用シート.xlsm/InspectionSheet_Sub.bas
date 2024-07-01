Attribute VB_Name = "InspectionSheet_Sub"


Sub MakeInspectionSheets()
    'Call CreateInspectionSheetIDs
    Call DuplicateAndRenameSheets
    Call TransferDataToTopImpactTest
    Call RenameAndRemoveDuplicateSheets
    Call TransferDataToDynamicSheets
    Call ImpactValueJudgement
    Call FormatNonContinuousCells
    MsgBox "検査票シートの作成が終了しました"
End Sub


Sub DuplicateAndRenameSheets()
    Dim wsLogHelmet As Worksheet, wsTemplate As Worksheet, wsDraft As Worksheet
    Dim i As Long
    Dim part1Result As Boolean
    Dim sheetName As String, value As String, part1 As String, part2 As String

    Const LOG_HELMET As String = "Log_Helmet"
    Const TEMPLATE_SHEET As String = "InspectionSheet"

    Set wsLogHelmet = ThisWorkbook.Sheets(LOG_HELMET)
    Set wsTemplate = ThisWorkbook.Sheets(TEMPLATE_SHEET)

    ' シートの複製と名前の設定
    For i = 2 To wsLogHelmet.Cells(wsLogHelmet.Rows.count, 2).End(xlUp).row
        value = wsLogHelmet.Cells(i, 2).value
        
        ' 文字列にキャストして安全に関数に渡す
        part1 = CStr(Split(value, "-")(1))
        part2 = CStr(Split(value, "-")(2))
        part1Result = CheckPart1(part1)
        sheetName = ExtractSheetName(value)

        ' デバッグ情報の出力
        Debug.Print "Row: " & i & ", Value: " & value & ", Part1: " & part1 & ", Part2: " & part2
        Debug.Print "Part1Result: " & part1Result & ", SheetName: " & sheetName
        Debug.Print "Should Duplicate: " & (part1Result Or (Not part1Result And part2 = "天"))
        
        ' part1ResultがTrueか、Falseでもpart2が"天"の場合にシートを複製
        If part1Result Or (Not part1Result And part2 = "天") Then
            If sheetName <> "" Then
                wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                Set wsDraft = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                If Not SheetExists(sheetName) Then
                    wsDraft.name = sheetName
                End If
                Debug.Print "Sheet Duplicated: " & sheetName
            Else
                Debug.Print "Sheet not duplicated due to empty sheet name."
            End If
        Else
            Debug.Print "Conditions not met for duplication."
        End If
    Next i
End Sub



Function ExtractSheetName(fullName As String) As String
    Dim parts As Variant
    parts = Split(fullName, "-")
    
    If UBound(parts) >= 2 Then
        ' CheckPart1の結果に関わらず、part(2)が"天"の場合はシート名を生成
        If parts(2) = "天" Then
            ' parts(1)が"F"を含む場合はその部分を除いてシート名を生成
            Dim cleanPart1 As String
            cleanPart1 = Replace(parts(1), "F", "")  ' "300F"から"F"を削除
            ExtractSheetName = parts(0) & "-" & cleanPart1 & "-" & parts(2)
        Else
            ExtractSheetName = ""  ' 条件に合致しない場合は空文字列を返す
        End If
    Else
        ExtractSheetName = ""  ' 適切なパーツがない場合は空文字列を返す
    End If
End Function



Function CheckPart1(part As String) As Boolean
    ' 文字列の末尾がFでなければTrue、FであればFalseを返す
    CheckPart1 = Not Right(part, 1) = "F"
End Function



Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

Sub DeleteSheet()
    Dim ws As Worksheet
    On Error Resume Next ' エラーが発生した場合、次の行へ進む
    Set ws = ThisWorkbook.Sheets("ID")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' 削除の確認メッセージを表示しない
        ws.Delete
        Application.DisplayAlerts = True ' メッセージ表示を元に戻す
    End If
    On Error GoTo 0 ' エラーハンドリングを元に戻す
End Sub


Sub TransferDataToTopImpactTest()
    '天頂試験のみのシートを作成する。
    '"Log_Helmet"からコピーした検査票に値を転記する。
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dashPosSource As Integer
    Dim dashPosDest As Integer
    Dim matchName As String
    Dim TemperatureCondition As String

    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("Log_Helmet")

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 2行目から最終行までループ
    For i = 2 To lastRow
        ' C列の1文字目が"F"でない行を探す
        If Left(wsSource.Cells(i, 3).value, 1) <> "F" Then
            ' MatchNameを取得（C列の1文字目から"-"まで）
            dashPosSource = InStr(wsSource.Cells(i, 3).value, "-")
            If dashPosSource > 0 Then
                matchName = Left(wsSource.Cells(i, 3).value, dashPosSource - 1)

                ' L列の値に基づいて条件を設定
                Select Case wsSource.Cells(i, 12).value
                    Case "高温"
                        TemperatureCondition = "Hot"
                    Case "低温"
                        TemperatureCondition = "Cold"
                    Case "浸せき"
                        TemperatureCondition = "Wet"
                    Case Else
                        TemperatureCondition = ""
                End Select

                ' ワークシートの名前をループして条件をチェック
                For Each wsDestination In ThisWorkbook.Sheets
                    dashPosDest = InStr(wsDestination.name, "-")
                    If dashPosDest > 0 Then
                        If Left(wsDestination.name, dashPosDest - 1) = matchName And InStr(wsDestination.name, TemperatureCondition) > 0 Then
                            ' 条件が当てはまったら転記
                            wsDestination.Range("C2").value = wsSource.Cells(i, 21).value '試験内容
                            wsDestination.Range("F2").value = wsSource.Cells(i, 6).value '検査日
                            wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                            wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                            wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                            wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                            wsDestination.Range("C4").value = wsSource.Cells(i, 16).value 'Lot
                            wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                            wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                            wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                            wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                            wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("A10").value = "※前処理：" & wsSource.Cells(i, 12).value
                            wsDestination.Range("A14").value = "検査対象外"
                            wsDestination.Range("A19").value = "検査対象外"
                        End If
                    End If
                Next wsDestination
            End If
        End If
    Next i
End Sub




Sub RenameAndRemoveDuplicateSheets()
    'フィルタリングし易いようにシート名を改変「F390F-Cold」の形式にする。
    Dim ws As Worksheet
    Dim parts() As String
    Dim newName As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 重複する名前を持つシートを特定し、削除
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.name, 1) = "F" Then
            parts = Split(ws.name, "-")
            If UBound(parts) >= 2 Then
                newName = parts(0) & "-" & parts(1)
                If dict.Exists(newName) Then
                    Application.DisplayAlerts = False
                    ws.Delete
                    Application.DisplayAlerts = True
                Else
                    dict.Add newName, newName
                End If
            End If
        End If
    Next ws

    ' 重複を削除した後、シート名を変更
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.name, 1) = "F" Or InStr(ws.name, "-") > 0 Then
            parts = Split(ws.name, "-")
            If UBound(parts) >= 2 Then
                newName = parts(0) & "-" & parts(1)
                On Error Resume Next
                ws.name = newName
                On Error GoTo 0
            End If
        End If
    Next ws
End Sub

Sub TransferDataToDynamicSheets()
    'F付き帽体の試験票を作成する。
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim modifiedSourceData As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).value
        checkData = wsSource.Cells(i, 5).value
        parts = Split(sourceData, "-")

        If UBound(parts) >= 2 Then
            ' データとシート名を作成
            modifiedSourceData = parts(0) & "-" & parts(1)
            destinationSheetName = modifiedSourceData

            ' modifiedSourceData と sourceData の最初の2つの部分が一致する場合にのみ転記
            If Left(sourceData, Len(modifiedSourceData)) = modifiedSourceData Then
                ' シートが存在するか確認し、存在する場合のみ転記
                If InspectionSheetExists(destinationSheetName) Then
                    Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)

                    ' parts(UBound(parts))に基づいて処理を分岐
                    Select Case parts(UBound(parts))
                        Case "天"
                            If checkData = "天頂" Then
                                ' 天に関するデータ転記の処理
                                wsDestination.Range("C2").value = wsSource.Cells(i, 21).value '試験内容
                                wsDestination.Range("F2").value = wsSource.Cells(i, 6).value '検査日
                                wsDestination.Range("H2").value = wsSource.Cells(i, 7).value '温度
                                wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                                wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                                wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                                wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                                wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                                wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                                wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                                wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                                wsDestination.Range("A10").value = "※前処理：" & wsSource.Cells(i, 12).value
                                wsDestination.Range("E11").value = wsSource.Cells(i, 8).value '衝撃値
                            End If

                        Case "前"
                            If checkData = "前頭部" Then
                                ' 前に関するデータ転記の処理
                                wsDestination.Range("E13").value = wsSource.Cells(i, 8).value '衝撃値
                                wsDestination.Range("E14").value = wsSource.Cells(i, 10).value '4.90kN
                                wsDestination.Range("E15").value = wsSource.Cells(i, 11).value '7.35kN
                                wsDestination.Range("A13").value = "前頭部"
                            End If

                        Case "後"
                            If checkData = "後頭部" Then
                                ' 後に関するデータ転記の処理
                                wsDestination.Range("E17").value = wsSource.Cells(i, 8).value '衝撃値
                                wsDestination.Range("E18").value = wsSource.Cells(i, 10).value '4.90kN
                                wsDestination.Range("E19").value = wsSource.Cells(i, 11).value '7.35kN
                                wsDestination.Range("A17").value = "後頭部"
                            End If

                        Case Else
                            ' その他の値の場合の処理（必要に応じて）
                    End Select
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' シートが存在するかどうかを確認する関数
Function InspectionSheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    InspectionSheetExists = Not sheet Is Nothing
End Function


Sub ImpactValueJudgement()
    '衝撃吸収試験の結果を各検査票シートの衝撃値から判定する。
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim resultE11 As Boolean, resultE14 As Boolean, resultE19 As Boolean

    ' "LOG_Helmet"シートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' C列の最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' C列の2行目から最終行までループ
    For i = 2 To lastRow
        sheetName = wsSource.Cells(i, "C").value
        ' 対象のシート名とIDをあわせる処理
        sheetName = Left(sheetName, Len(sheetName) - 2)

        ' 対象のシートを設定
        Set wsTarget = ThisWorkbook.Sheets(sheetName)

        ' D11, D14, D19の値を基に判定
        resultE11 = wsTarget.Range("E11").value <= 4.9
        resultE14 = IsEmpty(wsTarget.Range("E13")) Or wsTarget.Range("E13").value <= 9.81
        resultE19 = IsEmpty(wsTarget.Range("E17")) Or wsTarget.Range("E17").value <= 9.81

        ' 全ての条件がTrueの場合は"合格"、それ以外は"不合格"をG9に記入
        If resultE11 And resultE14 And resultE19 Then
            wsTarget.Range("H9").value = "合格"
        Else
            wsTarget.Range("H9").value = "不合格"
        End If
    Next i
End Sub


Sub FormatNonContinuousCells()
    ' コピーした検査票に書式を設定する。
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmetシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' B列の最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    ' B列の各行をループ
    For i = 2 To lastRow
        sheetName = wsSource.Cells(i, 2).value
        ' 対象のシート名とIDをあわせる処理
        sheetName = Left(sheetName, Len(sheetName) - 2)

        ' ワークシートが存在するかチェック
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        ' ワークシートが存在すれば、指定したセル範囲に書式を設定
        If Not wsTarget Is Nothing Then
            ' 範囲と書式設定を関連付け
            FormatRange wsTarget.Range("E7"), "游明朝", 12, True
            FormatRange wsTarget.Range("E8"), "游明朝", 12, True
            FormatRange wsTarget.Range("E9"), "游明朝", 12, True

            ' E13に値がない場合、A14:E14とB15:D16をグレーアウト
            If IsEmpty(wsTarget.Range("E13").value) Then
                wsTarget.Range("A13").value = "検査対象外"
                FormatRange wsTarget.Range("A13"), "游ゴシック", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B13:F13, B14:E15"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A13"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("E13:E15"), "游ゴシック", 10, False, RGB(255, 255, 255) 'E13:E15に注意
            End If

            ' E17に値がない場合、A19:E19とB20:D21をグレーアウト
            If IsEmpty(wsTarget.Range("E17").value) Then
                wsTarget.Range("A17").value = "検査対象外"
                FormatRange wsTarget.Range("A17"), "游ゴシック", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B17:F17, B18:E19"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A17"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("E17:E19"), "游ゴシック", 10, False, RGB(255, 255, 255) 'E17:E19に注意
            End If
            FormatSpecificEndStrings wsTarget.Range("A10"), "游ゴシック", 12, True '前処理を目立たせる_書くところがないのでここに書く
            With wsTarget.Range("C2:C4, F2:F4, H2:H4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("F3").NumberFormat = "0.0"" g"""
            wsTarget.Range("H2").NumberFormat = "0"" ℃"""
            wsTarget.Range("H3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("E11, E14, E19").NumberFormat = "0.00"" kN"""
            wsTarget.Range("E14:E15, E18:E19").NumberFormat = "0.00"" ms"""
            ' 他の範囲も同様に設定可能
            ' FormatRange wsTarget.Range("その他の範囲"), "フォント名", フォントサイズ, 太字かどうか, 背景色

            Set wsTarget = Nothing
        End If
    Next i
End Sub


Sub FormatSpecificEndStrings(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean)
    ' セルの特定の文字(前処理)に書式を適用するサブプロシージャ
    Dim cell As Range

    For Each cell In rng
        Dim text As String
        text = cell.value
        Dim textLength As Integer
        textLength = Len(text)

        If textLength >= 2 Then
            If Right(text, 2) = "高温" Or Right(text, 2) = "低温" Then
                With cell.Characters(Start:=textLength - 1, Length:=2).Font
                    .name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "浸せき" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub


' 範囲に書式を適用するためのサブプロシージャ
Sub FormatRange(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean, Optional bgColor As Variant)
    With rng
        .Font.name = fontName
        .Font.Size = fontSize
        .Font.Bold = isBold
        If Not IsMissing(bgColor) Then
            .Interior.Color = bgColor
        Else
            .Interior.colorIndex = xlColorIndexAutomatic ' 背景色を自動に設定
        End If
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

' セルの特定の文字列のみ(ここでは高温のみ)に書式を適用するプロシージャ浸せきがひらがなになったため、FormatSpecificTextにとって変わられた。
Sub FormatLastTwoCharacters(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean)
    Dim cell As Range
    Dim lastTwoChars As String

    For Each cell In rng
        If Len(cell.value) >= 2 Then
            lastTwoChars = Right(cell.value, 2)
            ' ここでlastTwoCharsに対して特定の書式を適用する
            ' ただし、VBAでは部分的なセルの書式設定は直接できないため、
            ' 文字列全体に書式を適用し、その後で最後の2文字だけ別の書式を適用する
            With cell
                .Font.name = "游ゴシック"
                .Font.Size = 10
                .Font.Bold = False
                ' 最後の2文字に特定の書式を適用する
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.name = "游ゴシック"
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.Size = 12
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.Bold = True
            End With
        End If
    Next cell
End Sub


Sub PrintFirstPageOfUniqueListedSheets()
    ' 指定された検査票の1ページ目を、重複なく1回ずつ印刷するプロシージャ
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim printedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set printedSheets = New Collection ' 印刷されたシート名を追跡するコレクション

    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 2).value

        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        On Error Resume Next
        ' コレクションに同じ名前が既に存在するかチェック
        printedSheets.Add sheetName, sheetName
        If Err.number = 0 Then ' 追加が成功した場合、シートはまだ印刷されていない
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' シートの1ページ目のみを印刷
            End If
        End If
        On Error GoTo 0 ' エラーハンドリングをリセット

        Set wsTarget = Nothing
    Next i
End Sub



Sub ModifyAndStoreChartTitles()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartTitles() As String
    Dim i As Integer
    Dim parts() As String
    Dim modifiedChartTitle As String

    Set ws = ThisWorkbook.Sheets("LOG_Helmet") ' 実際のシート名に置き換えてください

    ReDim chartTitles(1 To ws.ChartObjects.count)

    i = 1
    For Each chartObj In ws.ChartObjects
        ' チャートタイトルを"-"で分割
        parts = Split(chartObj.chart.ChartTitle.text, "-")

        ' 最初の2つの部分を組み合わせて新しいタイトルを生成
        If UBound(parts) >= 1 Then
            modifiedChartTitle = parts(0) & "-" & parts(1)
        Else
            ' 分割できない場合は元のタイトルを使用
            modifiedChartTitle = chartObj.chart.ChartTitle.text
        End If

        ' 改変後のタイトルを配列に格納
        chartTitles(i) = modifiedChartTitle
        i = i + 1
    Next chartObj

    ' テスト出力
    For i = 1 To UBound(chartTitles)
        Debug.Print "Chart" & i & ": " & chartTitles(i)
    Next i
End Sub
Sub CopyChartToMatchingSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim chartObj As ChartObject
    Dim parts() As String
    Dim modifiedChartTitle As String
    Dim wb As Workbook
    
    Set wb = ThisWorkbook ' 現在のワークブックを設定
    Set wsSource = wb.Sheets("LOG_Helmet") ' ソースシート名を指定

    ' ソースシートの全チャートをループ
    For Each chartObj In wsSource.ChartObjects
        ' チャートタイトルを"-"で分割
        parts = Split(chartObj.chart.ChartTitle.text, "-")
        
        ' 最初の2つの部分を組み合わせて新しいタイトルを生成
        If UBound(parts) >= 1 Then
            modifiedChartTitle = parts(0) & "-" & parts(1)
        Else
            ' 分割できない場合は元のタイトルを使用
            modifiedChartTitle = chartObj.chart.ChartTitle.text
        End If
        
        ' ワークブックの全シートをループ
        For Each wsDest In wb.Sheets
            ' チャートのタイトルがシート名と一致する場合、チャートをコピー＆ペースト
            If wsDest.name = modifiedChartTitle Then
                ' チャートをコピー
                ' チャートをコピー
Dim tryCount As Integer
tryCount = 0
Do
    On Error Resume Next
    chartObj.Copy
    If Err.number = 0 Then Exit Do ' コピーに成功したらループを抜ける
    On Error GoTo 0
    tryCount = tryCount + 1
    If tryCount > 5 Then ' 5回試行してダメならエラーを出力
        MsgBox "チャートのコピーに失敗しました: " & chartObj.name
        Exit Sub
    End If
    Application.Wait Now + TimeValue("00:00:01") ' 1秒待って再試行
Loop
                
                ' シートにペースト
                With wsDest
                    .Activate
                    .Paste
                    ' 貼り付けたチャートの位置を調整（例: A1の位置に配置）
                    .Shapes(.Shapes.count).Top = .Range("A1").Top
                    .Shapes(.Shapes.count).Left = .Range("A1").Left
                End With
            End If
        Next wsDest
    Next chartObj
End Sub


'Sub ClearDataFromAllListedSheetsWithMergedCells()
'    '転記した項目を消すプロシージャ
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim lastRow As Long
'    Dim i As Long
'    Dim sheetName As String
'
'    ' LOG_Helmetシートを設定
'    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
'
'    ' B列の最終行を取得
'    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
'
'    ' B列の各行をループ
'    For i = 2 To lastRow
'        sheetName = wsSource.Cells(i, 2).value
'
'        If Left(sheetName, 1) = "F" Then
'            sheetName = Left(sheetName, Len(sheetName) - 2)
'        End If
'
'        ' ワークシートが存在するかチェック
'        On Error Resume Next
'        Set wsTarget = ThisWorkbook.Sheets(sheetName)
'        On Error GoTo 0
'
'        ' ワークシートが存在すれば、指定した結合セルからデータをクリア
'        If Not wsTarget Is Nothing Then
'            ' ここで結合セルの範囲を指定してください
'            wsTarget.Range("C2:C4", "F2:F4", "H2:H4").ClearContents
'            wsTarget.Range("H7:H9").ClearContents
'            wsTarget.Range("E11:F11").ClearContents
'            wsTarget.Range("E13:E15").ClearContents
'            wsTarget.Range("F13").ClearContents
'            wsTarget.Range("E17:E19").ClearContents
'            wsTarget.Range("F17").ClearContents
'            wsTarget.Range("A10").ClearContents
'            ' 以下、必要な範囲に合わせて追加
'
'            Set wsTarget = Nothing
'        End If
'    Next i
'End Sub


Sub DeleteAllListedSheets()
    ' 複製された検査票を削除するプロシージャ
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmetシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' B列の最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    ' B列の各行をループ
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 2).value

        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        ' ワークシートが存在するかチェック
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        If Not wsTarget Is Nothing Then
            Application.DisplayAlerts = False ' 警告の表示をオフにする
            wsTarget.Delete ' シートを削除
            Application.DisplayAlerts = True ' 警告の表示をオンに戻す
        End If
        On Error GoTo 0

        Set wsTarget = Nothing
    Next i
End Sub

