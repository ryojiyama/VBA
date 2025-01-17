Attribute VB_Name = "Test"
Sub CheckAndMarkRecords()
    Dim wsSource As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, checkRow As Long
    Dim targetLastRow As Long
    Dim foundSheets As Collection
    Dim PLNum As String
    Dim hasFailedRecord As Boolean
    Dim failedRow As Range
    Dim clearRange As Range
    Dim isAllPass As Boolean    ' 全シート合格フラグを追加
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "D").End(xlUp).Row
    
    ' Excelのパフォーマンス向上のための設定
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 該当するシートを探索するためのコレクション作成
    Set foundSheets = New Collection
    
    ' D列から対象シートを探索
    For checkRow = 2 To lastRow
        PLNum = wsSource.Cells(checkRow, "D").value
        
        ' ワークブック内の全シートをチェック
        For Each ws In ThisWorkbook.Worksheets
            ' シート名が "PLNum_数字" の形式と一致するかチェック
            If ws.Name Like PLNum & "_[0-9]*" Then
                ' 重複を避けるため、既に追加されていないかチェック
                On Error Resume Next
                foundSheets.Add ws, ws.Name
                On Error GoTo ErrorHandler
            End If
        Next ws
    Next checkRow
    
    ' 見つかったシートがない場合の処理
    If foundSheets.count = 0 Then
        MsgBox "対象となるシートが見つかりません。", vbExclamation
        GoTo CleanExit
    End If
    
    ' 全シート合格フラグを初期化
    isAllPass = True
    
    ' 各シートに対して処理を実行
    For Each ws In foundSheets
        hasFailedRecord = False
        targetLastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
        
        ' 既存の"不合格"行を削除
        On Error Resume Next
        ws.Rows(targetLastRow + 1).Delete
        On Error GoTo ErrorHandler
        
        ' 既存の色付けをクリア
        Set clearRange = ws.Range(ws.Cells(30, "B"), ws.Cells(targetLastRow, "U"))
        clearRange.Interior.ColorIndex = xlNone
        
        ' H18セルの内容をクリア
        ws.Range("H18").value = ""
        
        ' 30行目から最終行までチェック
        For checkRow = 30 To targetLastRow
            ' D列の値が"PLNum"と一致するレコードをチェック
            If ws.Cells(checkRow, "D").value = PLNum Then
                ' J列とL列の値を取得
                Dim jValue As Variant
                Dim lValue As Variant
                
                jValue = ws.Cells(checkRow, "J").value
                lValue = ws.Cells(checkRow, "L").value
                
                ' J列の数値チェックと条件判定
                If Not IsNumeric(jValue) Then
                    jValue = 0
                End If
                
                ' L列の数値チェックと条件判定
                If Not IsNumeric(lValue) Then
                    lValue = 0
                End If
                
                ' 条件チェック
                If CDbl(jValue) >= 300 Or CDbl(lValue) >= 4 Then
                    ' B列からU列を色付け
                    ws.Range(ws.Cells(checkRow, "B"), _
                            ws.Cells(checkRow, "U")).Interior.Color = RGB(255, 153, 153)
                    hasFailedRecord = True
                End If
            End If
        Next checkRow
        
        ' 条件を満たすレコードが1つでもあった場合、不合格を入力
        If hasFailedRecord Then
            isAllPass = False   ' 不合格があった場合、全シート合格フラグをfalseに
            
            ' 最終行を再取得
            targetLastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
            
            ' 不合格行の設定
            Set failedRow = ws.Range(ws.Cells(targetLastRow + 1, "A"), _
                                   ws.Cells(targetLastRow + 1, "U"))
            
            With failedRow
                ' セルの結合
                .Merge
                ' 不合格テキストの入力
                .value = "不合格 ※ J列が300以上 または L列が4以上 のレコードが存在します"
                ' セルの書式設定
                With .Interior
                    .Color = RGB(255, 153, 153)
                End With
                With .Font
                    .Bold = True
                    .Size = 12
                    .Color = RGB(192, 0, 0)
                End With
                ' 配置設定
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' 行の高さを調整
            ws.Rows(targetLastRow + 1).RowHeight = 25
        End If
    Next ws
    
    ' すべてのシートの処理が終わった後、全シート合格なら合格を表示
    If isAllPass Then
        For Each ws In foundSheets
            With ws.Range("H18")
                .value = "合格"
                With .Font
                    .Bold = True
                    .Size = 12
                    .Color = RGB(0, 176, 80)  ' 緑色
                End With
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next ws
    End If
    
CleanExit:
    ' Excelの設定を元に戻す
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    ' エラー発生時の処理
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' productName_1のシートの体裁を整える。
Sub CustomizeReportIntroduction()

    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim destinationSheetName As String

    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row
    
    ' Excelのパフォーマンス向上のための設定
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSourceのC列をループしてデータを処理
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, "B").value
        checkData = wsSource.Cells(i, 5).value
        parts = Split(sourceData, "-")

        ' シート名の生成
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(0) & "-" & parts(1)

            ' 転記先シートの存在確認
            On Error Resume Next
            Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
            On Error GoTo 0

            ' シートが存在し、かつ条件が一致する場合にデータを転記
            If Not wsDestination Is Nothing Then
                Select Case parts(2)
                    Case "天"
                        If checkData = "天頂" Then
                            ' 天に関するデータ転記
                            wsDestination.Range("C2").value = wsSource.Cells(i, 21).value
                            wsDestination.Range("F2").value = wsSource.Cells(i, 6).value
                            wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                            wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                            wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                            wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                            wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                            wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                            wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                            wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                            wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                            wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("A10").value = "※前処理：" & wsSource.Cells(i, 12).value
                        End If
                    Case "前"
                        If checkData = "前頭部" Then
                            ' 前頭部に関するデータ転記
                            wsDestination.Range("E13").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E14").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E15").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A13").value = "前頭部"
                        End If
                    Case "後"
                        If checkData = "後頭部" Then
                            ' 後頭部に関するデータ転記
                            wsDestination.Range("E17").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E18").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E19").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A17").value = "後頭部"
                        End If
                End Select
            End If
        End If
    Next i
    
    ' Excelの設定を元に戻す
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



Function GetTargetSheetNames() As Collection
    ' CopiedSheetNamesシートのA列からシート名を取得
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetNames As New Collection
    
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        sheetNames.Add ws.Cells(i, 1).value
    Next i
    
    Set GetTargetSheetNames = sheetNames
End Function
    ' CopiedSheetNamesシートのA列に基づいて検査票に書式を設定する
Sub FormatNonContinuousCells()
    Dim wsTarget As Worksheet
    Dim i As Long
    Dim sheetName As String
    Dim targetSheets As Collection
    Dim rng As Range
    Dim cell As Range
    
    ' 処理するシート名を取得
    Set targetSheets = GetTargetSheetNames()
    
    ' 対象のシート名に基づいて処理を行う
    For i = 1 To targetSheets.count
        sheetName = targetSheets(i)
        
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
                FormatRange wsTarget.Range("E13:E15"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If

            ' E17に値がない場合、A19:E19とB20:D21をグレーアウト
            If IsEmpty(wsTarget.Range("E17").value) Then
                wsTarget.Range("A17").value = "検査対象外"
                FormatRange wsTarget.Range("A17"), "游ゴシック", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B17:F17, B18:E19"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A17"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("E17:E19"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If
            
            ' 特定の文字に書式を適用
            FormatSpecificEndStrings wsTarget.Range("A10"), "游ゴシック", 12, True
            
            ' セルの書式設定
            With wsTarget.Range("C2:C4, F2:F4, H2:H4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("F3").NumberFormat = "0.0"" g"""
            wsTarget.Range("H2").NumberFormat = "0"" ℃"""
            wsTarget.Range("H3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("E11, E14, E19").NumberFormat = "0.00"" kN"""
            
            ' E14:E15, E18:E19の値に応じて書式を設定
            Set rng = wsTarget.Range("E14:E15, E18:E19")
            For Each cell In rng
                If cell.value <= 0.01 Then
                    cell.value = "―"
                Else
                    cell.NumberFormat = "0.00"" ms"""
                End If
            Next cell
            
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
                    .Name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "浸せき" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .Name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub



