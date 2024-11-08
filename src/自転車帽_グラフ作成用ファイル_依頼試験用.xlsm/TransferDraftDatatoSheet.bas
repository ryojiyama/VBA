Attribute VB_Name = "TransferDraftDatatoSheet"
' "レポートグラフ"などのシートを作成し、値を転記するプロシージャ
Sub TransferDataBasedOnID()
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim preProcess As String ' 前処理
    Dim anvilType As String ' アンビル
    Dim dummyHead As String '人頭模型
    Dim testPoint As String '試験箇所
    Dim sampleName As String
    Dim maxValue As Double
    Dim tempArray As Variant
    Dim data As Collection
    
    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
    
    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割し、必要な情報を取得
        idParts = Split(wsSource.Cells(i, "B").value, "-")
        If UBound(idParts) >= 3 Then
            group = idParts(0)
        Else
            ' ID形式が不正な場合は次のループへ
            GoTo NextIteration
        End If
        
        ' シート名をグループに基づいて決定
        targetSheetName = GetTargetSheetName(group)
        If targetSheetName = "" Then
            GoTo NextIteration
        End If
        
        ' データをコレクションに追加
        maxValue = wsSource.Range("J" & i).value
        sampleName = wsSource.Range("D" & i).value
        testPoint = wsSource.Range("N" & i).value
        dummyHead = wsSource.Range("P" & i).value
        anvilType = wsSource.Range("O" & i).value
        preProcess = wsSource.Range("M" & i).value
        tempArray = Array( _
            idParts(0), _
            targetSheetName, _
            maxValue, _
            sampleName, _
            testPoint, _
            dummyHead, _
            anvilType, _
            preProcess _
        )
        data.Add tempArray
        
NextIteration:
    Next i
    
    ' データを各シートに転記
    TransferDataToSheets data
    
    ' リソースを解放
    Set wsSource = Nothing
    Set data = Nothing
End Sub

Function GetTargetSheetName(ByVal group As String) As String
'TransferDataBasedOnIDのサブ関数。グループに基づいてターゲットシート名を取得する
    Select Case group
        Case "天"
            GetTargetSheetName = "Sub1"
        Case "前"
            GetTargetSheetName = "Sub2"
        Case "後"
            GetTargetSheetName = "Sub3"
        Case Else
            GetTargetSheetName = "レポートグラフ"
    End Select
End Function

Sub TransferDataToSheets(ByVal data As Collection)
' データを各シートに転記するTransferDataBasedOnIDのサブプロシージャ
    Dim wsDest As Worksheet
    Dim dataItem As Variant
    Dim nextRow As Long
    
    For Each dataItem In data
        ' 変数の割り当て
        Dim groupName As String
        Dim targetSheetName As String
        Dim preProcess As String
        Dim topGap As String
        Dim testPoint As String
        Dim maxValue As Double, duration49kN As Double, duration73kN As Double
        Dim sampleName As String
        
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        sampleName = dataItem(3)
        testPoint = dataItem(4)
        dummyHead = dataItem(5)
        anvilType = dataItem(6)
        preProcess = dataItem(7)
    
        ' 目的のシートを取得または作成
        Set wsDest = GetOrCreateSheet(targetSheetName)
        
        ' ヘッダー行を設定（14行目）
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "最大値"
            wsDest.Range("C14").value = "帽体No."
            wsDest.Range("D14").value = "試験箇所"
            wsDest.Range("E14").value = "人頭模型"
            wsDest.Range("F14").value = "アンビル"
            wsDest.Range("G14").value = "前処理"
        End If
        
        ' 次の空行を取得しデータを転記
        nextRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = maxValue
        wsDest.Range("C" & nextRow).value = sampleName
        wsDest.Range("D" & nextRow).value = testPoint
        wsDest.Range("E" & nextRow).value = dummyHead
        wsDest.Range("F" & nextRow).value = anvilType
        wsDest.Range("G" & nextRow).value = preProcess
    Next dataItem
End Sub

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
' TransferDataToSheetsのサブ関数。指定されたシートを取得または作成
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
    GetOrCreateSheet.Visible = xlSheetVisible
    On Error GoTo 0
End Function


' "レポートグラフ"シートに"テンプレート"シートから行をコピーしてグループ分のみ挿入
Sub ProcessImpactSheets()
    ' 変数の宣言
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim startRow As Long ' 探索開始行
    Dim currentGroup As String ' 現在のグループの値
    Dim previousGroup As String ' 前のグループの値
    Dim groupCount As Long ' グループの個数をカウントする変数
    Dim groupSize As Long ' 各グループ内のデータ数
    Dim insertRowOffset As Long ' 挿入行のオフセット
    Dim lastRow As Long ' A列の最終行番号
    Dim rowsToInsert As Collection ' 挿入行番号のコレクション
    Dim groupIndices As Collection ' グループ番号のコレクション
    Dim groupIndex As Long ' グループ番号
    Dim insertRows As Long
    Dim row As Variant ' Collectionのループ用変数
    Dim index As Variant ' Collectionのグループ番号用変数
    Dim groupCounter As Long ' グループカウンター（挿入用）

    ' "試験結果"シートを変数に格納
    Set wsResult = Sheets("Bicycle_テンプレート")

    ' シート名に"レポートグラフ"を含むシートをループ処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "レポートグラフ") > 0 Then

            ' シートをアクティブにする
            ws.Activate

            ' A列の"Group"から始まる行を探索開始行とする
            startRow = Application.WorksheetFunction.Match("Group", ws.Range("A:A"), 0)

            ' startRowの値を確認
            Debug.Print "startRow (Groupの行): " & startRow

            ' A列の最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

            ' 探索範囲を確認
            Debug.Print "探索範囲: A" & startRow & ":A" & lastRow

            ' A列の"Group"以下のデータを昇順に並べ替える
            With ws.Sort
                .SortFields.Clear
                .SortFields.Add key:=ws.Range("A" & startRow + 1 & ":A" & lastRow), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SetRange ws.Range("A" & startRow & ":Z" & lastRow)
                .Header = xlYes ' ヘッダーがあることを指定
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With

            ' 挿入予定の行番号を保持するコレクション
            Set rowsToInsert = New Collection
            Set groupIndices = New Collection

            ' データを1行ずつ確認し、グループを判別
            previousGroup = "" ' 最初のループでは前のグループは存在しないため空文字列を設定
            groupCount = 0 ' グループのカウントを初期化
            groupIndex = 1 ' グループ番号の初期化
            groupSize = 0 ' 初期グループのサイズを初期化

            For i = startRow + 1 To lastRow
                ' Trimでスペースを除去して値を取得
                currentGroup = Trim(ws.Cells(i, "A").value)

                ' 空白セルを無視する
                If currentGroup <> "" Then
                    ' A列の値そのままを出力
                    Debug.Print "行番号: " & i & ", A列の値: '" & currentGroup & "'"

                    ' 同じグループかどうかを判定
                    If currentGroup = previousGroup Then
                        groupSize = groupSize + 1 ' 同じグループ内なのでカウントを増やす
                    Else
                        ' 前のグループの処理が完了した時点で、グループのサイズを確認
                        If groupSize > 0 Then
                            ' グループが4を超えている場合はエラーを出して無視
                            If groupSize > 4 Then
                                Debug.Print "エラー: グループ" & previousGroup & "が4を超えています。処理を無視します。"
                            Else
                                ' グループのサイズが4以下の場合は挿入対象
                                rowsToInsert.Add i - groupSize - 1 ' グループの開始行を保存
                                groupIndices.Add groupIndex ' グループ番号を保存
                            End If
                        End If
                        
                        ' 新しいグループに移るため、カウントをリセット
                        groupSize = 1
                        groupIndex = groupIndex + 1 ' 新しいグループ番号に進む
                    End If

                    ' 次のグループ判定のために現在のグループを保存
                    previousGroup = currentGroup
                End If
            Next i

            ' 最後のグループも確認
            If groupSize > 0 Then
                If groupSize > 4 Then
                    Debug.Print "エラー: グループ" & previousGroup & "が4を超えています。処理を無視します。"
                Else
                    rowsToInsert.Add lastRow - groupSize + 1 ' 最後のグループの開始行を保存
                    groupIndices.Add groupIndex ' 最後のグループ番号を保存
                End If
            End If

            ' 行挿入処理をまとめて実行
            insertRowOffset = 0
            groupCounter = 1 ' グループカウンターを初期化
            For Each row In rowsToInsert
                ' グループ番号を取得
                index = groupIndices.item(groupCounter)
                
                insertRows = wsResult.Range("A2:A7").Rows.Count
                ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                With wsResult
                    .Range(.Cells(2, "A"), .Cells(7, "G")).Copy
                End With
                ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                ' index を -1 して、正しい groupIndex を設定
                ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & (index - 1)

                ' 挿入行のオフセットを更新
                insertRowOffset = insertRowOffset + insertRows
                groupCounter = groupCounter + 1 ' グループカウンターを更新
            Next row

            ' ループの最後にカウンタをリセット
            groupCount = 0
            insertRows = 0
            insertRowOffset = 0
        End If
    Next ws
End Sub
' "レポートグラフ"シートの列/行のサイズを整える
Sub SetCellDimensions()
    ' ProcessImpactSheetsのサブルーチン。シート名に"Impact"を含むシートをループ処理
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "レポートグラフ") > 0 Then
            ' シートをアクティブにする
            ws.Activate

            ' I列の"Insert" + 数字が入っている行をループ処理
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
            Dim i As Long
            Dim currentInsertGroup As String
            Dim groupStartRow As Long

            For i = 2 To lastRow ' 2行目から開始
                ' I列の値が "Insert" で始まり、その後に数字が続く場合
                If ws.Cells(i, "I").value Like "Insert[0-9]*" Then
                    ' 新しいグループの場合
                    If ws.Cells(i, "I").value <> currentInsertGroup Then
                        ' 前のグループのセル高さ・幅を設定 (最初のグループ以外)
                        If currentInsertGroup <> "" Then
                            SetGroupCellDimensions ws, groupStartRow, i - 1
                        End If

                        ' 新しいグループの開始行を記録
                        currentInsertGroup = ws.Cells(i, "I").value
                        groupStartRow = i
                    End If
                End If
            Next i

            ' 最後のグループのセル高さ・幅を設定
            If currentInsertGroup <> "" Then
                SetGroupCellDimensions ws, groupStartRow, lastRow
            End If
        End If
    Next ws

End Sub

Sub SetGroupCellDimensions(ws As Worksheet, startRow As Long, endRow As Long)
    ' SetCellDimensionsのサブルーチン。グループのセル高さ・幅を設定する
    ' A列からG列の幅と各行の高さを指定された条件に合わせて設定
    Debug.Print "startRow: " & startRow
    ' A列の幅を列幅単位で指定
    ws.Columns(1).ColumnWidth = 2.3

    ' B列とE列の幅を列幅単位で指定
    ws.Columns(2).ColumnWidth = 11.8
    ws.Columns(5).ColumnWidth = 11.8

    ' C列とF列の幅を列幅単位で指定
    ws.Columns(3).ColumnWidth = 11
    ws.Columns(6).ColumnWidth = 11

    ' D列とG列の幅を列幅単位で指定
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(7).ColumnWidth = 16

    ' 行の高さをピクセル換算のポイントで設定
    Dim i As Long
    For i = startRow To endRow
        Select Case (i - startRow + 1) Mod 6
            Case 1, 2, 4, 5
                ws.Rows(i).RowHeight = 18
            Case 3, 6
                ws.Rows(i).RowHeight = 127.8
            Case 0
                ws.Rows(i).RowHeight = 127.8
        End Select
    Next i
End Sub
' "レポートグラフ"シートにヘッダーを追加する。(途中)
Sub AddHeaderToReportSheets()
    Dim ws As Worksheet
    Dim lastCol As Long

    ' 名前に"レポートグラフ"が含まれるシートを処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "レポートグラフ") > 0 Then
            ' 1行目に行を挿入
            ws.Rows(1).Insert Shift:=xlDown

            ' 最終列を取得
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            
            ' 真ん中の列を計算して、ダミーテキストを挿入
            ws.Cells(1, Application.RoundUp(lastCol / 2, 0)).value = "ダミーテキスト"
        End If
    Next ws
End Sub






