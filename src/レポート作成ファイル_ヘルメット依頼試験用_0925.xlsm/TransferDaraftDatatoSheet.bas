Attribute VB_Name = "TransferDaraftDatatoSheet"
' "Impact_Top"などのシートを作成し、値を転記するプロシージャ
Sub TransferDataBasedOnID()
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim preProcess As String
    Dim topGap As String
    Dim testPoint As String
    Dim sampleName As String
    Dim MaxValue As Double, duration49kN As Double, duration73kN As Double
    Dim tempArray As Variant
    Dim data As Collection
    
    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row
    
    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割し、必要な情報を取得
        idParts = Split(wsSource.Cells(i, "B").value, "-")
        If UBound(idParts) >= 3 Then
            group = idParts(2)
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
        MaxValue = wsSource.Range("H" & i).value
        duration49kN = wsSource.Range("J" & i).value
        duration73kN = wsSource.Range("K" & i).value
        preProcess = wsSource.Range("L" & i).value
        topGap = wsSource.Range("N" & i).value
        testPoint = wsSource.Range("E" & i).value
        sampleName = wsSource.Range("D" & i).value
        tempArray = Array( _
            idParts(0), _
            targetSheetName, _
            MaxValue, _
            duration49kN, _
            duration73kN, _
            preProcess, _
            topGap, _
            testPoint, _
            sampleName _
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
            GetTargetSheetName = "Impact_Top"
        Case "前"
            GetTargetSheetName = "Impact_Front"
        Case "後"
            GetTargetSheetName = "Impact_Back"
        Case Else
            GetTargetSheetName = ""
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
        Dim MaxValue As Double, duration49kN As Double, duration73kN As Double
        Dim sampleName As String
        
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        MaxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)
        preProcess = dataItem(5)
        topGap = dataItem(6)
        testPoint = dataItem(7)
        sampleName = dataItem(8)
    
        ' 目的のシートを取得または作成
        Set wsDest = GetOrCreateSheet(targetSheetName)
        
        ' ヘッダー行を設定（14行目）
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "帽体No."
            wsDest.Range("C14").value = "前処理"
            wsDest.Range("D14").value = "試験位置"
            wsDest.Range("E14").value = "MAX"
            wsDest.Range("F14").value = "天頂すきま"
            wsDest.Range("G14").value = "4.9kN"
            wsDest.Range("H14").value = "7.3kN"
        End If
        
        ' 次の空行を取得しデータを転記
        nextRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = sampleName
        wsDest.Range("C" & nextRow).value = preProcess
        wsDest.Range("D" & nextRow).value = testPoint
        wsDest.Range("E" & nextRow).value = MaxValue
        wsDest.Range("F" & nextRow).value = topGap
        wsDest.Range("G" & nextRow).value = duration49kN
        wsDest.Range("H" & nextRow).value = duration73kN
    Next dataItem
End Sub

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
' TransferDataToSheetsのサブ関数。指定されたシートを取得または作成
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetOrCreateSheet.Name = sheetName
    End If
    GetOrCreateSheet.Visible = xlSheetVisible
    On Error GoTo 0
End Function

' "試験結果"シートからテンプレートをコピーしてくるプロシージャ
Sub ProcessImpactSheets()
    ' 変数の宣言
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim count As Integer
    Dim insertRows As Long
    Dim startRow As Long ' 探索開始行
    Dim currentGroup As Variant ' 現在のグループの値
    Dim previousGroup As Variant ' 前のグループの値
    Dim groupCount As Long ' グループの個数をカウントする変数
    Dim groupValues As Object ' 各グループの値とカウントを格納するDictionary
    Dim insertRowOffset As Long ' 挿入行のオフセット

    ' "試験結果"シートを変数に格納
    Set wsResult = Sheets("試験結果")

    ' シート名に"Impact"を含むシートをループ処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
        
            ' シートをアクティブにする
            ws.Activate

            ' A列の"Group"から始まる行を探索開始行とする
            startRow = Application.WorksheetFunction.Match("Group", ws.Range("A:A"), 0)

            ' A列の"Group"から始まる行から最終行までループ処理
            count = 0 ' ループの最初に count をリセット
            previousGroup = "" ' 最初のループでは前のグループは存在しないため空文字列を設定
            groupCount = 0 ' グループの個数を初期化
            Set groupValues = CreateObject("Scripting.Dictionary") ' Dictionaryオブジェクトを作成
            insertRowOffset = 0
            insertCount = 0
            
            For i = startRow To ws.Cells(ws.Rows.count, "A").End(xlUp).row
                currentGroup = ws.Cells(i, "A").value

                ' 現在のグループと前のグループが同じ場合、カウントを増やす
                If currentGroup = previousGroup Then
                    count = count + 1
                Else
                    ' 現在のグループと前のグループが異なる場合、カウントをリセット
                    ' 新しいグループが開始されたため、前のグループの値とカウントをDictionaryに格納
                    If previousGroup <> "" Then
                        If Not groupValues.Exists(previousGroup) Then
                            groupValues.Add previousGroup, count
                        Else
                            groupValues(previousGroup) = groupValues(previousGroup) + count ' 値を更新する
                        End If

                        If groupCount > 0 Then '2回目以降のグループの場合
                            insertRowOffset = insertRowOffset + insertRows
                        End If
                    End If

                    count = 1
                    groupCount = groupCount + 1 ' 新しいグループが開始されたため、グループの個数を増やす

                End If

                ' 前のグループを更新
                previousGroup = currentGroup
            Next i

            ' 最後のグループの値とカウントをDictionaryに格納
            If Not groupValues.Exists(previousGroup) Then
                groupValues.Add previousGroup, count
            Else
                groupValues(previousGroup) = groupValues(previousGroup) + count
            End If

'            ' グループの数と各グループのカウントを出力
'            Debug.Print "シート: " & ws.Name
'            Debug.Print "グループの数: " & groupCount
            For Each key In groupValues.Keys
'                Debug.Print "グループ " & key & ": " & groupValues(key) & " 個"
'                Debug.Print "挿入位置: : " & insertRowOffset

                ' グループのカウントに基づいて挿入とコピーを行う
                Select Case groupValues(key)
                    Case 3
                        insertRows = wsResult.Range("A3:A5").Rows.count
                        ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                        With wsResult
                            .Range(.Cells(3, "A"), .Cells(5, "G")).Copy
                        End With
                        ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                        ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & key

                    Case 2
                        insertRows = wsResult.Range("A7:A9").Rows.count
                        ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                        With wsResult
                            .Range(.Cells(7, "A"), .Cells(9, "G")).Copy
                        End With
                        ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                        ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & key
                End Select
                
                ' 各グループの処理後に insertRowOffset を更新
                insertRowOffset = insertRowOffset + insertRows
            Next key

            ' ループの最後に各カウンタをリセット
            count = 0
            insertRows = 0
            insertRowOffset = 0
        End If
    Next ws
    Call SetCellDimensions
End Sub

Sub SetCellDimensions()
    ' ProcessImpactSheetsのサブルーチン。シート名に"Impact"を含むシートをループ処理
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' シートをアクティブにする
            ws.Activate

            ' I列の"Insert" + 数字が入っている行をループ処理
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
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
    ' A2:G4セル範囲を取得 (startRowとendRowに合わせて調整)
    Dim targetRange As Range
    Set targetRange = ws.Range("A" & startRow & ":G" & endRow)

    ' A列の幅を設定
    targetRange.Columns(1).ColumnWidth = 2.8 ' ピクセルをポイントに変換

    ' B列からG列の幅を設定
    targetRange.Columns(2).Resize(1, 6).ColumnWidth = 11.5 ' ピクセルをポイントに変換

    ' 1行目と3行目の高さを設定 (startRowに合わせて調整)
    targetRange.Rows(1).RowHeight = 18
    targetRange.Rows(3).RowHeight = 18

    ' 2行目の高さを設定 (startRowに合わせて調整)
    targetRange.Rows(2).RowHeight = 161
End Sub





