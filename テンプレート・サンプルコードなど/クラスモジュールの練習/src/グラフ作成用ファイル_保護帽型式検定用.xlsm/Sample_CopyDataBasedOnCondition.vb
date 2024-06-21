Sub ProcessSheetPairs()
    Dim sheetPairs As Variant
    Dim logSheetName As String
    Dim specSheetName As String
    Dim pair As Variant

    ' シートペアを定義
    sheetPairs = Array( _
        Array("LOG_Helmet", "Hel_SpecSheet"), _
        Array("LOG_FallArrest", "FallArr_SpecSheet"), _
        Array("LOG_Bicycle", "Bic_SpecSheet"), _
        Array("LOG_BaseBall", "Base_SpecSheet") _
    )

    ' 各シートペアを探索して処理
    For Each pair In sheetPairs
        logSheetName = pair(0)
        specSheetName = pair(1)

        ' シートペアが存在するかチェック
        If SheetExists(logSheetName) And SheetExists(specSheetName) Then
            ' シートペアが成立した場合に処理を実行
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
        End If
    Next pair
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Sub CopyDataBasedOnCondition(sheetNameLog As String, sheetNameSpec As String)
    Dim logSheet As Worksheet
    Dim helSpec As Worksheet
    Dim lastRowLog As Long
    Dim lastRowSpec As Long
    Dim i As Long, j As Long
    Dim matchCount As Long
    Dim columnsToCopy As Collection
    Dim colPair As Variant
    Dim logHeader As Range
    Dim helSpecHeader As Range
    Dim col As Range
    Dim colLog As Range

    ' ワークシートをセット
    Set logSheet = ThisWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ThisWorkbook.Worksheets(sheetNameSpec)

    ' LOGシートの最終行を取得
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).Row
    ' Specシートの最終行を取得
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).Row

    ' ヘッダー行を取得
    Set logHeader = logSheet.Rows(1)
    Set helSpecHeader = helSpec.Rows(1)

    ' 転記する列のペアをコレクションに定義
    Set columnsToCopy = New Collection

    ' ペアとなるヘッダー名を取得
    colPair = GetHeaderPairs(sheetNameLog, sheetNameSpec)

    ' 各ヘッダー行を走査して一致するヘッダーを見つける
    Dim pair As Variant
    For Each pair In colPair
        Dim logCol As Long
        Dim helSpecCol As Long
        logCol = 0
        helSpecCol = 0
        For Each col In logHeader.Cells
            If col.Value = pair(0) Then
                logCol = col.Column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.Value = pair(1) Then
                helSpecCol = col.Column
                Exit For
            End If
        Next col
        If logCol > 0 And helSpecCol > 0 Then
            columnsToCopy.Add Array(logCol, helSpecCol)
        End If
    Next pair

    ' 値を比較して転記
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").Value = helSpec.Cells(j, "H").Value Then
                ' H列の値が一致した場合、各列の内容を転記
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.Count
                    logSheet.Cells(i, columnsToCopy(k)(0)).Value = helSpec.Cells(j, columnsToCopy(k)(1)).Value
                Next k
            End If
        Next j

        ' 一致した値が複数存在する場合、文字を太字にする
        If matchCount > 1 Then
            Dim l As Long
            For l = 1 To columnsToCopy.Count
                logSheet.Cells(i, columnsToCopy(l)(0)).Font.Bold = True
            Next l
        End If
    Next i
End Sub

Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
        headerPairs = Array( _
            Array("試験ID(C)", "衝撃値(H)"), _
            Array("品番(D)", "品番"), _
            Array("試験内容(E)", "試験内容"), _
            Array("検査日(F)", "検査日"), _
            Array("温度(G)", "温度"), _
            Array("前処理(L)", "前処理"), _
            Array("重量(M)", "重量"), _
            Array("天頂すきま(N)", "天頂すきま"), _
            Array("帽体色(O)", "帽体色") _
            Array("試験区分(U)", "試験区分") _
        )
    ElseIf sheetNameLog = "LOG_FallArrest" And sheetNameSpec = "FallArr_SpecSheet" Then
        headerPairs = Array( _
            Array("別の最大値", "別の衝撃値"), _
            Array("別のDヘッダー名", "別のDヘッダー名"), _
            Array("別のEヘッダー名", "別のEヘッダー名"), _
            Array("別のFヘッダー名", "別のFヘッダー名"), _
            Array("別のGヘッダー名", "別のGヘッダー名"), _
            Array("別のLヘッダー名", "別のIヘッダー名"), _
            Array("別のMヘッダー名", "別のJヘッダー名"), _
            Array("別のNヘッダー名", "別のKヘッダー名"), _
            Array("別のOヘッダー名", "別のLヘッダー名"), _
            Array("別のUヘッダー名", "別のMヘッダー名") _
        )
    ' 必要に応じて他のシートペアを追加
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bic_SpecSheet" Then
        headerPairs = Array( _
            Array("値1", "値2"), _
            Array("ヘッダー1", "ヘッダー2") _
            ' 他のペアを追加
        )
    ElseIf sheetNameLog = "LOG_BaseBall" And sheetNameSpec = "Base_SpecSheet" Then
        headerPairs = Array( _
            Array("値A", "値B"), _
            Array("ヘッダーA", "ヘッダーB") _
            ' 他のペアを追加
        )
    Else
        headerPairs = Array()
    End If

    GetHeaderPairs = headerPairs
End Function
