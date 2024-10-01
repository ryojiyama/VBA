Attribute VB_Name = "SpecSheet"
 ' ☆品番、試験箇所などに応じたIDを作成する
Sub createID()

    Dim ws As Worksheet
    Dim i As Long
    Dim id As String
    Dim lastRow As Long

    ' "Bicycle_SpecSheet" を含むシートを順次処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Bicycle_SpecSheet", vbTextCompare) > 0 Then
            
            ' 最後の行を取得
            lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
            
            ' 各行に対してIDを生成
            For i = 2 To lastRow ' 1行目はヘッダと仮定
                id = GenerateID(ws, i)
                ' B列にIDをセット
                ws.Cells(i, 2).value = id
            Next i
            
        End If
    Next ws

End Sub



Function GenerateID(ws As Worksheet, rowIndex As Long) As String
    ' CreateID()のサブプロシージャ
    Dim id As String

    ' C列: 2桁以下の数字
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C列とD列の間に"-"
    
    ' D列の処理を変更
    id = id & ExtractNumber(ws.Cells(rowIndex, 4).value)
    id = id & "-" ' D列とE列の間に"-"
    
    ' N列（14列目）の条件
    id = id & GetColumnNValue(ws.Cells(rowIndex, 14).value)
    id = id & "-" ' E列とM列の間に"-"
    
    ' M列（13列目）の条件
    id = id & GetColumnMValue(ws.Cells(rowIndex, 13).value)
    id = id & "-" ' M列とO列の間に"-"
    
    ' O列（15列目）の条件
    id = id & GetColumnOValue(ws.Cells(rowIndex, 15).value)
    id = id & "-" ' O列とP列の間に"-"
    
    ' P列（16列目）の条件
    id = id & GetColumnPValue(ws.Cells(rowIndex, 16).value)
    
    ' 完成したIDを返す
    GenerateID = id
End Function

Function ExtractNumber(value As String) As String
    ' value を文字列に変換して返す
    ExtractNumber = CStr(value)
End Function

Function GetColumnCValue(value As Variant) As String
    ' GenerateIDのサブ関数
    If Len(value) <= 2 Then
        GetColumnCValue = Right("00" & value, 2)
    Else
        GetColumnCValue = "??"
    End If
End Function

Function GetColumnNValue(value As Variant) As String
    ' Valueに "前頭部" が含まれている場合は "前" を返す
    If InStr(value, "前頭部") > 0 Then
        GetColumnNValue = "前"
    ' Valueに "後頭部" が含まれている場合は "後" を返す
    ElseIf InStr(value, "後頭部") > 0 Then
        GetColumnNValue = "後"
    ' Valueに "左側頭部" が含まれている場合は "左" を返す
    ElseIf InStr(value, "左側頭部") > 0 Then
        GetColumnNValue = "左"
    ' Valueに "右側頭部" が含まれている場合は "右" を返す
    ElseIf InStr(value, "右側頭部") > 0 Then
        GetColumnNValue = "右"
    ' それ以外の場合は "?" を返す
    Else
        GetColumnNValue = "??"
    End If
End Function


Function GetColumnMValue(value As Variant) As String
    ' GenerateIDのサブ関数
    Select Case value
        Case "高温"
            GetColumnMValue = "Hot"
        Case "低温"
            GetColumnMValue = "Cold"
        Case "浸せき"
            GetColumnMValue = "Wet"
        Case Else
            GetColumnMValue = "?"
    End Select
End Function


Function GetColumnOValue(value As Variant) As String
    ' Valueが "平" の場合は "平" を返す
    If value = "平" Then
        GetColumnOValue = "平"
    ' Valueが "球" の場合は "球" を返す
    ElseIf value = "球" Then
        GetColumnOValue = "球"
    ' それ以外の場合は "その他" を返す
    Else
        GetColumnOValue = "その他"
    End If
End Function


Function GetColumnPValue(value As Variant) As String
    ' Valueが "A", "E", "J", "M", "O" の場合はそのまま返す
    If value = "A" Then
        GetColumnPValue = "A"
    ElseIf value = "E" Then
        GetColumnPValue = "E"
    ElseIf value = "J" Then
        GetColumnPValue = "J"
    ElseIf value = "M" Then
        GetColumnPValue = "M"
    ElseIf value = "O" Then
        GetColumnPValue = "O"
    ' それ以外の場合は "その他" を返す
    Else
        GetColumnPValue = "その他"
    End If
End Function




' ☆SpecSheetに転記するプロシージャの本体。
Sub SyncSpecSheetToLogHel()
    ' 同値が見つかった場合はエラーメッセージを表示して処理を中断
    If HighlightDuplicateValues Then
        MsgBox "衝撃値で同値が見つかりました。小数点下二桁に影響が出ない範囲で修正してください。", vbCritical
        Exit Sub
    End If

    ' 表に空欄がある場合にエラーメッセージを出して中断
    If Not LocateEmptySpaces Then
        MsgBox "空欄があります。まずはそれを埋めてください。", vbCritical
        Exit Sub
    End If
    Call createID              ' B列にIDを作成する。
    Call ProcessSheetPairs          ' 転記処理をするプロシージャ

End Sub
Function HighlightDuplicateValues() As Boolean
    ' SyncSpecSheetToLogHelのサブプロシージャ
    Dim sheetName As String
    sheetName = "Bicycle_SpecSheet"

    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundDuplicate As Boolean
    foundDuplicate = False ' 同値が見つかったかどうかのフラグを初期化

    ' シートオブジェクトを設定
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row

    ' 以前の色をクリア
    For i = 2 To lastRow
        ws.Cells(i, "J").Interior.colorIndex = xlNone
        ws.Cells(i, "K").Interior.colorIndex = xlNone
    Next i

    ' 色のインデックスを初期化
    Dim colorIndex As Integer
    colorIndex = 3 ' Excelの色インデックスは3から始まる

    ' J+K列の2行目から最終行までループ
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            ' J列とK列の値を組み合わせて比較
            If ws.Cells(i, "J").value & ws.Cells(i, "K").value = ws.Cells(j, "J").value & ws.Cells(j, "K").value And ws.Cells(i, "J").value <> "" And ws.Cells(i, "K").value <> "" Then
                ' 同値を持つセルが見つかった場合、フラグをTrueに設定し、セルに色を塗る
                foundDuplicate = True
                ws.Cells(i, "J").Interior.colorIndex = colorIndex
                ws.Cells(j, "J").Interior.colorIndex = colorIndex
                ws.Cells(i, "K").Interior.colorIndex = colorIndex
                ws.Cells(j, "K").Interior.colorIndex = colorIndex
            End If
        Next j
        ' 同値が見つかった場合、次の色に変更
        If foundDuplicate And ws.Cells(i, "J").Interior.colorIndex <> xlNone Then
            colorIndex = colorIndex + 1
            ' 色インデックスの最大値を超えないようにチェック
            If colorIndex > 56 Then colorIndex = 3 ' 色インデックスをリセット
        End If
    Next i

    ' 同値が一つも見つからなかった場合、J列とK列のセルの色をクリア
    If Not foundDuplicate Then
        For i = 2 To lastRow
            ws.Cells(i, "J").Interior.Color = xlNone
            ws.Cells(i, "K").Interior.Color = xlNone
        Next i
    End If

    ' 同値が見つかったかどうかに基づいて結果を返す
    HighlightDuplicateValues = foundDuplicate
End Function


Function LocateEmptySpaces() As Boolean
    ' SyncSpecSheetToLogHelのサブプロシージャ
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim cell As Range
    Dim errorMsg As String

    ' エラーメッセージ用の文字列を初期化
    errorMsg = ""

    ' シートオブジェクトを設定
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row

    ' 最終列を"M"(試験区分)に固定
    Dim lastCol As Long
    lastCol = ws.Columns("M").column

    ' 指定範囲をループ
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' 空白のチェック
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "空白セル: " & cell.Address(False, False) & vbNewLine
            End If

            ' 列G、H、J、Kで数値の確認
            If j = Columns("G").column Or j = Columns("H").column Or j = Columns("J").column Or j = Columns("K").column Then
                If Not IsNumeric(cell.value) Then
                    ' セルの書式を標準に設定
                    cell.NumberFormat = "General"

                    ' 数値に変換
                    If IsNumeric(cell.value) Then
                        cell.value = CDbl(cell.value)
                    Else
                        cell.value = 0
                    End If
                    cell.Interior.colorIndex = 6 ' 黄色に色付け
                    errorMsg = errorMsg & "数値に変換したセル: " & cell.Address(False, False) & vbNewLine
                End If
            End If

            ' 列N、O、Pで文字列の確認
            If j = Columns("N").column Or j = Columns("O").column Or j = Columns("P").column Then
                If Not VarType(cell.value) = vbString Then
                    ' 文字列に変換
                    cell.value = CStr(cell.value)
                    cell.Interior.colorIndex = 6 ' 黄色に色付け
                    errorMsg = errorMsg & "文字列に変換したセル: " & cell.Address(False, False) & vbNewLine
                End If
            End If
        Next j
    Next i

    ' エラーメッセージがあれば表示し、Falseを返す
    If Len(errorMsg) > 0 Then
        LocateEmptySpaces = False
        MsgBox errorMsg, vbCritical
    Else
        LocateEmptySpaces = True
    End If
End Function

' 転記処理をするプロシージャ
Sub ProcessSheetPairs()
    Dim sheetPairs As Variant
    Dim logSheetName As String
    Dim specSheetName As String
    Dim pair As Variant

    ' シートペアを定義
    sheetPairs = Array( _
        Array("LOG_Helmet", "Hel_SpecSheet"), _
        Array("LOG_FallArrest", "FallArr_SpecSheet"), _
        Array("LOG_Bicycle", "Bicycle_SpecSheet"), _
        Array("LOG_BaseBall", "Base_SpecSheet") _
    )

    ' 各シートペアを探索して処理
    For Each pair In sheetPairs
        logSheetName = pair(0)
        specSheetName = pair(1)
'        Debug.Print logSheetName
'        Debug.Print specSheetName
        ' シートペアが存在するかチェック
        If SheetExists(logSheetName) And SheetExists(specSheetName) Then
            ' シートペアが成立した場合に処理を実行
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
        End If
    Next pair
End Sub

Function SheetExists(sheetName As String) As Boolean
    'ProcessSheetPairsのサブプロシージャ
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


Sub CopyDataBasedOnCondition(sheetNameLog As String, sheetNameSpec As String)
    'ProcessSheetPairsのサブプロシージャ
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
    Dim structureCol As Long
    Dim penetrationCol As Long
    Dim zairyoCol As Long
    Dim logSum As Double
    Dim specSum As Double

    ' ワークシートをセット
    Set logSheet = ThisWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ThisWorkbook.Worksheets(sheetNameSpec)

    ' LOGシートの最終行を取得
    lastRowLog = logSheet.Cells(logSheet.Rows.count, "J").End(xlUp).row
    ' Specシートの最終行を取得
    lastRowSpec = helSpec.Cells(helSpec.Rows.count, "J").End(xlUp).row

    ' ヘッダー行を取得
    Set logHeader = logSheet.Rows(1)
    Set helSpecHeader = helSpec.Rows(1)

    ' 転記する列のペアをコレクションに定義
    Set columnsToCopy = New Collection

    ' ペアとなるヘッダー名を取得
    colPair = GetHeaderPairs(sheetNameLog, sheetNameSpec)

    ' ペアが正しく取得されているか確認
    If UBound(colPair) = -1 Then
        MsgBox "ヘッダーのペアが見つかりませんでした: " & sheetNameLog & " と " & sheetNameSpec
        Exit Sub
    End If

    ' 各ヘッダー行を走査して一致するヘッダーを見つける
    Dim pair As Variant
    For Each pair In colPair
        Dim logCol As Long
        Dim helSpecCol As Long
        logCol = 0
        helSpecCol = 0
        For Each col In logHeader.Cells
            If col.value = pair(0) Then
                logCol = col.column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.value = pair(1) Then
                helSpecCol = col.column
                Exit For
            End If
        Next col
        If logCol > 0 And helSpecCol > 0 Then
            columnsToCopy.Add Array(logCol, helSpecCol)
        Else
            MsgBox "ヘッダーが見つかりませんでした: " & pair(0) & " または " & pair(1)
        End If
    Next pair

    ' 値を比較して転記
    For i = 2 To lastRowLog
        matchCount = 0
        logSum = logSheet.Cells(i, "J").value + logSheet.Cells(i, "K").value

        For j = 2 To lastRowSpec
            specSum = helSpec.Cells(j, "J").value + helSpec.Cells(j, "K").value

            ' J+Kの合計が一致する場合に転記処理を行う
            If logSum = specSum Then
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.count
                    logSheet.Cells(i, columnsToCopy(k)(0)).value = helSpec.Cells(j, columnsToCopy(k)(1)).value
                Next k
            End If
        Next j

        ' 一致した値が複数存在する場合、文字を太字にする
        If matchCount > 1 Then
            Dim l As Long
            For l = 1 To columnsToCopy.count
                logSheet.Cells(i, columnsToCopy(l)(0)).Font.Bold = True
            Next l
        End If
    Next i

    ' 追加機能: 「構造_検査結果」と「耐貫通_検査結果」の列に「合格」を入力
    structureCol = FindHeaderColumn(logHeader, "外観検査")
    penetrationCol = FindHeaderColumn(logHeader, "あごひも検査")
    zairyoCol = FindHeaderColumn(logHeader, "材料・付属品検査")

    If structureCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, structureCol).value = "合格"
        Next i
    Else
        MsgBox "ヘッダー「構造_検査結果」が見つかりませんでした。"
    End If

    If penetrationCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, penetrationCol).value = "合格"
        Next i
    Else
        MsgBox "ヘッダー「耐貫通_検査結果」が見つかりませんでした。"
    End If
    
    If zairyoCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, zairyoCol).value = "合格"
        Next i
    Else
        MsgBox "ヘッダー「材料・付属品検査」が見つかりませんでした。"
    End If

    ' 転記が行われたことを確認
    MsgBox "転記が完了しました: " & sheetNameLog & " から " & sheetNameSpec
End Sub


' 指定したヘッダー名を持つ列の番号を取得する関数
Function FindHeaderColumn(headerRow As Range, headerName As String) As Long
    Dim col As Range
    For Each col In headerRow.Cells
        If col.value = headerName Then
            FindHeaderColumn = col.column
            Exit Function
        End If
    Next col
    FindHeaderColumn = -1 ' ヘッダーが見つからなかった場合
End Function

Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    'ProcessSheetPairsのサブプロシージャ
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
            headerPairs = Array( _
                Array("試料ID", "試験ID(C)"), _
                Array("品番", "品番(D)"), _
                Array("試験内容", "試験位置(E)"), _
                Array("検査日", "検査日(F)"), _
                Array("温度", "温度(G)"), _
                Array("最大値(kN)", "衝撃値(H)"), _
                Array("前処理", "前処理(L)"), _
                Array("重量", "重量(M)"), _
                Array("天頂すきま", "天頂すきま(N)"), _
                Array("帽体色", "帽体色(O)"), _
                Array("ロットNo.", "製造ロット(P)"), _
                Array("帽体ロット", "帽体ロット(Q)"), _
                Array("内装ロット", "内装ロット(R)"), _
                Array("構造_検査結果", "構造/結果(S)"), _
                Array("耐貫通_検査結果", "耐貫通/結果(U)"), _
                Array("試験区分", "試験内容(U)") _
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
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bicycle_SpecSheet" Then
        headerPairs = Array( _
            Array("ID", "試験ID(C)"), _
            Array("試料ID", "試料ID"), _
            Array("品番", "品番"), _
            Array("ロット番号", "ロット番号"), _
            Array("試験日", "試験日"), _
            Array("温度", "温度"), _
            Array("湿度", "湿度"), _
            Array("重量", "重量"), _
            Array("前処理", "前処理"), _
            Array("試験箇所", "試験箇所"), _
            Array("アンビル", "アンビル"), _
            Array("人頭模型", "人頭模型") _
        )
    ElseIf sheetNameLog = "LOG_BaseBall" And sheetNameSpec = "Base_SpecSheet" Then
        headerPairs = Array( _
            Array("別の最大値", "別の衝撃値"), _
            Array("別のDヘッダー名", "別のDヘッダー名"), _
            Array("別のUヘッダー名", "別のMヘッダー名") _
        )
    Else
        headerPairs = Array()
    End If

    GetHeaderPairs = headerPairs
End Function










