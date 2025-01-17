Attribute VB_Name = "SpecSheet"
 ' ☆品番、試験箇所などに応じたIDを作成する
Sub CreateID(sheetName As String)
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim id As String
    
    ' 引数で渡されたシート名を使用
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    
    ' 最後の行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' 各行に対してIDを生成
    For i = 2 To lastRow ' 1行目はヘッダと仮定
        id = GenerateID(ws, i)
        ' B列にIDをセット
        ws.Cells(i, 2).value = id
    Next i
End Sub
Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' CreateID()のサブプロシージャ
    Dim id As String

    ' C列: 2桁以下の数字
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C列とD列の間に"-"
    ' D列の処理を変更
    id = id & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    id = id & "-" ' FmとE列の間に"-"
    id = id & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E列の条件
    id = id & "-" ' FmとE列の間に"-
    id = id & GetColumnIValue(ws.Cells(rowIndex, 9).value) ' I列の条件
    id = id & "-" ' I列とL列の間に"-"
    id = id & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L列の条件

    GenerateID = id
End Function
Function ExtractNumberWithF(value As String) As String
' GenerateIDのサブ関数
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' 正規表現オブジェクトの作成
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{3,6}"
    regex.Global = True

    ' 数字部分を抽出
    Set matches = regex.Execute(value)
    If matches.Count > 0 Then
        numPart = matches(0).value
    Else
        numPart = "000000" ' デフォルト値またはエラーハンドリング
    End If

    ' Fの存在チェック
    hasF = InStr(value, "F") > 0

    ' Fがある場合は数字の後にFをつける
    If hasF Then
        ExtractNumberWithF = numPart & "F"
    Else
        ExtractNumberWithF = numPart
    End If
End Function
Function GetColumnCValue(value As Variant) As String
' GenerateIDのサブ関数
    If Len(value) <= 2 Then
        GetColumnCValue = Right("00" & value, 2)
    Else
        GetColumnCValue = "??"
    End If
End Function
Function GetColumnEValue(value As Variant) As String
    ' GenerateIDのサブ関数
    If InStr(value, "天頂") > 0 Then
        GetColumnEValue = "天"
    ElseIf InStr(value, "前頭部") > 0 Then
        GetColumnEValue = "前"
    ElseIf InStr(value, "後頭部") > 0 Then
        GetColumnEValue = "後"
    ElseIf InStr(value, "側面") > 0 Then
        Dim Parts() As String
        Parts = Split(value, "_")
        
        If UBound(Parts) >= 1 Then
            Dim angle As String
            Dim direction As String
            
            ' 角度を抽出
            angle = Replace(Parts(0), "側面", "")
            
            ' 方向を抽出と整形
            direction = Parts(1)
            direction = Replace(direction, "前", "前")
            direction = Replace(direction, "後", "後")
            direction = Replace(direction, "左", "左")
            direction = Replace(direction, "右", "右")
            
            GetColumnEValue = "側" & angle & direction
        Else
            GetColumnEValue = "側"
        End If
    Else
        GetColumnEValue = "?"
    End If
End Function
Function GetColumnIValue(value As Variant) As String
' GenerateIDのサブ関数
    Select Case value
        Case "高温"
            GetColumnIValue = "Hot"
        Case "低温"
            GetColumnIValue = "Cold"
        Case "浸せき"
            GetColumnIValue = "Wet"
        Case "常温"
            GetColumnIValue = "Nrml"
        Case Else
            GetColumnIValue = "?"
    End Select
End Function
Function GetColumnLValue(value As Variant) As String
' GenerateIDのサブ関数
    If value = "白" Then
        GetColumnLValue = "White"
    Else
        GetColumnLValue = "OthClr"
    End If
End Function
' 衝撃値のみをLOGシートに転記する。
Sub TransferValuesBetweenSheets()

    ' シートペアの配列を作成
    Dim sheetPairs As Variant
    sheetPairs = Array( _
        Array("Hel_SpecSheet", "LOG_Helmet"), _
        Array("Bicycle_SpecSheet", "LOG_Bicycle"), _
        Array("Fall_SpecSheet", "LOG_FallArrest"), _
        Array("BaseBall_SpecSheet", "LOG_BaseBall"))
    
    Dim i As Long
    Dim specSheet As Worksheet
    Dim logSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim cell As Range
    
    ' シートペアをループして処理
    For i = LBound(sheetPairs) To UBound(sheetPairs)
        ' SpecSheet と LOG_ シートを設定
        On Error Resume Next
        Set specSheet = ActiveWorkbook.Sheets(sheetPairs(i)(0))
        Set logSheet = ActiveWorkbook.Sheets(sheetPairs(i)(1))
        On Error GoTo 0
        
        ' シートが存在する場合に処理を実行
        If Not specSheet Is Nothing And Not logSheet Is Nothing Then
            ' SpecSheetのH列の最終行を取得
            lastRow = specSheet.Cells(specSheet.Rows.Count, "H").End(xlUp).row
            
            ' H列のデータを転記する範囲を設定
            Set dataRange = specSheet.Range("H2:H" & lastRow) ' H2から最終行まで
                       
            ' SpecSheetからLOG_シートへ値を転記
            logSheet.Range("H2").Resize(dataRange.Rows.Count).value = dataRange.value
        Else
            ' シートが存在しない場合のデバッグ出力
            Debug.Print "シートが見つかりませんでした: " & sheetPairs(i)(0) & " または " & sheetPairs(i)(1)
        End If
        
        ' オブジェクトのクリア
        Set specSheet = Nothing
        Set logSheet = Nothing
    Next i

End Sub





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

    Call UpdateCrownClearance ' 天頂すきまを調整
    Call ProcessSheetPairs   ' 転記処理をするプロシージャ

End Sub
Function HighlightDuplicateValues() As Boolean
    ' SyncSpecSheetToLogHelのサブプロシージャ
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundDuplicate As Boolean
    foundDuplicate = False ' 同値が見つかったかどうかのフラグを初期化

    ' シートオブジェクトを設定
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' 色のインデックスを初期化
    Dim colorIndex As Integer
    colorIndex = 3 ' Excelの色インデックスは3から始まる

    ' H列の2行目から最終行までループ
    For i = 2 To lastRow
        ' M列の値をチェックし、"依頼"が含まれる場合はフラグをFalseに設定
        If InStr(ws.Cells(i, "M").value, "依頼") > 0 Then
            foundDuplicate = False
        Else
            For j = i + 1 To lastRow
                If ws.Cells(i, "H").value = ws.Cells(j, "H").value And ws.Cells(i, "H").value <> "" Then
                    ' 同値を持つセルが見つかった場合、フラグをTrueに設定し、セルに色を塗る
                    foundDuplicate = True
                    ws.Cells(i, "H").Interior.colorIndex = colorIndex
                    ws.Cells(j, "H").Interior.colorIndex = colorIndex
                End If
            Next j
            ' 同値が見つかった場合、次の色に変更
            If foundDuplicate And ws.Cells(i, "H").Interior.colorIndex <> xlNone Then
                colorIndex = colorIndex + 1
                ' 色インデックスの最大値を超えないようにチェック
                If colorIndex > 56 Then colorIndex = 3 ' 色インデックスをリセット
            End If
        End If
    Next i

    ' 同値が一つも見つからなかった場合、H列のセルの色を白に設定
    If Not foundDuplicate Then
        For i = 2 To lastRow
            ws.Cells(i, "H").Interior.Color = xlNone
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
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' 最終列を"M"(試験区分)に固定
    Dim lastCol As Long
    lastCol = ws.columns("M").column

    ' 指定範囲をループ
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' 空白のチェック
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "空白セル: " & cell.Address(False, False) & vbNewLine
            End If

            ' 列G(温度)、H(衝撃値)、J(重量)、K(天頂すきま)で数値の確認
            If j = columns("G").column Or j = columns("H").column Or j = columns("J").column Or j = columns("K").column Then
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

            ' 列N(製造ロット)、O(帽体ロット)、P(内装ロット)で文字列の確認
            If j = columns("N").column Or j = columns("O").column Or j = columns("P").column Then
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
' 天頂すき間を"Setting"シートのデータに合わせて調整する。
Sub UpdateCrownClearance()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim colGenshoNoSukima As Integer
    Dim colKaisu As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim i As Long
    Dim skipCount As Long

    ' シートをセット
    Set wsHelSpec = ActiveWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ActiveWorkbook.Sheets("Setting")

    ' ヘッダーの列番号を取得
    colHinban = GetColumnIndex(wsHelSpec, "品番(D)")
    colBoutai = GetColumnIndex(wsSetting, "帽体No.")
    colTencho = GetColumnIndex(wsHelSpec, "天頂肉厚")
    colTenchoSukima = GetColumnIndex(wsHelSpec, "天頂すきま(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "測定すきま")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "天頂肉厚")
    colGenshoNoSukima = GetColumnIndex(wsHelSpec, "原初のすきま")
    colKaisu = GetColumnIndex(wsHelSpec, "回数")

    ' 必要な列が見つかったかを確認
    If colHinban = 0 Or colBoutai = 0 Or colTencho = 0 Or _
       colTenchoSukima = 0 Or colSokuteiSukima = 0 Or _
       colTenchoNikui = 0 Or colGenshoNoSukima = 0 Or _
       colKaisu = 0 Then
        MsgBox "必要な列が見つかりません。ヘッダーを確認してください。", vbCritical
        Exit Sub
    End If

    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row

    ' "品番(D)" 列の値を探索し、転記
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell

    skipCount = 0

    ' "天頂すきま(N)" の値を "測定すきま" にコピーし、値を計算
    For i = 2 To lastRowHelSpec
        ' 回数が記入済みの場合はスキップし、カウントを増やす
        If wsHelSpec.Cells(i, colKaisu).value <> "" Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' "原初のすき間" が空欄の場合のみコピー (回数列の状態に関わらず実行)
        If wsHelSpec.Cells(i, colGenshoNoSukima).value = "" Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = wsHelSpec.Cells(i, colTenchoSukima).value
            wsHelSpec.Cells(i, colGenshoNoSukima).value = wsHelSpec.Cells(i, colTenchoSukima).value
        End If

        ' 各セルの値を取得 (原初のすき間の値を取得)
        tenchoSukimaValue = wsHelSpec.Cells(i, colGenshoNoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value

        ' "原初のすき間"の値から"天頂肉厚"の値を引く
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If

        ' 回数に済を代入
        wsHelSpec.Cells(i, colKaisu).value = "済"

        ' Q列とR列に"合格"の値を代入
        wsHelSpec.Cells(i, 17).value = "合格" ' Q列は17番目の列
        wsHelSpec.Cells(i, 18).value = "合格" ' R列は18番目の列

NextRow:
    Next i

    ' メッセージの表示
    If skipCount > 0 Then
        MsgBox "修正はすでに行われました。（" & skipCount & "行スキップされました）"
    Else
        MsgBox "天頂すき間が正しいかチェックをお願いします。"
    End If

End Sub

' 列番号を取得する関数
Function GetColumnIndex(targetSheet As Worksheet, headerName As String) As Integer
    Dim headerRange As Range
    Set headerRange = targetSheet.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerRange Is Nothing Then
        GetColumnIndex = headerRange.column
    Else
        GetColumnIndex = 0
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
        Array("LOG_Bicycle", "Bic_SpecSheet"), _
        Array("LOG_BaseBall", "Base_SpecSheet") _
    )

    ' 各シートペアを探索して処理
    For Each pair In sheetPairs
        logSheetName = pair(0)
        specSheetName = pair(1)
'        Debug.Print logSheetName
'        Debug.Print specSheetName
        ' シートペアが存在するかチェック
        If sheetExists(logSheetName) And sheetExists(specSheetName) Then
            ' シートペアが成立した場合に処理を実行
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
            Debug.Print logSheetName
        End If
    Next pair
End Sub

Function sheetExists(sheetName As String) As Boolean
    'ProcessSheetPairsのサブプロシージャ
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    sheetExists = Not ws Is Nothing
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

    ' ワークシートをセット
    Set logSheet = ActiveWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ActiveWorkbook.Worksheets(sheetNameSpec)

    ' LOGシートの最終行を取得
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).row
    ' Specシートの最終行を取得
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).row

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
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").value = helSpec.Cells(j, "H").value Then
                ' H列の値が一致した場合、各列の内容を転記
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.Count
                    logSheet.Cells(i, columnsToCopy(k)(0)).value = helSpec.Cells(j, columnsToCopy(k)(1)).value
                Next k
                ' C列の値もB列にコピー
                logSheet.Cells(i, "B").value = logSheet.Cells(i, "C").value
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

    ' 追加機能: 「構造_検査結果」と「耐貫通_検査結果」の列に「合格」を入力
    structureCol = FindHeaderColumn(logHeader, "構造_検査結果")
    penetrationCol = FindHeaderColumn(logHeader, "耐貫通_検査結果")

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
                Array("試験区分", "試験内容(U)"), _
                Array("ストライカ高さ", "ストライカ高さ(V)"), _
                Array("内装種類", "内装種類(W)"), _
                Array("前処理時間", "前処理時間(X)"), _
                Array("備考", "備考(Z)") _
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
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bic_SpecSheet" Then
        headerPairs = Array( _
            Array("別の最大値", "別の衝撃値"), _
            Array("別のDヘッダー名", "別のDヘッダー名"), _
            Array("別のUヘッダー名", "別のMヘッダー名") _
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

Sub HighlightMismatchedRows()
    Dim sheetPairs As Variant
    Dim logSheet As Worksheet
    Dim specSheet As Worksheet
    Dim logLastRow As Long, specLastRow As Long
    Dim i As Long, j As Long
    Dim logValue1 As Variant, specValue1 As Variant
    Dim logValue2 As Variant, specValue2 As Variant
    Dim logValue3 As Variant, specValue3 As Variant
    Dim mismatchFound As Boolean ' 全体の不一致検知フラグ
    
    ' ペアごとの比較列の定義を別の配列で設定
    Dim helmetColumns As Variant
    Dim fallArrestColumns As Variant
    Dim bicycleColumns As Variant
    Dim baseBallColumns As Variant
    
    ' シートペアの定義
    sheetPairs = Array( _
        Array("LOG_Helmet", "Hel_SpecSheet"), _
        Array("LOG_FallArrest", "FallArr_SpecSheet"), _
        Array("LOG_Bicycle", "Bicycle_SpecSheet"), _
        Array("LOG_BaseBall", "BaseBall_SpecSheet") _
    )
    
    ' 各シートペアに対応する列の定義
    helmetColumns = Array("J", "I", "O", "J", "I", "O")
    fallArrestColumns = Array("K", "L", "M", "K", "L", "M")
    bicycleColumns = Array("J", "I", "O", "J", "I", "O")
    baseBallColumns = Array("N", "O", "P", "N", "O", "P")
    
    mismatchFound = False ' 初期化
    
    ' 各シートペアをループ
    For j = LBound(sheetPairs) To UBound(sheetPairs)
        Dim logCol1 As String, logCol2 As String, logCol3 As String
        Dim specCol1 As String, specCol2 As String, specCol3 As String
        Dim columns As Variant
        
        ' ペアに応じて対応する列を選択
        Select Case sheetPairs(j)(0)
            Case "LOG_Helmet"
                columns = helmetColumns
            Case "LOG_FallArrest"
                columns = fallArrestColumns
            Case "LOG_Bicycle"
                columns = bicycleColumns
            Case "LOG_BaseBall"
                columns = baseBallColumns
            Case Else
                Debug.Print "ペアが見つかりませんでした: " & sheetPairs(j)(0)
        End Select
        
        ' 列の割り当て
        If Not IsEmpty(columns) Then
            logCol1 = columns(0)
            logCol2 = columns(1)
            logCol3 = columns(2)
            specCol1 = columns(3)
            specCol2 = columns(4)
            specCol3 = columns(5)
            
            ' シートの存在確認
            On Error Resume Next
            Set logSheet = ActiveWorkbook.Sheets(sheetPairs(j)(0))
            Set specSheet = ActiveWorkbook.Sheets(sheetPairs(j)(1))
            On Error GoTo 0
            
            If Not logSheet Is Nothing And Not specSheet Is Nothing Then
                logLastRow = logSheet.Cells(logSheet.Rows.Count, "C").End(xlUp).row
                specLastRow = specSheet.Cells(specSheet.Rows.Count, "C").End(xlUp).row

                ' LOGシートの2行目以降をループ
                For i = 2 To logLastRow
                    If i <= specLastRow Then
                        ' 別のプロシージャで比較を行う
                        If CompareRows(logSheet, specSheet, i, logCol1, logCol2, logCol3, specCol1, specCol2, specCol3) Then
                            logSheet.Range(logSheet.Cells(i, "D"), logSheet.Cells(i, "O")).Interior.Color = RGB(255, 0, 0)
                            mismatchFound = True
                            Debug.Print "不一致行: " & i & " (シート: " & logSheet.Name & ")"
                        End If
                    End If
                Next i
            End If
        End If
    Next j
    
    ' 不一致行がない場合にメッセージボックスとハイライトリセットを実行
    If Not mismatchFound Then
        MsgBox "不一致行は見つかりませんでした。", vbInformation, "結果"
        
        ' ハイライトリセット: すべてのシートのハイライトをクリア
        For j = LBound(sheetPairs) To UBound(sheetPairs)
            Set logSheet = ActiveWorkbook.Sheets(sheetPairs(j)(0))
            On Error Resume Next ' シートが存在しない場合を考慮
            logLastRow = logSheet.Cells(logSheet.Rows.Count, "C").End(xlUp).row
            logSheet.Range(logSheet.Cells(2, "D"), logSheet.Cells(logLastRow, "O")).Interior.colorIndex = xlNone
        Next j
    End If
End Sub

' 比較ロジックを別のプロシージャに分割
Function CompareRows(logSheet As Worksheet, specSheet As Worksheet, row As Long, _
                     logCol1 As String, logCol2 As String, logCol3 As String, _
                     specCol1 As String, specCol2 As String, specCol3 As String) As Boolean
    Dim logValue1 As Variant, specValue1 As Variant
    Dim logValue2 As Variant, specValue2 As Variant
    Dim logValue3 As Variant, specValue3 As Variant
    
    ' LOGシートとSpecSheetの値を取得
    logValue1 = logSheet.Cells(row, logCol1).value
    logValue2 = logSheet.Cells(row, logCol2).value
    logValue3 = logSheet.Cells(row, logCol3).value
    specValue1 = specSheet.Cells(row, specCol1).value
    specValue2 = specSheet.Cells(row, specCol2).value
    specValue3 = specSheet.Cells(row, specCol3).value
    
    ' デバッグ出力
'    Debug.Print "LOGシート行: " & row & " - " & logValue1 & ", " & logValue2 & ", " & logValue3
'    Debug.Print "SpecSheet行: " & row & " - " & specValue1 & ", " & specValue2 & ", " & specValue3

    ' 比較処理
    If logValue1 <> specValue1 Or logValue2 <> specValue2 Or logValue3 <> specValue3 Then
        CompareRows = True
    Else
        CompareRows = False
    End If
End Function










