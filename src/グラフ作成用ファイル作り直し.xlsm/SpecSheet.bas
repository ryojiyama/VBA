Attribute VB_Name = "SpecSheet"
 ' ☆品番、試験箇所などに応じたIDを作成する
Sub CreateID()
   
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim ID As String

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")

    ' 最後の行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' 各行に対してIDを生成
    For i = 2 To lastRow ' 1行目はヘッダと仮定
        ID = GenerateID(ws, i)
        ' B列にIDをセット
        ws.Cells(i, 2).value = ID
    Next i
End Sub

Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' CreateID()のサブプロシージャ
    Dim ID As String

    ' C列: 2桁以下の数字
    ID = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    ID = ID & "-" ' C列とD列の間に"-"
    ' D列の処理を変更
    ID = ID & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    ID = ID & "-" ' FmとE列の間に"-"
    ID = ID & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E列の条件
    ID = ID & "-" ' FmとE列の間に"-
    ID = ID & GetColumnIValue(ws.Cells(rowIndex, 9).value) ' I列の条件
    ID = ID & "-" ' I列とL列の間に"-"
    ID = ID & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L列の条件

    GenerateID = ID
End Function
Function ExtractNumberWithF(value As String) As String
' GenerateIDのサブ関数
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' 正規表現オブジェクトの作成
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "d{3,6}"
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
    ElseIf InStr(value, "側頭部") > 0 Then
        Dim pos As Integer
        pos = InStr(value, "_")
        If pos > 0 Then
            GetColumnEValue = "側" & Mid(value, pos)
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

' ☆SpecSheetに転記するプロシージャの本体。アイコンに紐づけ。
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

    Call ProcessSheetPairs          ' 転記処理をするプロシージャ
    Call CustomizeSheetFormats      ' 各列に書式設定をする
    Call TransformIDs               ' B列にIDを作成する。
    Call Utlities.FillBlanksWithHyphenInMultipleSheets
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
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' 色のインデックスを初期化
    Dim colorIndex As Integer
    colorIndex = 3 ' Excelの色インデックスは3から始まる

    ' H列の2行目から最終行までループ
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            If ws.Cells(i, "H").value = ws.Cells(j, "H").value And ws.Cells(i, "H").value <> "" Then
                ' 同値を持つセルが見つかった場合、フラグをTrueに設定し、セルに色を塗る
                foundDuplicate = True
                ws.Cells(i, "H").Interior.colorIndex = colorIndex
                ws.Cells(j, "H").Interior.colorIndex = colorIndex
                ws.Cells(i, "H").Interior.colorIndex = colorIndex ' 同値が見つかったセルに色を塗る
            End If
        Next j
        ' 同値が見つかった場合、次の色に変更
        If foundDuplicate And ws.Cells(i, "H").Interior.colorIndex <> xlNone Then
            colorIndex = colorIndex + 1
            ' 色インデックスの最大値を超えないようにチェック
            If colorIndex > 56 Then colorIndex = 3 ' 色インデックスをリセット
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
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' 最終列を"M"(試験区分)に固定
    Dim lastCol As Long
    lastCol = ws.Columns("M").Column

    ' 指定範囲をループ
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' 空白のチェック
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "空白セル: " & cell.Address(False, False) & vbNewLine
            End If

            ' 列G、H、J、Kで数値の確認
            If j = Columns("G").Column Or j = Columns("H").Column Or j = Columns("J").Column Or j = Columns("K").Column Then
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
                cell.NumberFormat = "General"
            End If

            ' 列N、O、Pで文字列の確認
            If j = Columns("N").Column Or j = Columns("O").Column Or j = Columns("P").Column Then
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
        If SheetExists(logSheetName) And SheetExists(specSheetName) Then
            ' シートペアが成立した場合に処理を実行
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
            Debug.Print logSheetName
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

    ' ワークシートをセット
    Set logSheet = ThisWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ThisWorkbook.Worksheets(sheetNameSpec)

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
                logCol = col.Column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.value = pair(1) Then
                helSpecCol = col.Column
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

    ' 転記が行われたことを確認
    MsgBox "転記が完了しました: " & sheetNameLog & " から " & sheetNameSpec
End Sub


Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    'ProcessSheetPairsのサブプロシージャ
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
            headerPairs = Array( _
                Array("試料ID", "試験ID(C)"), _
                Array("品番", "品番(D)"), _
                Array("試験内容", "試験内容(E)"), _
                Array("検査日", "検査日(F)"), _
                Array("温度", "温度(G)"), _
                Array("前処理", "前処理(L)"), _
                Array("重量", "重量(M)"), _
                Array("天頂すきま", "天頂すきま(N)"), _
                Array("帽体色", "帽体色(O)"), _
                Array("試験区分", "試験区分(U)") _
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

' 各列に書式設定をする
Sub CustomizeSheetFormats()

    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            ' Determine the data type based on the column header and set the format accordingly
            Select Case True
                Case InStr(cell.value, "検査日") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "yyyy-mm-dd"
                Case InStr(cell.value, "温度") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "最大値(kN)") > 0, InStr(cell.value, "重量") > 0, _
                     InStr(cell.value, "天頂すきま") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "最大値を記録した時間") > 0, _
                     InStr(cell.value, "4.9kNの継続時間") > 0, _
                     InStr(cell.value, "7.3kNの継続時間") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "ID") > 0, InStr(cell.value, "試料ID") > 0, _
                     InStr(cell.value, "品番") > 0, InStr(cell.value, "試験位置") > 0, _
                     InStr(cell.value, "前処理") > 0, InStr(cell.value, "帽体色") > 0, _
                     InStr(cell.value, "製品ロット") > 0, InStr(cell.value, "帽体ロット") > 0, _
                     InStr(cell.value, "内装ロット") > 0, InStr(cell.value, "構造検査") > 0, _
                     InStr(cell.value, "貫通検査") > 0, InStr(cell.value, "試験区分") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "@"
            End Select
        Next cell
    Next sheet


End Sub

Sub CustomizeSheetFormats_Old()

    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "最大値(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.00 "
            ElseIf InStr(1, cell.value, "最大値(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0 "
            ElseIf InStr(1, cell.value, "時間") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "温度") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "重量") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "ロット") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "@"
            ElseIf InStr(1, cell.value, "天頂すきま") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            End If
        Next cell
    Next sheet
End Sub
' B列にIDを作成する。
Sub TransformIDs()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim newID As String
    
    ' LOG_Helmetシートを設定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' 2行目から最終行までループ（1行目はヘッダーと仮定）
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "C").value
        
        ' IDを変換
        newID = GenerateNewID(cellValue)
        
        ' 新しいIDをセルにセット
        ws.Cells(i, "B").value = newID
    Next i
End Sub

Function GenerateNewID(cellValue As String) As String
    'TransformIDsのサブプロシージャ
    Dim numPart As String
    Dim otherPart As String
    Dim newID As String
    Dim matches As Object
    Dim reNum As Object
    Dim reOther As Object
    Dim startIndex As Long
    
    ' 数値部分の正規表現オブジェクトを作成
    Set reNum = CreateObject("VBScript.RegExp")
    reNum.Global = False
    reNum.IgnoreCase = False
    reNum.Pattern = "d{3,5}F?"
    
    ' 数値部分を抽出
    If reNum.Test(cellValue) Then
        Set matches = reNum.Execute(cellValue)
        numPart = ExtractNumberPart(matches(0).value)
        newID = numPart
        
        ' 特定の文字列に続く部分を抽出
        otherPart = ExtractOtherPart(cellValue, reNum.Execute(cellValue)(0).FirstIndex + 1)
        
        ' デバッグ用の出力
        Debug.Print numPart
        Debug.Print otherPart
        
        ' 新しいIDを結合
        GenerateNewID = newID & otherPart
    Else
        ' 数値部分が見つからない場合は元の値を返す
        GenerateNewID = cellValue
    End If
End Function

Function ExtractNumberPart(numPart As String) As String
        'TransformIDsのサブプロシージャ
    Dim hasF As Boolean
    ' 数字部分の末尾がFの場合
    hasF = Right(numPart, 1) = "F"
    If hasF Then
        ' 末尾のFを除去して数値部分を取得
        numPart = Left(numPart, Len(numPart) - 1)
        ' 新しいIDを生成（前後にFを追加）
        ExtractNumberPart = "F" & numPart & "F"
    Else
        ' 末尾にFがない場合はそのまま使用
        ExtractNumberPart = numPart
    End If
End Function

Function ExtractOtherPart(cellValue As String, startIndex As Long) As String
    'TransformIDsのサブプロシージャ
    Dim reOther As Object
    Dim matches As Object
    Dim otherPart As String
    Dim endIndex As Long
    
    ' 特定の文字列に続く部分を抽出するための正規表現
    Set reOther = CreateObject("VBScript.RegExp")
    reOther.Global = False
    reOther.IgnoreCase = False
    reOther.Pattern = "-(天|前|後|側)"
    
    If reOther.Test(cellValue) Then
        startIndex = reOther.Execute(cellValue)(0).FirstIndex + 1
        otherPart = Mid(cellValue, startIndex)
        
        ' 最後の'-'以降の文字を取り除く
        endIndex = InStrRev(otherPart, "-")
        If endIndex > 0 Then
            otherPart = Left(otherPart, endIndex - 1)
        End If
        ExtractOtherPart = otherPart
    Else
        ExtractOtherPart = ""
    End If
End Function







