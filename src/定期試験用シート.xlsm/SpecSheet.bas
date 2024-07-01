Attribute VB_Name = "SpecSheet"
Sub SetupTestSamples()
    Call CreateInspectionSheetIDs
    Call InsertXLookupAndUpdateKColumn
End Sub


Sub SyncSpecSheetToLogHel()
    ' アイコンに紐づけ。SpecSheetに転記するプロシージャのまとめ
    ' 同値が見つかった場合はエラーメッセージを表示して処理を中断
    If HighlightDuplicateValues Then
        MsgBox "衝撃値で同値が見つかりました。小数点下二桁に影響が出ない範囲で修正してください。", vbCritical
        Exit Sub
    End If
    
    Dim errMsg As String
    errMsg = LocateEmptySpaces()
    
    If errMsg <> "" Then
        ' エラーメッセージがある場合、それを表示
        MsgBox "以下の問題があります。まずはこれらを解決してください：" & vbNewLine & errMsg, vbCritical
        Exit Sub
    Else
    End If
    
    Call CopyDataBasedOnCondition
    Call CustomizeSheetFormats
    MsgBox "転記が終了しました。"
End Sub


Sub CreateInspectionSheetIDs_0410Before()
    ' SpecSheetのB列に試験IDを作成する。これは転記するときのキーとして使用する。
    
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' D列の最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row

    Dim i As Long
    For i = 2 To lastRow
        ' D列に値がある行の場合のみ処理
        If ws.Cells(i, "D").value <> "" Then
            ' S列に式を設定
            ws.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""高温"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""低温"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""浸せき"", ""Wet"", """")))"

            ' IDを作成
            Dim id As String
            id = ws.Cells(i, "D").value & "-" & ws.Cells(i, "S").value & "-" & Left(ws.Cells(i, "E").value, 1)

            ' D列の値に"F"が含まれている場合、IDの先頭に"F"を追加
            If InStr(ws.Cells(i, "D").value, "F") > 0 Then
                id = "F" & id
            End If

            ' 作成したIDをB列に設定
            ws.Cells(i, "B").value = id
            ws.Cells(i, "Q").value = "合格"
            ws.Cells(i, "R").value = "合格"
        End If
    Next i
End Sub

Sub CreateInspectionSheetIDs()
    Dim wsSpecSheet As Worksheet
    Set wsSpecSheet = ThisWorkbook.Sheets("Hel_SpecSheet")

    Dim wsSetting As Worksheet
    Set wsSetting = ThisWorkbook.Sheets("Setting")

    Dim lastRow As Long
    lastRow = wsSpecSheet.Cells(wsSpecSheet.Rows.count, "D").End(xlUp).row

    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    For i = 2 To lastRow
        If wsSpecSheet.Cells(i, "D").value <> "" Then
            wsSpecSheet.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""高温"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""低温"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""浸せき"", ""Wet"", """")))"
            Dim id As String
            id = wsSpecSheet.Cells(i, "D").value & "-" & wsSpecSheet.Cells(i, "S").value & "-" & Left(wsSpecSheet.Cells(i, "E").value, 1)

            foundMatch = False
            For j = 2 To wsSetting.Cells(wsSetting.Rows.count, "H").End(xlUp).row
                If wsSpecSheet.Cells(i, "D").value = wsSetting.Cells(j, "H").value Then
                    foundMatch = True
                    If InStr(wsSetting.Cells(j, "J").value, "x") > 0 Then
                        id = "F" & id
                    End If
                    Exit For
                End If
            Next j

            If Not foundMatch Then
                MsgBox "エラー: D列の値がSettingシートのH列と一致する項目がありません。処理を中止します。"
                Exit Sub
            End If

            wsSpecSheet.Cells(i, "B").value = id
            wsSpecSheet.Cells(i, "Q").value = "合格"
            wsSpecSheet.Cells(i, "R").value = "合格"
        End If
    Next i
End Sub

Sub InsertXLookupAndUpdateKColumn()
    ' "Hel_SpecSheet"の天頂隙間を調整する
    ' 調整した天頂隙間の行に"Changed"を入れてわかりやすくしている。
    Dim wsHelSpecSheet As Worksheet
    Dim wsSetting As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim formulaResult As Variant
    Dim kValue As Variant
    
    ' シートの設定
    Set wsHelSpecSheet = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' D列の最終行を探索
    lastRow = wsHelSpecSheet.Cells(wsHelSpecSheet.Rows.count, "D").End(xlUp).row
    
    ' D列を探索し、値がある各行に対して処理を実行
    For i = 2 To lastRow
        If wsHelSpecSheet.Cells(i, "D").value <> "" Then
            ' T列にXLOOKUP関数を代入
            wsHelSpecSheet.Cells(i, "T").Formula = "=XLOOKUP(TEXT(Hel_SpecSheet!D" & i & ", ""0""), " & _
                "TEXT(Setting!$H$2:$H$49, ""0""), " & _
                "Setting!$I$2:$I$49, """")"

            ' XLOOKUP関数の結果を取得
            formulaResult = wsHelSpecSheet.Cells(i, "T").value
            
            ' K列の値を取得
            kValue = wsHelSpecSheet.Cells(i, "K").value
            
            ' K列の値からT列の値を引いて、結果をK列に代入
            wsHelSpecSheet.Cells(i, "K").value = kValue - formulaResult
            
            ' U列に'Changed'を代入
            wsHelSpecSheet.Cells(i, "U").value = "Changed"
        End If
    Next i
End Sub


Function HighlightDuplicateValues() As Boolean
    ' シート名を変数で定義
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
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row
    
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


Function LocateEmptySpaces() As String
    ' "Hel_SpecSheet"に空欄またはデータ型の誤りがないかをチェック
    ' 変数宣言と初期化
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    Dim lastCol As Long
    lastCol = ws.Columns("S").Column
    Dim errorMsg As String
    errorMsg = ""
    
    ' 指定範囲をループしてエラーチェック
    For i = 2 To lastRow
        For j = 2 To lastCol
            Dim cell As Range
            Set cell = ws.Cells(i, j)
            ' 空白のチェック
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "空白セル: " & cell.Address(False, False) & vbNewLine
            End If
            ' 数値チェック
            If (j = 7 Or j = 8 Or j = 10 Or j = 11) And Not IsNumeric(cell.value) Then
                errorMsg = errorMsg & "数値でないセル: " & cell.Address(False, False) & vbNewLine
            End If
            ' 文字列チェック
'            If (j = 14 Or j = 15 Or j = 16) And Not VarType(cell.Value) = vbString Then
'                errorMsg = errorMsg & "文字列でないセル: " & cell.Address(False, False) & vbNewLine
'            End If
        Next j
    Next i
    
    ' エラーメッセージがあればそれを返し、なければ空の文字列を返す
    LocateEmptySpaces = errorMsg
End Function



Sub CopyDataBasedOnCondition()
    'SpecSheetの内容をLogシートに転記する
    Dim logSheet As Worksheet
    Dim helSpec As Worksheet
    Dim lastRowLog As Long
    Dim lastRowSpec As Long
    Dim i As Long, j As Long
    Dim matchCount As Long

    ' ワークシートをセット
    Set logSheet = ThisWorkbook.Worksheets("LOG_Helmet")
    Set helSpec = ThisWorkbook.Worksheets("Hel_SpecSheet")

    ' LOG_Helmetの最終行を取得
    lastRowLog = logSheet.Cells(logSheet.Rows.count, "H").End(xlUp).row
    ' Hel_SpecSheetの最終行を取得
    lastRowSpec = helSpec.Cells(helSpec.Rows.count, "H").End(xlUp).row

    ' LOG_HelmetのH列の値を整える
'    For i = 2 To lastRowLog
'        logSheet.Cells(i, "H").Value = Application.Round(logSheet.Cells(i, "H").Value, 2)
'    Next i

    ' 値を比較して転記
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").value = helSpec.Cells(j, "H").value Then
                ' H列の値が一致した場合、各列の内容を転記
                matchCount = matchCount + 1
                logSheet.Cells(i, "B").value = helSpec.Cells(j, "B").value
                logSheet.Cells(i, "C").value = helSpec.Cells(j, "B").value
                logSheet.Cells(i, "D").value = helSpec.Cells(j, "D").value
                logSheet.Cells(i, "E").value = helSpec.Cells(j, "E").value
                logSheet.Cells(i, "F").value = helSpec.Cells(j, "F").value
                logSheet.Cells(i, "G").value = helSpec.Cells(j, "G").value
                logSheet.Cells(i, "L").value = helSpec.Cells(j, "I").value
                logSheet.Cells(i, "M").value = helSpec.Cells(j, "J").value
                logSheet.Cells(i, "N").value = helSpec.Cells(j, "K").value '天頂すきま
                logSheet.Cells(i, "O").value = helSpec.Cells(j, "L").value
                logSheet.Cells(i, "U").value = helSpec.Cells(j, "M").value '試験内容
                logSheet.Cells(i, "P").value = helSpec.Cells(j, "N").value '製造ロット
                logSheet.Cells(i, "Q").value = helSpec.Cells(j, "O").value
                logSheet.Cells(i, "R").value = helSpec.Cells(j, "P").value
                logSheet.Cells(i, "S").value = helSpec.Cells(j, "Q").value '構造結果
                logSheet.Cells(i, "T").value = helSpec.Cells(j, "R").value
                'logSheet.Cells(i, "U").Value = helSpec.Cells(j, "S").Value
                'logSheet.Cells(i, "U").Value = helSpec.Cells(j, "U").Value
                
            End If
        Next j
        
        ' 一致した値が複数存在する場合、文字を太字にする
        If matchCount > 1 Then
            logSheet.Cells(i, "C").Font.Bold = True
            logSheet.Cells(i, "D").Font.Bold = True
            logSheet.Cells(i, "E").Font.Bold = True
            logSheet.Cells(i, "F").Font.Bold = True
            logSheet.Cells(i, "G").Font.Bold = True
            logSheet.Cells(i, "L").Font.Bold = True
            logSheet.Cells(i, "M").Font.Bold = True
            logSheet.Cells(i, "N").Font.Bold = True
            logSheet.Cells(i, "O").Font.Bold = True
        End If
    Next i
End Sub


Sub CustomizeSheetFormats()
' 各列に書式設定をする
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "最大値(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.00 ""kN"""
            ElseIf InStr(1, cell.value, "最大値(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0 ""G"""
            ElseIf InStr(1, cell.value, "時間") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""ms"""
            ElseIf InStr(1, cell.value, "温度") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""℃"""
            ElseIf InStr(1, cell.value, "重量") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""g"""
            ElseIf InStr(1, cell.value, "ロット") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "@"
            ElseIf InStr(1, cell.value, "天頂すきま") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 ""mm"""
            End If

        Next cell
    Next sheet
End Sub

Sub UniformizeLineGraphAxes()

    ' Display input dialog to set the maximum value for the axes
    Dim MaxValue As Double
    MaxValue = InputBox("Y軸の最大値を入力してください。(整数)", "最大値を入力")
    
    ' Loop through all the charts in the active sheet
    Dim chartObj As ChartObject
    For Each chartObj In ActiveSheet.ChartObjects
        With chartObj.chart.Axes(xlValue)
            ' Set the Y-axis maximum value
            .MaximumScale = MaxValue
        End With
    Next chartObj

End Sub

