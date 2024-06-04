Attribute VB_Name = "SpecSheet"
Sub CreateIDplusDrawBorders()
    '転記作業とフォーマット整理、アイコンに紐づけ
    Call CreateID
    Call DrawBordersWithHairline
End Sub

Sub CreateID()
    '品番、試験箇所などに応じたIDを作成する
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
        
        ' C列: 2桁以下の数字
        If Len(ws.Cells(i, 3).Value) <= 2 Then
            ID = Right("00" & ws.Cells(i, 3).Value, 2)
        Else
            ID = "??"
        End If
        
        ' C列とD列の間に"-"
        ID = ID & "-"
        
        ' D列: 左から4文字目から6文字目の文字列
        ID = ID & Mid(ws.Cells(i, 4).Value, 4, 3)
        
        ' E列の条件
        Select Case ws.Cells(i, 5).Value
            Case "天頂"
                ID = ID & "T"
            Case "前頭部"
                ID = ID & "F"
            Case "後頭部"
                ID = ID & "R"
            Case Else
                ID = ID & "?"
        End Select
        
        ' I列の条件
        Select Case ws.Cells(i, 9).Value
            Case "高温"
                ID = ID & "H"
            Case "低温"
                ID = ID & "L"
            Case "浸せき"
                ID = ID & "W"
            Case Else
                ID = ID & "?"
        End Select
        
        ' I列とL列の間に"-"
        ID = ID & "-"
        
        ' L列の条件
        If ws.Cells(i, 12).Value = "白" Then
            ID = ID & "W"
        Else
            ID = ID & "O"
        End If
        
        ' B列にIDをセット
        ws.Cells(i, 2).Value = ID
    Next i

End Sub

Sub DrawBordersWithHairline()
    ' シート「Hel_SpecSheet」を選択
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' シートの既存の罫線を全て消去
    ws.Cells.Borders.LineStyle = xlNone

    ' C列の最終行を探索
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' 範囲 Cells(2, "B"):Cells(lastRow, "M") に新たに罫線を引く（1行目は除外）
    With ws.Range(ws.Cells(2, "B"), ws.Cells(lastRow, "M"))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlHairline
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlHairline

        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlHairline
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
End Sub



Sub SyncSpecSheetToLogHel()
    ' アイコンに紐づけ。SpecSheetに転記するプロシージャのまとめ
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

    Call CopyDataBasedOnCondition
    Call CustomizeSheetFormats
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
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' 色のインデックスを初期化
    Dim colorIndex As Integer
    colorIndex = 3 ' Excelの色インデックスは3から始まる

    ' H列の2行目から最終行までループ
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            If ws.Cells(i, "H").Value = ws.Cells(j, "H").Value And ws.Cells(i, "H").Value <> "" Then
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
            ws.Cells(i, "H").Interior.color = xlNone
        Next i
    End If

    ' 同値が見つかったかどうかに基づいて結果を返す
    HighlightDuplicateValues = foundDuplicate
End Function

Function LocateEmptySpaces() As Boolean
    ' "Hel_SpecSheet"に空欄がないかをチェック
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
    lastCol = ws.Columns("M").column

    ' 指定範囲をループ
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' 空白のチェック
            If IsEmpty(cell.Value) Then
                errorMsg = errorMsg & "空白セル: " & cell.Address(False, False) & vbNewLine
            End If

            ' 列G、H、J、Kで数値の確認
            If j = Columns("G").column Or j = Columns("H").column Or j = Columns("J").column Or j = Columns("K").column Then
                If Not IsNumeric(cell.Value) Then
                    errorMsg = errorMsg & "数値でないセル: " & cell.Address(False, False) & vbNewLine
                End If
            End If

            ' 列N、O、Pで文字列の確認
            If j = Columns("N").column Or j = Columns("O").column Or j = Columns("P").column Then
                If Not VarType(cell.Value) = vbString Then
                    errorMsg = errorMsg & "文字列でないセル: " & cell.Address(False, False) & vbNewLine
                End If
            End If
        Next j
    Next i

    ' エラーメッセージがあれば表示し、Falseを返す
    If Len(errorMsg) > 0 Then
        LocateEmptySpaces = False
        Exit Function
    Else
        LocateEmptySpaces = True
    End If
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
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).row
    ' Hel_SpecSheetの最終行を取得
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).row

    ' LOG_HelmetのH列の値を整える
'    For i = 2 To lastRowLog
'        logSheet.Cells(i, "H").Value = Application.Round(logSheet.Cells(i, "H").Value, 2)
'    Next i

    ' 値を比較して転記
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").Value = helSpec.Cells(j, "H").Value Then
                ' H列の値が一致した場合、各列の内容を転記
                matchCount = matchCount + 1
                logSheet.Cells(i, "C").Value = helSpec.Cells(j, "B").Value
                logSheet.Cells(i, "D").Value = helSpec.Cells(j, "D").Value
                logSheet.Cells(i, "E").Value = helSpec.Cells(j, "E").Value
                logSheet.Cells(i, "F").Value = helSpec.Cells(j, "F").Value
                logSheet.Cells(i, "G").Value = helSpec.Cells(j, "G").Value
                logSheet.Cells(i, "L").Value = helSpec.Cells(j, "I").Value
                logSheet.Cells(i, "M").Value = helSpec.Cells(j, "J").Value
                logSheet.Cells(i, "N").Value = helSpec.Cells(j, "K").Value
                logSheet.Cells(i, "O").Value = helSpec.Cells(j, "L").Value
                logSheet.Cells(i, "U").Value = helSpec.Cells(j, "M").Value
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
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.Value, "最大値(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.00 ""kN"""
            ElseIf InStr(1, cell.Value, "最大値(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0 ""G"""
            ElseIf InStr(1, cell.Value, "時間") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""ms"""
            ElseIf InStr(1, cell.Value, "温度") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""℃"""
            ElseIf InStr(1, cell.Value, "重量") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "0.0 ""g"""
            ElseIf InStr(1, cell.Value, "ロット") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = "@"
            ElseIf InStr(1, cell.Value, "天頂すきま") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.column).End(xlUp))
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
    Dim ChartObj As ChartObject
    For Each ChartObj In ActiveSheet.ChartObjects
        With ChartObj.chart.Axes(xlValue)
            ' Set the Y-axis maximum value
            .MaximumScale = MaxValue
        End With
    Next ChartObj

End Sub

