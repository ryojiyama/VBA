Attribute VB_Name = "Utlities"
' レポートグラフの印刷範囲を設定する
Sub SetPrintAreaForGroups()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim groupCount As Integer
    Dim groupRows() As Long
    Dim i As Long
    Dim printStart As Long
    Dim printEnd As Long
    Dim j As Integer
    Dim fileName As String
    Dim pdfFolder As String
    Dim currentGroup As String
    Dim previousGroup As String

    ' 保存するPDFのフォルダパスを設定
    pdfFolder = ThisWorkbook.Path & "\PDFs\"
    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder ' フォルダが存在しない場合、作成
    End If

    ' シートをループして"レポートグラフ"を含むシートを処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "レポートグラフ") > 0 Then
            ' シートの最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row

            ' グループの行を保持するための配列を動的に確保
            ReDim groupRows(0)

            ' "Insert + 数値"が含まれる行を見つけて配列に保存
            groupCount = 0 ' 初期化
            previousGroup = ""
            For i = 1 To lastRow
                currentGroup = ws.Cells(i, "I").value
                If InStr(currentGroup, "Insert") > 0 And currentGroup <> previousGroup Then
                    groupCount = groupCount + 1
                    ReDim Preserve groupRows(groupCount)
                    groupRows(groupCount - 1) = i
                    previousGroup = currentGroup ' グループが変わった時のみ更新
                End If
            Next i

            ' グループ数に応じて印刷範囲を設定
            If groupCount = 0 Then
                MsgBox "印刷範囲に該当するグループが見つかりませんでした。"
                Exit For ' シートがない場合は次に進む
            End If

            ' グループを2つずつ印刷範囲に設定
            For j = 0 To groupCount - 1 Step 2
                printStart = groupRows(j) ' グループの開始行を設定

                If j + 2 < groupCount Then
                    ' 次の次のグループの開始行の1行前を終了行とする
                    printEnd = groupRows(j + 2) - 1
                Else
                    ' 最後のグループが単独の場合、最後の行まで含める
                    printEnd = lastRow
                End If

                ' 印刷範囲を設定
                ws.PageSetup.PrintArea = ws.Range("A" & printStart & ":G" & printEnd).Address

                ' グループと印刷範囲をデバッグウインドウに表示
                If j + 1 < groupCount Then
                    Debug.Print "シート: " & ws.Name & ", グループ: " & ws.Cells(groupRows(j), "I").value & " - " & ws.Cells(groupRows(j + 1), "I").value & ", 印刷範囲: " & ws.PageSetup.PrintArea
                Else
                    Debug.Print "シート: " & ws.Name & ", グループ: " & ws.Cells(groupRows(j), "I").value & ", 印刷範囲: " & ws.PageSetup.PrintArea
                End If

                ' PDF出力（コメントアウト中）
                 fileName = pdfFolder & ws.Name & "_Group_" & j + 1 & ".pdf"
                 ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard
            Next j
        End If
    Next ws
End Sub






' レポートグラフシートの内容を削除する。
Sub DeleteContentFromReportGraphSheets()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' シートをループして、名前に「レポートグラフ」が含まれているシートを処理
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "レポートグラフ") > 0 Then
            ' A列の最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

            ' 最終行が1以上であれば、行を削除する
            If lastRow > 0 Then
                ws.Rows("1:" & lastRow).Delete
            End If
        End If
    Next ws
End Sub


