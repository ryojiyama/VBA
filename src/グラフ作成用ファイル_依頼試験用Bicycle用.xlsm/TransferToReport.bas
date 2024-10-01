Attribute VB_Name = "TransferToReport"
' レポート本文の表に結果を挿入する。ヘルメットのものをそのままコピー修正が必要
Sub TransferDataWithMappingAndFormatting()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim i As Long
    Dim destRow As Long
    Dim transferredRows As Long
    Const MAX_ROWS As Long = 12 ' 最大転記行数を12に設定
    Dim startRow As Long
    Dim mappingDict As Object ' マッピング用のディクショナリ

    ' シートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet") ' 転記元シート
    Set wsDest = ThisWorkbook.Sheets("レポート本文") ' 転記先シート

    ' 転記元の最終行を取得
    lastRowSource = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).row

    ' 転記先の開始行を設定（9行目から開始）
    destRow = 9
    startRow = destRow ' 新しく追加した行の開始行を記録
    transferredRows = 0 ' 転記した行数をカウント

    ' マッピング用のディクショナリを取得
    Set mappingDict = GetMappingDictionary()

    ' 転記元の2行目から最終行までループ
    For i = 2 To lastRowSource
        ' 転記した行が12行に達したら中止
        If transferredRows >= MAX_ROWS Then
            MsgBox "転記は最大" & MAX_ROWS & "行までに制限されています。処理を中止しました。", vbExclamation
            Exit For
        End If

        ' 転記先の行を追加（destRowの位置に新しい行を挿入）
        wsDest.Rows(destRow).Insert Shift:=xlDown

        ' 転記を実行（ディクショナリに基づく転記）
        Call TransferMappedValues(wsSource, wsDest, i, destRow, mappingDict)

        ' 次の転記先行に進む
        destRow = destRow + 1
        transferredRows = transferredRows + 1 ' 転記した行数をカウント
    Next i

    ' 新しく追加された行にフォーマットを適用
    Call ApplyFormattingToNewRows(wsDest, startRow, destRow - 1)

    ' 全行が転記された場合はメッセージを表示
    If transferredRows < MAX_ROWS Then
        MsgBox "データの転記が完了しました。", vbInformation
    End If

End Sub

Private Function GetMappingDictionary() As Object
    ' 転記元と転記先のマッピングをディクショナリで設定
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 転記元の列 → 転記先の列を列名として明示的に記述
    dict.Add "D", "B" ' 転記元のD列 → 転記先のB列
    dict.Add "E", "C" ' 転記元のE列 → 転記先のC列
    dict.Add "L", "D" ' 転記元のL列 → 転記先のD列
    dict.Add "H", "E" ' 転記元のH列 → 転記先のE列
    dict.Add "M", "F" ' 転記元のM列 → 転記先のF列
    dict.Add "N", "G" ' 転記元のN列 → 転記先のG列
    ' 必要に応じて他の列のマッピングも追加

    Set GetMappingDictionary = dict
End Function

Private Sub TransferMappedValues(wsSource As Worksheet, wsDest As Worksheet, sourceRow As Long, destRow As Long, mappingDict As Object)
    ' マッピングに基づいて値を転記する
    Dim key As Variant

    ' マッピングディクショナリをループして転記を実行
    For Each key In mappingDict.Keys
        wsDest.Cells(destRow, Columns(mappingDict(key)).column).value = wsSource.Cells(sourceRow, Columns(key).column).value
    Next key
End Sub
Private Sub ApplyFormattingToNewRows(ws As Worksheet, startRow As Long, endRow As Long)
    ' 新しく追加された行にフォーマットを適用し、I列に印をつける
    Dim currentRow As Long
    Dim targetRange As Range
    Dim eRange As Range, fRange As Range, gRange As Range, iRange As Range
    
    ' 1行ずつ処理
    For currentRow = startRow To endRow
        ' 現在の行の範囲を取得
        Set targetRange = ws.Range("B" & currentRow & ":G" & currentRow)
        
        ' フォーマットを適用
        With targetRange
            .Font.Name = "游ゴシック" ' フォント名を設定
            .Font.ThemeFont = xlThemeFontMinor ' Lightウェイトにする（テーマフォント）
            .Font.Bold = False ' 太字を解除
            .Font.Color = RGB(0, 0, 0) ' フォントの色を黒に設定
            
            ' 背景色を行ごとに変更
            If currentRow Mod 2 = 0 Then
                ' 偶数行：薄い青色
                .Interior.Color = RGB(220, 230, 241)
            Else
                ' 奇数行：薄い灰色
                .Interior.Color = RGB(255, 255, 255)
            End If
            
            .Borders.LineStyle = xlContinuous ' 罫線を設定
        End With
        
        ' E列に 0.00 "kN" の書式設定
        Set eRange = ws.Range("E" & currentRow)
        eRange.NumberFormat = "0.00 ""kN"""
        eRange.HorizontalAlignment = xlRight ' 右寄せ

        ' F列に 0.0 "g" の書式設定
        Set fRange = ws.Range("F" & currentRow)
        fRange.NumberFormat = "0.0 ""g"""
        fRange.HorizontalAlignment = xlRight ' 右寄せ

        ' G列に 0.0 "mm" の書式設定
        Set gRange = ws.Range("G" & currentRow)
        gRange.NumberFormat = "0.0 ""mm"""
        gRange.HorizontalAlignment = xlRight ' 右寄せ

        ' I列に "Insert + 行番号" の印を付ける
        Set iRange = ws.Range("I" & currentRow)
        iRange.value = "Insert " & currentRow
    Next currentRow
End Sub





