Attribute VB_Name = "TransferToReport"
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
    dict.Add "B", "C" ' 転記元のB列 → 転記先のC列

    ' 必要に応じて他の列のマッピングも追加

    Set GetMappingDictionary = dict
End Function

Private Sub TransferMappedValues(wsSource As Worksheet, wsDest As Worksheet, sourceRow As Long, destRow As Long, mappingDict As Object)
    ' マッピングに基づいて値を転記する
    Dim key As Variant

    ' マッピングディクショナリをループして転記を実行
    For Each key In mappingDict.Keys
        wsDest.Cells(destRow, Columns(mappingDict(key)).Column).value = wsSource.Cells(sourceRow, Columns(key).Column).value
    Next key
End Sub

Private Sub ApplyFormattingToNewRows(ws As Worksheet, startRow As Long, endRow As Long)
    ' 新しく追加された行にフォーマットを適用する
    Dim targetRange As Range

    ' 新しく追加された行の範囲を設定
    Set targetRange = ws.Range("B" & startRow & ":G" & endRow)

    ' フォーマットを適用
    With targetRange
        .Font.Bold = True ' フォントを太字に設定
        .Interior.color = RGB(220, 230, 241) ' 背景色を設定（青系）
        .Borders.LineStyle = xlContinuous ' 罫線を設定
    End With
End Sub



