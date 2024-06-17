# 2024/01/13

```vb
Dim j as long
```

|  A       |  B   |  C   |  D   |     H     |     I      |
|----------|------|------|------|-----------|------------|
| 型式名    | No.1 | No.2 | No.3 |  登録日    |  更新費用   |
|----------|------|------|------|-----------|------------|
| No.100   | 3707 |  -   |  -   | 2022-01-01|  $200      |
| No.105   | 4468 | 4469 |  -   | 2022-02-15|  $200      |
| No.110   | 2668 |  -   | TF644| 2022-03-20|  $200      |
| No.110F  | 2669 | 2670 | TF644| 2022-03-20|  $200      |
| No.110S  | 3920 | 3921 | TF924| 2022-05-05|  $200      |
------------------------------------------------------------

以下の条件を満たすVBAのコードを作成してください。主な機能は"InspectionSheet"のB列に所定のIDを作成するものです。
# 条件
- 範囲はD列の2列目から最終行まで。
- D列に値がある行のS列に関数"=IF(I1="高温", "Hot", IF(I1="低温", "Cold", IF(I1="浸漬", "Wet", "")))
"を入れてください。
- B列に"D2の数値のみ & - & E列の最初の１文字 & - & S列の値"になるようにIDを作成してください。
- また、D列の値に"F"が含まれている場合、完成したIDの先頭にFをつけてください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

```vb
Sub CreateInspectionSheetIDs()
    ' シート名を変数で指定できるようにする
    Dim sheetName As String
    sheetName = "InspectionSheet"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' D列の最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        ' D列に値がある行の場合のみ処理
        If ws.Cells(i, "D").Value <> "" Then
            ' S列に式を設定
            ws.Cells(i, "S").Formula = "=IF(I1=""高温"", ""Hot"", IF(I1=""低温"", ""Cold"", IF(I1=""浸漬"", ""Wet"", """")))"

            ' IDを作成
            Dim id As String
            id = ws.Cells(i, "D").Value & "-" & Left(ws.Cells(i, "E").Value, 1) & "-" & ws.Cells(i, "S").Value

            ' D列の値に"F"が含まれている場合、IDの先頭に"F"を追加
            If InStr(ws.Cells(i, "D").Value, "F") > 0 Then
                id = "F" & id
            End If

            ' 作成したIDをB列に設定
            ws.Cells(i, "B").Value = id
        End If
    Next i
End Sub
```

ws.Cells(i, "S").Formula = "=IF(INDIRECT(""R"" & i & ""C9"", FALSE)=""高温"", ""Hot"", IF(INDIRECT(""R"" & i & ""C9"", FALSE)=""低温"", ""Cold"", IF(INDIRECT(""R"" & i & ""C9"", FALSE)=""浸漬"", ""Wet"", """")))"

ws.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""高温"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""低温"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""浸漬"", ""Wet"", """")))"


以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"から"InspectionSheet"に指定の値を転記するものです。
# 条件
- "Log_Helmet"のC列の値を参照し、C列の値のうち、先頭に"F"がつくもの以外の数を数え、その数だけ"InspectionSheet"を複製します。
- そしてC列の先頭にFがつき、L列が高温の場合の数を数え、その数だけ"InspectionSheet"を複製します。。
- 複製したシートは"Inspection_Draft"の右側に"Draft_1", "Draft_2"と名前をつけ、並べていきます。名前の重複が起こらないようにしてください。
- シートの複製が終わったら複製したシートの名前を"LOG_Helmet"シートのC列の2行目の値"に変更していきます。変更するときにシートの名前が被らないようにしてください。また、"LOG_Helemt"のL列が高温以外の列はスキップしてください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- コード内のシート名は後で変更できるようにしてください。

# コード
```vb
Sub DuplicateAndRenameSheets()
    Dim wsLogHelmet As Worksheet, wsTemplate As Worksheet, wsDraft As Worksheet
    Dim countNonF As Long, countHighTemp As Long, i As Long, sheetIndex As Long
    Dim sheetName As String

    ' シート名の設定
    Const LOG_HELMET As String = "Log_Helmet"
    Const TEMPLATE_SHEET As String = "InspectionSheet"
    Const DRAFT_PREFIX As String = "Inspection_Draft"

    Set wsLogHelmet = ThisWorkbook.Sheets(LOG_HELMET)
    Set wsTemplate = ThisWorkbook.Sheets(TEMPLATE_SHEET)

    ' C列の値をチェックし、条件に応じて数をカウント
    For i = 1 To wsLogHelmet.Cells(wsLogHelmet.Rows.Count, 3).End(xlUp).Row
        If Left(wsLogHelmet.Cells(i, 3).Value, 1) <> "F" Then
            countNonF = countNonF + 1
        ElseIf wsLogHelmet.Cells(i, 3).Value Like "F*" And wsLogHelmet.Cells(i, 12).Value = "高温" Then
            countHighTemp = countHighTemp + 1
        End If
    Next i

    ' シートの複製
    For i = 1 To countNonF + countHighTemp
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsDraft = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsDraft.Name = DRAFT_PREFIX & "_" & i
    Next i

    ' シート名の変更
    sheetIndex = 1
    For i = 1 To wsLogHelmet.Cells(wsLogHelmet.Rows.Count, 3).End(xlUp).Row
        If wsLogHelmet.Cells(i, 12).Value = "高温" Or Left(wsLogHelmet.Cells(i, 3).Value, 1) <> "F" Then
            Do
                sheetName = wsLogHelmet.Cells(i, 3).Value
                If Not SheetExists(sheetName) Then
                    ThisWorkbook.Sheets(DRAFT_PREFIX & "_" & sheetIndex).Name = sheetName
                    sheetIndex = sheetIndex + 1
                    Exit Do
                Else
                    sheetIndex = sheetIndex + 1
                End If
            Loop While sheetIndex <= countNonF + countHighTemp
        End If
    Next i
End Sub

Function SheetExists(sheetName As String) As Boolean
Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function
```


このコードの解説を以下に示します：

- `Sub DuplicateAndRenameSheets()`はメインのプロシージャで、指定された条件に基づいてシートの複製と名前変更を行います。
- まず、必要なワークシートを変数に設定し、"Log_Helmet"シートのC列とL列を参照して条件に合致する行の数をカウントします。
- 次に、"InspectionSheet"を基に指定された数だけシートを複製し、"Draft_1", "Draft_2"等の名前を付けます。
- 最後に、"Log_Helmet"シートのC列の値に基づいて各複製シートの名前を変更します。この際、既に存在するシート名との重複を避けるために`SheetExists`関数を使用します。

注意点：
- 実際のExcelファイルによっては、このコードを調整する必要があります。
- エラーハンドリングは基本的なものですが、より詳細なエラーチェックが必要な場合は追加してください。
- コードは、"Log_Helmet"シートのC列とL列のデータに依存しています。これらの列に予期せぬ値がある場合、コードが正しく動作しない可能性があります。

以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"から他のシートに指定の値を転記するものです。
プロシージャの名前は"TransferDataToTopImpactTest"にしてください。
# 条件
- "Log_Helmet"のC列の値を参照し、C列の２行目から最終行までの項目の数だけ、転記作業を実行します。
- C列のそれぞれの行の１文字目に"F"が含まれていない行を探します。この行を1行ずつ１シートに転記します。
- C列の１文字目から"-"までとシートの名前の１文字目から"-"までを探します。この条件を仮に"MatchName"とします。
- 条件が当てはまった行のD列の値を条件の当てはまったシートのB3に「D列の値 & "_" & O列の値」を入力する。
- 条件が当てはまった行のF列の値を条件の当てはまったシートのE2に入力する。
- 条件が当てはまった行のG列の値を条件の当てはまったシートのG2に入力する。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

プロシージャの呼び出し、または引数が不正です。

了解しました。ただもう一つ条件を加えたいです。
Left(wsDest.Name, dashPosDest - 1) = matchName
に加えて、[LOG_Helmet]シートのL列の値が"高温"の場合は当てはまるシート名に"Hot"が含まれ、"低温"の場合は"Cold"、"浸漬"の場合は"Wet"が
含まれるシートを転記の対象にしてください。

wsDest.Range("C9").Value = "製品規格【産業用安全帽】AⅡQ-09-08に寄る。"
を"AⅡQ-09-08に寄る。"の部分をセル内で改行して表示したい。その部分のコードのみ提示してください。

以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"から他のシートに指定の値を転記するものです。
プロシージャの名前は"TransferDataToAllImpactTest"にしてください。
# 条件
- "Log_Helmet"のC列の値を参照し、C列の２行目から最終行までの項目の数だけ、転記作業を実行します。
- C列のそれぞれの行の１文字目に"F"が含まれている行を探します。この行を1行ずつ１シートに転記します。
- C列のそれぞれの行の１文字目から"-"までとシートの名前の１文字目から"-"までを探します。この条件を仮に"MatchName"とします。
- "MatchName"に合致する条件に加え、"LOG_Helemt"のE列の値が"天頂"の場合は該当する行の"LOG_Helemt_H列"の値を転記先のシートのD11に転記します。
- "MatchName"に合致する条件に加え、"LOG_Helemt"のE列の値が"前頭部"の場合は該当する行の"LOG_Helemt_H列"の値を転記先のシートのD14に転記します。
- "MatchName"に合致する条件に加え、"LOG_Helemt"のE列の値が"後頭部"の場合は外合する行の"LOG_Helemt_H列"の値を転記先のシートのD19に転記します。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

```vb
Sub TransferDataToAllSheets()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceName As String
    Dim matchName As String, sheetMatchName As String

    sourceName = "LOG_Helmet"  ' 転記元のシート名
    Set wsSource = ThisWorkbook.Sheets(sourceName)

    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' ワークブック内の全てのシートをループ
    For Each wsDest In ThisWorkbook.Sheets
        ' 現在のシート名から MatchName を取得
        If InStr(wsDest.Name, "-") > 0 Then
            sheetMatchName = Left(wsDest.Name, InStr(wsDest.Name, "-") - 1)

            ' C列をループし条件に合うデータを転記
            For i = 2 To lastRow
                ' MatchName の取得
                matchName = Left(wsSource.Cells(i, 3).Value, InStr(wsSource.Cells(i, 3).Value, "-") - 1)

                ' 条件に合致する場合に転記
                If matchName = sheetMatchName Then
                    ' E列の値に基づいて処理を分岐
                    Select Case wsSource.Cells(i, 5).Value
                        Case "天頂"
                            ' 天頂の場合の転記先を指定
                            wsDest.Cells(3, 1).Value = wsSource.Cells(i, 8).Value
                        Case "前頭部"
                            ' 前頭部の場合の転記先を指定
                            wsDest.Cells(4, 1).Value = wsSource.Cells(i, 8).Value
                        Case "後頭部"
                            ' 後頭部の場合の転記先を指定
                            wsDest.Cells(5, 1).Value = wsSource.Cells(i, 8).Value
                    End Select
                End If
            Next i
        End If
    Next wsDest
End Sub
```

```vb
Sub FormatTransferredCells()
    Dim wsLogHelmet As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmetシートを設定
    Set wsLogHelmet = ThisWorkbook.Sheets("LOG_Helmet")

    ' B列の最終行を取得
    lastRow = wsLogHelmet.Cells(wsLogHelmet.Rows.Count, "B").End(xlUp).row

    ' B列の各行をループ
    For i = 1 To lastRow
        sheetName = wsLogHelmet.Cells(i, 2).Value

        ' ワークシートが存在するかチェック
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        ' ワークシートが存在すれば、指定した結合セルの書式を設定
        If Not wsTarget Is Nothing Then
            ' 書式を設定する範囲を指定
            With wsTarget.Range("B2:G9")
                ' フォント設定
                .Font.name = "游ゴシック"  ' フォントの種類
                .Font.Size = 12       ' フォントのサイズ
                .Font.Bold = True     ' 太字にする

                ' 背景色を設定
                .Interior.color = RGB(255, 255, 0)  ' 背景色（黄色）

                ' 境界線を設定
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            End With

            Set wsTarget = Nothing
        End If
    Next i
End Sub

If matchName = sheetMatchName Then
    ' E列の値に基づいて処理を分岐
    Select Case wsSource.Cells(i, 5).Value
        Case "天頂"
            ' 天頂の場合の転記先を指定
            wsTestForm.Range("D11").Value = wsSource.Cells(i, 8).Value
        Case "前頭部"
            ' 前頭部の場合の転記先を指定
            wsTestForm.Range("D14").Value = wsSource.Cells(i, 8).Value
        Case "後頭部"
            ' 後頭部の場合の転記先を指定
            wsTestForm.Range("D19").Value = wsSource.Cells(i, 8).Value
    End Select
End If
の部分ですが、天頂という言葉をws.Source.Cells(i, 5)から検索した後に転記する値はそれに対応した行になります。
つまり転記元の値はws.Source.Cells(i, 5)に天頂という言葉が入っている行から引っ張ってきてください。
指示が不明瞭ならその部分を質問してください。


アプローチを変更します。
まず最初にwsSource.Cells(i, 5).Valueの中から天頂などの言葉を探します。その後にwsSource.Cells(i, 12).Valueの
中から、高温、低温、浸漬という言葉を探し、それごとに転記作業を行うように変更してください。


If matchName = sheetMatchName Then
    ' E列の値に基づいて処理を分岐
    Select Case wsSource.Cells(i, 5).Value
        Case "天頂"
            ' 天頂の場合の処理
            Select Case wsSource.Cells(i, 12).Value
                Case "高温"
                    ' 高温の場合の転記先を指定
                    wsTestForm.Range("D11").Value = wsSource.Cells(i, 8).Value
                Case "低温"
                    ' 低温の場合の転記先を指定
                    wsTestForm.Range("D14").Value = wsSource.Cells(i, 8).Value
                Case "浸漬"
                    ' 浸漬の場合の転記先を指定
                    wsTestForm.Range("D19").Value = wsSource.Cells(i, 8).Value
        Case "前頭部"
            ' 前頭部の場合の処理
            Select Case wsSource.Cells(i, 12).Value
                Case "高温"
                    ' 高温の場合の転記先を指定
                    ' （転記先のセルを指定）
                Case "低温"
                    ' 低温の場合の転記先を指定
                    ' （転記先のセルを指定）
                Case "浸漬"
                    ' 浸漬の場合の転記先を指定
                    ' （転記先のセルを指定）
            End Select

        Case "後頭部"
            '後頭部の場合の処理
            Select Case wsSource.Cells(i, 12).Value
                Case "高温"
                    ' 高温の場合の転記先を指定
                    ' （転記先のセルを指定）
                Case "低温"
                    ' 低温の場合の転記先を指定
                    ' （転記先のセルを指定）
                Case "浸漬"
                    ' 浸漬の場合の転記先を指定
                    ' （転記先のセルを指定）
            End Select

        ' その他のケース（必要に応じて）
    End Select
End If

'"F170F-*-Hot"の形式でソートする方法
Sub TransferDataToAllSheets()
    Dim wsSource As Worksheet, wsTestForm As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceName As String
    Dim matchString As String

    sourceName = "LOG_Helmet"  ' 転記元のシート名
    Set wsSource = ThisWorkbook.Sheets(sourceName)

    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row

    ' ワークブック内の全てのシートをループ
    For Each wsTestForm In ThisWorkbook.Sheets
        ' wsSourceのC列をループし条件に合うデータを転記
        For i = 2 To lastRow
            matchString = wsSource.Cells(i, 3).Value
            ' C列の値が"F170F-*-Hot"パターンに一致するかチェック
            If matchString Like "F170F-*-Hot" And wsTestForm.name Like "F170F-*-Hot" Then
                ' E列の値に基づいて処理を分岐
                Select Case wsSource.Cells(i, 5).Value
                    Case "天頂"
                        ' 天頂の場合の処理
                        wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value '衝撃値
                    Case "前頭部"
                        ' 前頭部の場合の処理
                        wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value '衝撃値
                    Case "後頭部"
                        ' 後頭部の場合の処理
                        wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value '衝撃値
                End Select
            End If
        Next i
    Next wsTestForm
End Sub



```
それでは"F170F-*-Hot" でソートする方法をもう少し拡張したいと思います。
F170FのFで囲まれた数字ですがこれを任意の3~4の数字に変更できますか？
試しに If matchString Like "F170F-*-Hot" And wsTestForm.name Like "F170F-*-Hot" Then
の部分のみ変更してください。


```vb
If (matchString Like "F###*-Hot" Or matchString Like "F####*-Hot") And _
   (wsTestForm.name Like "F###*-Hot" Or wsTestForm.name Like "F####*-Hot") Then
```
そのようにした場合、例えば、F170F-天-HotとF390F-天-Hotの区別はつくでしょうか？

F170F-天-Hot,F390F-天-Hot,F391F-天-Hotの3パターンを作成し、実行したところF391F-天-Hotのパターンのみが3つのシートに適用されました。
改善方法はありますか？
If (matchString Like "F[0-9][0-9][0-9]F-*-Hot" Or matchString Like "F[0-9][0-9][0-9][0-9]F-*-Hot") And _
   (wsTestForm.Name Like "F[0-9][0-9][0-9]F-*-Hot" Or wsTestForm.Name Like "F[0-9][0-9][0-9][0-9]F-*-Hot") Then

わかりました。アプローチを変更したいと思います。
以下の条件を満たす分岐を作成してください。
# 条件
1. wsSource.Cells(i, 3).Valueの"F170F-天-Hot"の1文字目から"-"までの文字とwsTestForm.nameの1文字目から"-"までの文字を比較し、一致するグループを作成します。
2. そのグループの中からwsSource.Cells(i, 12).Valueが一致するグループを作成します。
3. そしてCase文を用いてwsSource.Cells(i, 5).Valueの値がそれぞれ、"天頂","前頭部", "後頭部",の条件で分岐を作成します。

ひとまずこの条件を満たす分岐の部分のみを簡単にVBAで書いてください。

```vb
Dim sourcePrefix As String
Dim testFormPrefix As String
Dim category As String

' wsSourceのC列とwsTestFormの名前を比較するためのプレフィックスを取得
sourcePrefix = Left(wsSource.Cells(i, 3).Value, InStr(wsSource.Cells(i, 3).Value, "-") - 1)
testFormPrefix = Left(wsTestForm.Name, InStr(wsTestForm.Name, "-") - 1)

' プレフィックスが一致するかチェック
If sourcePrefix = testFormPrefix Then
    ' wsSourceのL列の値を取得
    category = wsSource.Cells(i, 12).Value

    ' 一致するカテゴリに基づいて処理を分岐
    If category = "適切なカテゴリ1" Then
        ' E列の値に基づいて処理を分岐
        Select Case wsSource.Cells(i, 5).Value
            Case "天頂"
                ' 天頂の場合の処理
                ' ...
            Case "前頭部"
                ' 前頭部の場合の処理
                ' ...
            Case "後頭部"
                ' 後頭部の場合の処理
                ' ...
        End Select
    ElseIf category = "適切なカテゴリ2" Then
        ' 他のカテゴリに対する処理
        ' ...
    End If
End If
```
了解しました。変数categoryをprepTempとし、"適切なカテゴリ1,2,3"をそれぞれ、高温、低温、浸漬にします。
これらを踏まえて全体のコードを示していただけますか。

このコードの特定のパターンのフィルタリングには問題があるので条件を変更したいと思います。
ひとまず下記の条件を満たす分岐の部分のみを簡単にVBAで書いてください。
# 条件
1. wsSource.Cells(i, 3).Valueの"F170F-天-Hot"の1文字目から"-"までの文字とwsTestForm.nameの1文字目から"-"までの文字を比較し、一致するグループを作成します。
2. そのグループの中からwsSource.Cells(i, 12).Valueにはそれぞれ高温、低温、浸漬の3種類の値が入っているので、1.のグループの中から値が一致する行のグループを作成します。。
3. そしてCase文を用いてwsSource.Cells(i, 5).Valueの値がそれぞれ、"天頂","前頭部", "後頭部",の条件で分岐を作成します。
この分岐処理に問題がある場合はお知らせください。

```vb
For Each wsTestForm In ThisWorkbook.Sheets
    For i = 2 To lastRow
        ' "-"までの文字列を取得
        Dim sourcePrefix As String
        sourcePrefix = Left(wsSource.Cells(i, 3).Value, InStr(1, wsSource.Cells(i, 3).Value, "-") - 1)

        Dim testFormPrefix As String
        testFormPrefix = Left(wsTestForm.Name, InStr(1, wsTestForm.Name, "-") - 1)

        ' 条件1: wsSourceとwsTestFormのプレフィックスが一致するか
        If sourcePrefix = testFormPrefix Then
            ' 条件2: wsSourceの12列目の値が一致するか
            If wsSource.Cells(i, 12).Value = wsTestForm.Cells(1, 12).Value Then  ' 例として、wsTestFormの12列目1行目の値と比較
                ' 条件3: ケース文による分岐
                Select Case wsSource.Cells(i, 5).Value
                    Case "天頂"
                        ' 天頂の場合の処理
                        ' ... (処理内容を記述)
                    Case "前頭部"
                        ' 前頭部の場合の処理
                        ' ... (処理内容を記述)
                    Case "後頭部"
                        ' 後頭部の場合の処理
                        ' ... (処理内容を記述)
                End Select
            End If
        End If
    Next i
Next wsTestForm

Sub TransferDataToAllSheets()
    Dim wsSource As Worksheet, wsTestForm As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceName As String
    Dim sourcePrefix As String, testFormPrefix As String

    sourceName = "LOG_Helmet"  ' 転記元のシート名
    Set wsSource = ThisWorkbook.Sheets(sourceName)

    ' 最終行を取得
    ' lastRow のデバッグ出力
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row
    Debug.Print "Last row in LOG_Helmet: " & lastRow

    ' ワークブック内の全シート数のデバッグ出力
    Debug.Print "Total sheets in workbook: " & ThisWorkbook.Sheets.Count

    ' ワークブック内の全てのシートをループ
    For Each wsTestForm In ThisWorkbook.Sheets
        ' wsSourceのC列をループし条件に合うデータを転記
        For i = 2 To lastRow
            Dim sourcePrefix As String
            sourcePrefix = Left(wsSource.Cells(i, 3).Value, InStr(1, wsSource.Cells(i, 3).Value, "-") - 1)

            Dim testFormPrefix As String
            For Each wsTestForm In ThisWorkbook.Sheets
                ' シート名の最初の文字が "F" であるかどうかをチェック
                If Left(wsTestForm.Name, 1) = "F" Then
                    ' ハイフンが存在するかどうかをチェック
                    If InStr(1, wsTestForm.Name, "-") > 0 Then
                        ' ハイフンがある場合、ハイフンまでの文字列を testFormPrefix に割り当て
                        testFormPrefix = Left(wsTestForm.Name, InStr(1, wsTestForm.Name, "-") - 1)
                    Else
                        ' ハイフンがない場合、このシートをスキップ
                        Continue For
                    End If
                Else
                    ' シート名の最初の文字が "F" でない場合、このシートをスキップ
                    Continue For
                End If

            If sourcePrefix = testFormPrefix Then
                ' 条件2: wsSourceの12列目の値に基づく分岐
                Select Case wsSource.Cells(i, 12).Value
                    Case "高温"
                        ' 高温の場合の処理
                        Select Case wsSource.Cells(i, 5).Value
                            Case "天頂"
                                wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            Case "前頭部"
                                wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                            Case "後頭部"
                                wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                        End Select
                    Case "低温"
                        ' 低温の場合の処理
                        Select Case wsSource.Cells(i, 5).Value
                            Case "天頂"
                                wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            Case "前頭部"
                                wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                            Case "後頭部"
                                wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                        End Select
                    Case "浸漬"
                        ' 浸漬の場合の処理
                        Select Case wsSource.Cells(i, 5).Value
                            Case "天頂"
                                wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            Case "前頭部"
                                wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                            Case "後頭部"
                                wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                        End Select
                End Select
            End If

            Next wsTestForm

        Next i
    Next wsTestForm
End Sub

Sub TransferDataToAllSheets()
    Dim wsSource As Worksheet, wsTestForm As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceName As String
    Dim sourcePrefix As String, testFormPrefix As String

    sourceName = "LOG_Helmet"  ' 転記元のシート名
    Set wsSource = ThisWorkbook.Sheets(sourceName)

    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).row

    ' ワークブック内の全てのシートをループ
    For Each wsTestForm In ThisWorkbook.Sheets
        ' シート名の最初の文字が "F" で、ハイフンが存在するかをチェック
        If Left(wsTestForm.Name, 1) = "F" Then
            If InStr(1, wsTestForm.Name, "-") > 0 Then
                testFormPrefix = Left(wsTestForm.Name, InStr(1, wsTestForm.Name, "-") - 1)
            Else
                ' ハイフンがない場合、このシートをスキップ
                Exit For
            End If
            ' wsSourceのC列をループし条件に合うデータを転記
            For i = 2 To lastRow
                sourcePrefix = Left(wsSource.Cells(i, 3).Value, InStr(1, wsSource.Cells(i, 3).Value, "-") - 1)

                If sourcePrefix = testFormPrefix Then
                    ' 条件2: wsSourceの12列目の値に基づく分岐
                    Select Case wsSource.Cells(i, 12).Value
                    Debug.print "条件2：" i
                        Case "高温"
                            ' 高温の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                     Debug.print "条件天頂：" i
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                        Case "低温"
                            ' 低温の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                        Case "浸漬"
                            ' 浸漬の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                    End Select
                End If

            Next i
        End If
    Next wsTestForm
End Sub
```

```vb
Case "高温"
    ' 高温の場合の処理
    Select Case wsSource.Cells(i, 5).Value
        Case "天頂"
            Debug.Print "条件:高温のシート："; wsTestForm.name
            Debug.Print "条件:高温天頂："; i
            wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
        Case "前頭部"
            wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
        Case "後頭部"
            wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
    End Select

```
以上のコードを条件に沿うように修正したいと思います。修正案のみを提示してください。
# 条件
iの値をwsSource.Cells(i, 5).Valueを探索した値に合わせて変化させたいです。例えば"前頭部"を見つけた場合は8行目にあるのでCells(i, 5)のiの値は8になります。
```vb
' wsSourceの各行をループ
For i = 2 To lastRow
    ' wsSourceの5列目の値に基づく条件チェック
    Select Case wsSource.Cells(i, 5).Value
        Case "天頂"
            ' 高温かつ天頂の場合の処理
            If wsSource.Cells(i, 12).Value = "高温" Then
                Debug.Print "条件:高温のシート：" & wsTestForm.Name
                Debug.Print "条件:高温天頂：" & i
                wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
            End If
        Case "前頭部"
            ' 高温かつ前頭部の場合の処理
            If wsSource.Cells(i, 12).Value = "高温" Then
                wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
            End If
        Case "後頭部"
            ' 高温かつ後頭部の場合の処理
            If wsSource.Cells(i, 12).Value = "高温" Then
                wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
            End If
    End Select
Next i
```
                If sourcePrefix = testFormPrefix Then
                    ' 条件2: wsSourceの12列目の値に基づく分岐
                    Select Case wsSource.Cells(i, 12).Value
                        Case "高温"
                            ' 高温の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                    Debug.Print "条件:高温のシート："; wsTestForm.name
                                    Debug.Print "条件:高温天頂："; i
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                        Case "低温"
                            ' 低温の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                    Debug.Print "条件：低温天頂："; i
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                        Case "浸漬"
                            ' 浸漬の場合の処理
                            Select Case wsSource.Cells(i, 5).Value
                                Case "天頂"
                                    Debug.Print "条件：浸漬天頂："; i
                                    wsTestForm.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                                Case "前頭部"
                                    wsTestForm.Cells(14, 4).Value = wsSource.Cells(i, 8).Value
                                Case "後頭部"
                                    wsTestForm.Cells(19, 4).Value = wsSource.Cells(i, 8).Value
                            End Select
                    End Select
                End If

wsSourceのC列には以下のように並んでいます。
F170F-天-Hot
F170F-天-Cold
F170F-天-Wet
F170F-前-Hot
F170F-前-Cold
F170F-前-Wet
F170F-後-Hot
F170F-後-Cold
F170F-後-Wet
F170F-天-Hot、F170F-前-Hot、F170F-後-Hotの行のデータをwsTestFormシートの各部分に転記したいのです。どのようなアプローチがありますか？アイデアをいくつか教えて下さい。

シート名の決定の際に最初の文字列のみで判断するのは、不正確な結果をもたらすかもしれません。"F170F-天-Hot"のうち、"F170F"だけでなく"Hot"も加えることは可能でしょうか？
シート名は別のプロシージャによって作成されています。名前は"F170F-天-Hot"のような法則性でつけられており、変更することは難しいです。
シート名は法則性を持っていますが、数値などは多種多様です。"F170F-天-Cold","F3220F-前-Hot","F390F-後-Wet"などいろいろなパターンがあります。
```vb
Sub TransferDataToDynamicSheets()
    Dim wsSource As Worksheet, wsTestForm As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        parts = Split(sourceData, "-")

        ' シート名の動的生成（例："F170F-Hot"）
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(0) & "-天-" & parts(UBound(parts))

            ' シートが存在するか確認し、存在する場合のみ転記
            If SheetExists(destinationSheetName) Then
                Set wsTestForm = ThisWorkbook.Sheets(destinationSheetName)
                Debug.Print "wsTestForm" wsTestForm
                Debug.Print "destinationSheetName" destinationSheetName
                ' 必要なデータを転記
                ' 例: wsTestForm.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value
            End If
        End If
    Next i
End Sub

' シートが存在するかどうかを確認する関数
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function
```

次のステップです。F170F-天-Hot、F170F-前-Hot、F170F-後-Hotの行のデータを"F170F-天-Hot"シートの各部分に転記したいのです。
コードの例でいうと、part(0) & -*- & part(2)のようになるでしょうか。アプローチのアイデアを教えて下さい。

"F170F-天-Cold","F3220F-前-Hot","F390F-後-Wet"などいろいろなパターンがあるシート名にも対応したいです。

```vb
Sub TransferDataToDynamicSheet()
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set wsDestination = ThisWorkbook.Sheets("F170F-Hot-天") ' 目的のシートを設定
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        parts = Split(sourceData, "-")

        ' データ形式が条件に一致するかチェック
        If UBound(parts) >= 2 And (parts(2) = "天" Or parts(2) = "前" Or parts(2) = "後") Then
            ' 条件に基づくデータ転記
            ' 例: wsDestination.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value
            ' 適切な行と列のインデックスを設定する必要があります
        End If
    Next i
End Sub

```
 destinationSheetName = parts(0) & "-天-" & parts(UBound(parts))でシート名の動的生成を行っていますが、
 この部分を拡張し、下記の条件に適合させてください。
 # 条件
 - F170F-天-Hot、F170F-前-Hot、F170F-後-Hotの行のデータを"F170F-天-Hot"シート一枚に転記する。
 - destinationSheetName = parts(0) & "-前-" & parts(UBound(parts)),destinationSheetName = parts(0) & "-後-" & parts(UBound(parts))の名前を擬似的に作成し、F170F-前-Hot、F170F-後-Hotと一致できるようにする。
 - 上記と一致したら'F170F-天-Hot'シートに転記する。
このアプローチにするとどのような問題が起きるか教えて下さい。

では複雑性を避けるためにマッチングするの最適な名前を考えようと思います。目的はF170F-天-Hot、F170F-前-Hot、F170F-後-Hotの行のデータを"F170F-天-Hot"シート一枚に転記する。ですが、それぞれの名前は変更しても構いません。どのようなアイデアがありますか？


F170F-Hot-天、F170F-Hot-前、F170F-Hot-後の行のデータを"F170F-Hot-天"シート一枚に転記する。
"F170F-天-Cold","F3220F-前-Hot","F390F-後-Wet"などいろいろなパターンがあるシート名にも対応したいです。


以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"から他のシートに指定の値を転記するものです。
プロシージャの名前は"TransferDataToAllImpactTest"にしてください。
# 条件
- "Log_Helmet"のC列の値を参照し、C列の２行目から最終行までの項目の数だけ、転記作業を実行します。
- C列のそれぞれの行の１文字目に"F"が含まれている行を探します(例：F170F-Hot-後)。この行を１シートに転記します。
- 'F170F-Hot-天'、'F170F-Hot-前'、'F170F-Hot-後'の行のデータを"F170F-Hot-天"シート一枚に転記する。
- "F170F-Cold-天","F3220F-Hot-前","F390F-後-Wet"などいろいろなパターンがあるシート名にも対応してください。
- そのために"parts = Split(sourceData, "-")"などでの名称操作も考慮してください。
- 当てはまるシートがない場合は作成せず、GoToやExitForでスキップするようにしてください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

```vb
Sub TransferDataToAllImpactTest()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim sheetName As String

    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("Log_Helmet")

    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' C列のデータを順にチェック
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value

        ' C列の1文字目が"F"であるか確認
        If Left(sourceData, 1) = "F" Then
            ' ハイフンで文字列を分割
            parts = Split(sourceData, "-")

            ' シート名を設定（例: F170F-Hot-天）
            sheetName = parts(0) & "-" & parts(1) & "-" & parts(2)

            ' 当てはまるシートがあるか確認
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0

            ' シートが存在する場合、データを転記
            If Not wsTarget Is Nothing Then
                ' ここに転記処理を記述
                ' 例: wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Offset(1, 0).Value = sourceData

                ' ターゲットシートをリセット
                Set wsTarget = Nothing
            End If
        End If
    Next i
End Sub

- 'F170F-Hot-天'、'F170F-Hot-前'、'F170F-Hot-後'の行のデータを"F170F-Hot-天"シート一枚に転記する。
- "F170F-Cold-天","F3220F-Hot-前","F390F-後-Wet"などいろいろなパターンがあるシート名にも対応してください。
この両方の条件に対応するために以下のアプローチを試みてください。

' シート名を設定（例: F170F-Hot-天）
sheetName = parts(0) & "-" & parts(1) & "-" & parts(2)
をsheetName = parts(0) & "-" & parts(1)とし、
C列の文字列をwsSource.Cells(i, 3).Value = parts(0) & "-" & parts(1) のように改変した値と比較し、
一致したときに処理を行う。

これは試作のコードであり、エラーが出ても状況は回復可能です。失敗を恐れず取り組んでください。




Sub TransferDataToDynamicSheets()
    Dim wsSource As Worksheet, wsTestForm As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        parts = Split(sourceData, "-")

        ' シート名の動的生成（例："F170F-Hot"）
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(0) & "-" & parts(UBound(parts))

            ' シートが存在するか確認し、存在する場合のみ転記
            If SheetExists(destinationSheetName) Then
                Set wsTestForm = ThisWorkbook.Sheets(destinationSheetName)
                Debug.Print "wsTestForm" wsTestForm
                Debug.Print "destinationSheetName" destinationSheetName
                ' 必要なデータを転記
                ' 例: wsTestForm.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value
            End If
        End If
    Next i
End Sub

' シートが存在するかどうかを確認する関数
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

それでは上記のコードを参考に条件を満たすVBAのコードを作成してください。
# 条件
- C列のデータもparts(0) & "-" & parts(1)のように作り替え、シート名と照合する。
- parts(UBound(parts))は、天, 前, 後、の3種類。それに合わせてCase文で処理を分岐する。

Sub TransferDataToDynamicSheets()
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim modifiedSourceData As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        parts = Split(sourceData, "-")

        If UBound(parts) >= 2 Then
            ' データとシート名を作成
            modifiedSourceData = parts(0) & "-" & parts(1)
            destinationSheetName = modifiedSourceData

            ' シートが存在するか確認し、存在する場合のみ転記
            If SheetExists(destinationSheetName) Then
                Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)

                ' parts(UBound(parts))に基づいて処理を分岐
                Select Case parts(UBound(parts))
                    Case "天"
                        ' 天に関するデータ転記の処理
                        Debug.Print "wsDestination:天："wsDestination.name
                        ' 例: wsDestination.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value

                    Case "前"
                        ' 前に関するデータ転記の処理
                        Debug.Print "wsDestination:前："wsDestination.name
                        ' 例: wsDestination.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value

                    Case "後"
                        ' 後に関するデータ転記の処理
                        Debug.Print "wsDestination:後："wsDestination.name
                        ' 例: wsDestination.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value

                    Case Else
                        ' その他の値の場合の処理（必要に応じて）
                End Select
            End If
        End If
    Next i
End Sub

' シートが存在するかどうかを確認する関数
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

了解しました。ここで必要なことは、シート名とカスタムしたC行の値が一致することを確認することです。つまり、modifiedSourceData = parts(0) & "-" & parts(1)と一致したシート名に転記作業を行いたいです。コードの修正は可能でしょうか？

「F170F-Hot-天」,「F170F-Hot-前」,「F170F-Hot-後」
If Left(sourceData, Len(modifiedSourceData)) = modifiedSourceDataの部分がよくわかりません。「F170F-Hot-天」,「F170F-Hot-前」,「F170F-Hot-後」のサンプルを用いて説明をお願いします。


Sub TransferDataToDynamicSheets()
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim modifiedSourceData As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        parts = Split(sourceData, "-")

        If UBound(parts) >= 2 Then
            ' データとシート名を作成
            modifiedSourceData = parts(0) & "-" & parts(1)
            destinationSheetName = modifiedSourceData

            ' シートが存在するか確認し、存在する場合のみ転記
            If InspectionSheetExists(destinationSheetName) Then
                Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)

                ' modifiedSourceData と sourceData の最初の2つの部分が一致する場合にのみ転記
                If Left(sourceData, Len(modifiedSourceData)) = modifiedSourceData Then
                    ' ここに転記処理を記述
                    ' 例: wsDestination.Cells(適切な行, 適切な列).Value = wsSource.Cells(i, 適切な列).Value
                End If
            End If
        End If
    Next i
End Sub

' シートが存在するかどうかを確認する関数
Function InspectionSheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    InspectionSheetExists = Not sheet Is Nothing
End Function



よくわかりました。コードは正常に機能しました。次のステップに移りたいと思います。
Case "天"で分岐をしていますが、これに条件を加えたいと思います。
wsSourceのE列とこのCaseの分岐をリンクさせたいです。つまり
Case "天"の場合wsSource.Cells(i, 5)の値が"天頂"
Case "前"の場合wsSource.Cells(i, 5)の値が"前頭部"
Case "後"の場合wsSource.Cells(i, 5)の値が"後頭部"
これらに一致する行の値を転記するようにできるでしょうか？


Sub TransferDataToDynamicSheets()
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim modifiedSourceData As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' wsSourceのC列をループ
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).Value
        checkData = wsSource.Cells(i, 5).Value
        parts = Split(sourceData, "-")

        If UBound(parts) >= 2 Then
            ' データとシート名を作成
            modifiedSourceData = parts(0) & "-" & parts(1)
            destinationSheetName = modifiedSourceData

            ' modifiedSourceData と sourceData の最初の2つの部分が一致する場合にのみ転記
            If Left(sourceData, Len(modifiedSourceData)) = modifiedSourceData Then
                ' シートが存在するか確認し、存在する場合のみ転記
                If InspectionSheetExists(destinationSheetName) Then
                    Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)

                    ' parts(UBound(parts))に基づいて処理を分岐
                    Select Case parts(UBound(parts))
                        Case "天"
                            If checkData = "天頂" Then
                                ' 天に関するデータ転記の処理
                                wsDestination.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            End If

                        Case "前"
                            If checkData = "前頭部" Then
                                ' 前に関するデータ転記の処理
                                wsDestination.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            End If

                        Case "後"
                            If checkData = "後頭部" Then
                                ' 後に関するデータ転記の処理
                                wsDestination.Cells(11, 4).Value = wsSource.Cells(i, 8).Value
                            End If

                        Case Else
                            ' その他の値の場合の処理（必要に応じて）
                    End Select
                End If
            End If
        End If
    Next i
End Sub

' シートが存在するかどうかを確認する関数
Function InspectionSheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    InspectionSheetExists = Not sheet Is Nothing
End Function

"C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\Scripts\USB-U_powershellExecutionPolicyBypass.bat"


以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"シートの表の値に一致するシートの各セルの値を判定することです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は"ImpactValueJudgement"にしてください。
# 条件
- "Log_Helmet"のC列の値を参照し、C列の２行目から最終行までの項目の数だけ、作業を実行します。
- C列の値と一致するシートに処理を施します。
- C列の値のうち、先頭に"F"がつく場合は、最後から2文字をその値から切り取り、残った文字とシート名を照合してください。例："F170F-Hot-天"は"F170F-Hot"となります。
- 各シートのD11の値を参照し、その値が4.9以下の場合True、それ以上の場合をFalseにします。
- 各シートのD14の値を参照し、その値が9.81以下の場合True、それ以上の場合をFalseにします。
- 各シートのD19の値を参照し、その値が9.81以下の場合True、それ以上の場合をFalseにします。
- D14,D19に値がない場合は"True"としてください。
- すべての条件がTrueの場合のみ、G9セルに"合格"と表記します。それ以外の場合は”不合格”と表記します。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。


Sub ImpactValueJudgement()
    Dim wsLog As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim resultD11 As Boolean, resultD14 As Boolean, resultD19 As Boolean

    ' "Log_Helmet"シートを設定
    Set wsLog = ThisWorkbook.Sheets("Log_Helmet")

    ' C列の最終行を取得
    lastRow = wsLog.Cells(wsLog.Rows.Count, "C").End(xlUp).Row

    ' C列の2行目から最終行までループ
    For i = 2 To lastRow
        sheetName = wsLog.Cells(i, "C").Value

        ' シート名が"F"で始まる場合、最後から2文字を削除
        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        ' 対象のシートを設定
        Set targetSheet = ThisWorkbook.Sheets(sheetName)

        ' D11, D14, D19の値を基に判定
        resultD11 = targetSheet.Range("D11").Value <= 4.9
        resultD14 = IsEmpty(targetSheet.Range("D14")) Or targetSheet.Range("D14").Value <= 9.81
        resultD19 = IsEmpty(targetSheet.Range("D19")) Or targetSheet.Range("D19").Value <= 9.81

        ' 全ての条件がTrueの場合は"合格"、それ以外は"不合格"をG9に記入
        If resultD11 And resultD14 And resultD19 Then
            targetSheet.Range("G9").Value = "合格"
        Else
            targetSheet.Range("G9").Value = "不合格"
        End If
    Next i
End Sub


' 天に関するデータ転記の処理
wsDestination.Range("B2").Value = wsSource.Cells(i, 21).Value
wsDestination.Range("E2").Value = wsSource.Cells(i, 6).Value '検査日
wsDestination.Range("G2").Value = wsSource.Cells(i, 7).Value '温度
wsDestination.Range("B3").Value = wsSource.Cells(i, 15).Value
wsDestination.Range("E3").Value = wsSource.Cells(i, 13).Value
wsDestination.Range("G3").Value = wsSource.Cells(i, 14).Value
wsDestination.Range("B4").Value = wsSource.Cells(i, 16).Value
wsDestination.Range("E4").Value = wsSource.Cells(i, 17).Value
wsDestination.Range("G4").Value = wsSource.Cells(i, 18).Value
wsDestination.Range("G7").Value = wsSource.Cells(i, 19).Value
wsDestination.Range("G8").Value = wsSource.Cells(i, 20).Value
wsDestination.Range("D11").Value = wsSource.Cells(i, 8).Value '衝撃値

' 前に関するデータ転記の処理
wsDestination.Range("D14").Value = wsSource.Cells(i, 8).Value '衝撃値
wsDestination.Range("D15").Value = wsSource.Cells(i, 11).Value
wsDestination.Range("D16").Value = wsSource.Cells(i, 10).Value
wsDestination.Range("A14").Value = "前頭部"

' 後に関するデータ転記の処理
wsDestination.Range("D19").Value = wsSource.Cells(i, 8).Value '衝撃値
wsDestination.Range("D20").Value = wsSource.Cells(i, 11).Value
wsDestination.Range("D21").Value = wsSource.Cells(i, 10).Value
wsDestination.Range("A19").Value = "後頭部"


以下の条件を満たすVBAのコードを作成してください。主な機能は"LOG_Helemt"シートの表の値に一致するシートの各セルの値を変更することです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は"AdjustHelmetTop"にしてください。
# 条件
- "Setting"のC列の値を参照し、H列の２行目から最終行までの項目の数だけ、ソートします。
- H列の値とI列の値はリンクしています。
- 該当するシートのB3の値の先頭から"_"までの文字とH列
- 各シートのD11の値を参照し、その値が4.9以下の場合True、それ以上の場合をFalseにします。
- 各シートのD14の値を参照し、その値が9.81以下の場合True、それ以上の場合をFalseにします。
- 各シートのD19の値を参照し、その値が9.81以下の場合True、それ以上の場合をFalseにします。
- D14,D19に値がない場合は"True"としてください。
- すべての条件がTrueの場合のみ、G9セルに"合格"と表記します。それ以外の場合は”不合格”と表記します。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。


Sub SetYAxis(ByRef chart As chart, ByVal maxVal As Double)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    If maxVal <= 4.95 Then
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1# '1.0刻み
    ElseIf maxVal <= 9.81 Then
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2# '2.0刻み
    Else
        yAxis.MaximumScale = Int(maxVal) + 1
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub


Sub FormatNonContinuousCells()
    ' コピーした検査票に書式を設定する。
    Dim wsLogHelmet As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmetシートを設定
    Set wsLogHelmet = ThisWorkbook.Sheets("LOG_Helmet")

    ' B列の最終行を取得
    lastRow = wsLogHelmet.Cells(wsLogHelmet.Rows.Count, "B").End(xlUp).row

    ' B列の各行をループ
    For i = 2 To lastRow
        sheetName = wsLogHelmet.Cells(i, 2).Value

        ' シート名が"F"で始まる場合、最後から2文字を削除
        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        ' ワークシートが存在するかチェック
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        ' ワークシートが存在すれば、指定したセル範囲に書式を設定
        If Not wsTarget Is Nothing Then
            ' 範囲と書式設定を関連付け
            FormatRange wsTarget.Range("G7"), "游明朝", 12, True
            FormatRange wsTarget.Range("G8"), "游明朝", 12, True
            FormatRange wsTarget.Range("G9"), "游明朝", 12, True

            ' D14に値がない場合、A14:E14とB15:D16をグレーアウト
            If IsEmpty(wsTarget.Range("D14").Value) Then
                wsTarget.Range("A14").Value = "検査対象外"
                FormatRange wsTarget.Range("A14"), "游ゴシック", 10, False
                FormatRange wsTarget.Range("A14:E14, B15:D16"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A14"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("B14:E14, B15:D16"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If

            ' D19に値がない場合、A19:E19とB20:D21をグレーアウト
            If IsEmpty(wsTarget.Range("D19").Value) Then
                wsTarget.Range("A19").Value = "検査対象外"
                FormatRange wsTarget.Range("A19"), "游ゴシック", 10, False
                FormatRange wsTarget.Range("A19:E19, B20:D21"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A19"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("B19:E19, B20:D21"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If
            FormatLastTwoCharacters wsTarget.Range("A10"), "游ゴシック", 10, True '前処理を目立たせる_書くところがないのでここに書く
            With wsTarget.Range("B2:B4, E2:E4, G2:G4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("E3").NumberFormat = "0.0"" g"""
            wsTarget.Range("G2").NumberFormat = "0"" ℃"""
            wsTarget.Range("G3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("D11, D14, D19").NumberFormat = "0.00"" kN"""
            'wsTarget.Range("D14").NumberFormat = "0.00"" kN"""
            wsTarget.Range("D15:D16, D20:D21").NumberFormat = "0.00"" ms"""
            'wsTarget.Range("D19").NumberFormat = "0.00"" kN"""
            'wsTarget.Range("D20:D21").NumberFormat = "0.00"" ms"""
            ' 他の範囲も同様に設定可能
            ' FormatRange wsTarget.Range("その他の範囲"), "フォント名", フォントサイズ, 太字かどうか, 背景色

            Set wsTarget = Nothing
        End If
    Next i
End Sub


Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "QC_試験グラフ作成試験結果_データベース.xlsm"

    ' ワークブックを開く
    Set sourceWorkbook = OpenWorkbook("", "グラフ作成用ファイル_保護帽定期試験用")
    Set targetWorkbook = OpenWorkbook(oneDrivePath, "試験結果_データベース.xlsm")

    Application.ScreenUpdating = False
    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Sub MigrateData(ByRef sourceWB As Workbook, ByRef targetWB As Workbook)
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    sourceSheets = Array("LOG_Helmet", "Setting", "Hel_SpecSheet", "InspectionSheet")
    targetSheets = Array("HeLmetTestData", "BaseBallTestData", "biCycleHelmetTestData", "FallArrestTestData")
    IDPrefixes = Array("HBT-", "BBT-", "CHT-", "FAT-")

    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWB.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWB.Sheets(targetSheets(i))

        CopyData sourceSheet, targetSheet, IDPrefixes(i)
    Next i
End Sub

Sub CopyData(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim IDGenRow As Long
    Dim currentID As String

    ' 転記元のシートの最終行と最終列を取得
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column

    ' 転記先のシートの最終行を取得
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).Row + 1

    For IDGenRow = 2 To lastRow
        ' 新しいIDを生成して転記先のシートのC列にセット
        currentID = GetNewID(targetSheet, IDPrefix)
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' 転記元から転記先へデータをコピー
        sourceSheet.Range(sourceSheet.Cells(IDGenRow, "C"), sourceSheet.Cells(IDGenRow, "U")).Copy _
            Destination:=targetSheet.Cells(targetLastRow, "D")

        ' D列から最終列までをD列から最終列へコピー
        If lastColumn > 4 Then ' 4列より多い場合のみ実行
            sourceSheet.Range(sourceSheet.Cells(IDGenRow, "D"), sourceSheet.Cells(IDGenRow, lastColumn)).Copy _
                Destination:=targetSheet.Cells(targetLastRow, "E")
        End If

        targetLastRow = targetLastRow + 1
    Next IDGenRow

    ' 転記した範囲を削除（データのみを削除）
    sourceSheet.Range(sourceSheet.Cells(2, "B"), sourceSheet.Cells(lastRow, lastColumn)).ClearContents

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & sourceSheet.Name, vbCritical
    Application.ScreenUpdating = True
End Sub

Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(name)
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Open(path & name)
    End If

    Set OpenWorkbook = wb
End Function

Function GetNewID(ByVal targetSheet As Worksheet, ByVal IDPrefix As String) As String
    Dim lastRow As Long
    Dim currentID As String
    Dim currentNumber As Integer

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).Row
    If lastRow > 1 Then
        currentID = targetSheet.Cells(lastRow, "C").Value
        currentNumber = Val(Mid(currentID, Len(IDPrefix) + 1)) + 1
    Else
        currentNumber = 1
    End If
    GetNewID = IDPrefix & Format(currentNumber, "00000")
End Function
```
------------------------------------------------------------------------------------------------------
```vb
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim selectedWb As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' ユーザーにワークブックを選択させる
    Set selectedWb = CheckAndChooseWorkbook()
    If selectedWb Is Nothing Then Exit Sub ' ユーザーが選択しないか、無効な選択をした場合は処理を中止

    Set sourceWorkbook = selectedWb

    ' OneDriveのパスを取得し、目標のパスを生成
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\QC_試験グラフ作成試験結果_データベース.xlsm"

    ' 転記先ワークブックを開く
    Set targetWorkbook = Workbooks.Open(myPath)

    ' 転記処理を実行
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Function ChooseWorkbook(wbNames As Collection) As Workbook
    Dim i As Integer
    Dim msg As String
    Dim chosenIndex As Variant

    ' 選択肢のメッセージを作成
    For i = 1 To wbNames.Count
        msg = msg & i & ". " & wbNames(i) & vbCrLf
    Next i
    msg = msg & "Please enter the number of the workbook to use:"

    ' ユーザーに選択肢を提示し、選択を受け取る
    chosenIndex = InputBox(msg, "Select Workbook")

    ' 入力値の検証
    If chosenIndex = "" Or Not IsNumeric(chosenIndex) Or chosenIndex < 1 Or chosenIndex > wbNames.Count Then
        MsgBox "Invalid selection. Process aborted.", vbCritical
        Set ChooseWorkbook = Nothing
    Else
        ' 選択されたワークブックを返す
        Set ChooseWorkbook = Workbooks(wbNames(chosenIndex))
    End If
End Function

Sub CheckAndChooseWorkbook()
    Dim wb As Workbook
    Dim wbNames As New Collection
    Dim selectedWb As Workbook

    ' 条件に一致するワークブックの名前を集める
    For Each wb In Application.Workbooks
        If wb.Name Like "グラフ作成用ファイル_*.xlsm" Then
            wbNames.Add wb.Name
        End If
    Next wb

    ' 複数のワークブックが見つかった場合、選択肢を提示
    If wbNames.Count > 1 Then
        Set selectedWb = ChooseWorkbook(wbNames)
        If Not selectedWb Is Nothing Then
            MsgBox "Selected: " & selectedWb.Name
            ' ここで選択されたワークブックに対する処理を実行
        End If
    ElseIf wbNames.Count = 1 Then
        Set selectedWb = Workbooks(wbNames(1))
        MsgBox "Automatically selected the only matching workbook: " & selectedWb.Name
        ' ここで唯一のワークブックに対する処理を実行
    Else
        MsgBox "No valid workbook found matching the pattern. Process aborted.", vbCritical
    End If
End Sub

Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    'CreateIndividualChartのサブプロシージャ
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.ChartTitle.text = ws.Cells(i, "B").Value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 0.75

    SetYAxis chart, maxVal
    SetXAxis chart

End Sub

Sub SetYAxis(ByRef chart As chart, ByVal maxVal As Double)
    'ConfigureChartで使用する関数
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    If maxVal <= 4.95 Then
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1# '1.0刻み
    ElseIf maxVal <= 9.81 Then
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2# '2.0刻み
    Else
        yAxis.MaximumScale = Int(maxVal) + 1
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.color = RGB(89, 89, 89)
        .Font.Size = 8
    End With
End Sub
以上のコードはグラフの書式設定をするVBAのコードです。Y軸の目盛りを決定するのに、現在は数値を使用していますが、これをE軸の値に変更したい。
以下の条件に沿うようにコードの変更をお願いします。
# 条件
- グラフデータがある行のE列の値によって最大値を決める。
- 値が"天頂"の場合はyAxis.MaximumScale = 5 yAxis.MajorUnit = 1#
- そうでない場合はyAxis.MaximumScale = 10 yAxis.MajorUnit = 2#
- 行はchart.ChartTitle.text = ws.Cells(i, "B").Value に使用されている変数"i"と同値である。



Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"
    Debug.Print myPath

    ' ワークブックを開く
    Set sourceWorkbook = OpenWorkbook("", "グラフ作成用ファイル_保護帽定期試験用.xlsm")
    Set targetWorkbook = OpenWorkbook(oneDrivePath, "試験結果_データベース.xlsm")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub
Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(name)
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Open(path & name)
    End If

    Set OpenWorkbook = wb
End Function

これがOpenWorkbookのコードです。
エラーの内容は実行時エラー"1004"で、
申し訳ございません。C:\Users\QC07\OneDrive - トーヨーセフティーホールディングス株式会社試験結果_データベース.xlsmが見つかりません。名前が変更されたか、移動や削除が行われた可能性があります。
ですが、正しいパスは"C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\試験結果_データベース.xlsm"です。


同じエラーが出るのでこのようにしてみました。イミディエイトウィンドウの結果も載せます。
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"
    Debug.Print "myPath" & myPath

    ' ワークブックを開く
    Set sourceWorkbook = OpenWorkbook("", "グラフ作成用ファイル_保護帽定期試験用.xlsm")
    Set targetWorkbook = OpenWorkbook(oneDrivePath, "試験結果_データベース.xlsm")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    ' パスとファイル名の結合時に区切り文字を確実に挿入
    fullPath = path & "\" & name
    Debug.Print "fullPath" & fullPath

    On Error Resume Next
    Set wb = Workbooks(name)
    On Error GoTo 0

    If wb Is Nothing Then
        ' fullPath変数を使用してファイルを開く
        Set wb = Workbooks.Open(fullPath)
    End If

    Set OpenWorkbook = wb
End Function

# イミディエイトウィンドウ
myPathC:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\試験結果_データベース.xlsm
fullPath\グラフ作成用ファイル_保護帽定期試験用.xlsm
fullPathC:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\試験結果_データベース.xlsm


' ワークブックを開く
' 正しいパスを指定してsourceWorkbookを開くためのコードをここに追加
' 例: Set sourceWorkbook = OpenWorkbook(correctSourcePath, "グラフ作成用ファイル_保護帽定期試験用.xlsm")

' myPathを使用してtargetWorkbookを開く
Set targetWorkbook = OpenWorkbook(myPath, "")


Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    If name = "" Then
        fullPath = path
    Else
        fullPath = path & "\" & name
    End If
    Debug.Print "fullPath" & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function


Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"

    ' sourceWorkbookを開くための正しいパスをここに設定する
    ' この例では、sourceWorkbookがどこにあるか具体的なパスが指定されていないため、適切なパスを指定する必要があります
    Set sourceWorkbook = OpenWorkbook("ここにsourceWorkbookの正確なパス", "グラフ作成用ファイル_保護帽定期試験用.xlsm")

    ' myPathを使用してtargetWorkbookを開く
    Set targetWorkbook = OpenWorkbook(myPath, "")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub


"C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\試験結果_データベース.xlsm"

"C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\グラフ作成用ファイル_保護帽定期試験用.xlsm"


Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"

    ' sourceWorkbookを開く
    Set sourceWorkbook = OpenWorkbook("C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\", "グラフ作成用ファイル_保護帽定期試験用.xlsm")

    ' myPathを使用してtargetWorkbookを開く
    Set targetWorkbook = OpenWorkbook(myPath, "")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub
```

それではユーザーに必要なグラフを選択してもらいましょう。リストを記したシートのD列、E列、L列を選択し、当てはまる行のグラフを作るための工程を示してください。ユーザフォームが必要ならできる範囲でそれを示してください。



以下の条件を満たすVBAのコードを作成してください。コードの目的は列の値から同値を見つけることです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "Hel_SpecSheet"のH列の2行目から最終行までの値を探索する
- その中から同じ数値を見つける。
- 同じ数値のセルを同じ色に塗る。
- 色はその数値のペアごとに変更してください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。
```VB

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

    ' 最終列を"T"に固定
    Dim lastCol As Long
    lastCol = ws.Columns("T").column

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
                logSheet.Cells(i, "B").Value = helSpec.Cells(j, "B").Value
                logSheet.Cells(i, "C").Value = helSpec.Cells(j, "B").Value
                logSheet.Cells(i, "D").Value = helSpec.Cells(j, "D").Value
                logSheet.Cells(i, "E").Value = helSpec.Cells(j, "E").Value
                logSheet.Cells(i, "F").Value = helSpec.Cells(j, "F").Value
                logSheet.Cells(i, "G").Value = helSpec.Cells(j, "G").Value
                logSheet.Cells(i, "L").Value = helSpec.Cells(j, "I").Value
                logSheet.Cells(i, "M").Value = helSpec.Cells(j, "J").Value
                logSheet.Cells(i, "N").Value = helSpec.Cells(j, "K").Value '天頂すきま
                logSheet.Cells(i, "O").Value = helSpec.Cells(j, "L").Value
                logSheet.Cells(i, "U").Value = helSpec.Cells(j, "M").Value '試験内容
                logSheet.Cells(i, "P").Value = helSpec.Cells(j, "N").Value '製造ロット
                logSheet.Cells(i, "Q").Value = helSpec.Cells(j, "O").Value
                logSheet.Cells(i, "R").Value = helSpec.Cells(j, "P").Value
                logSheet.Cells(i, "S").Value = helSpec.Cells(j, "Q").Value '構造結果
                logSheet.Cells(i, "T").Value = helSpec.Cells(j, "R").Value
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
            ' D列の2行目から最終行までの書式設定を追加
            If cell.column = 4 Then  ' D列を識別
                Set rng = ws.Range(ws.Cells(2, cell.column), ws.Cells(ws.Rows.Count, cell.column).End(xlUp))
                rng.NumberFormat = """No.""@"
                rng.Value = Evaluate("IF(ROW(" & rng.Address & ")>1,""No."" & " & rng.Address & ","""")")
            End If
        Next cell
    Next sheet
End Sub
```

以下の条件を満たすVBAのコードを作成してください。コードの目的は所定の時間になると通知することです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は、作ったコードの機能を簡略にして英語で表現してください。
# 条件
- 現在の西暦年とE列の値で年月を作成します。E列の値が2なら"2024/02/01"となります。
- その時刻になると"ryoji.yamanaka@tbshldgs.jp"にメールがいくように。
- "検査実施月"シートのE列の2行目から最終行までを探索し、当てはまるようなら実行する
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。


|  A       |  B   |  C               |
|----------|------|----------------- |
| 製品名    | 検査月| 作成班            |
|----------|------|----------------- |
| No.100   |   2  | hogo2@tshldgs.jp |
| No.105   |   2  | hogo3@tshldgs.jp |
| No.110   |   4  | hogo3@tshldgs.jp |
| No.110F  |   4  | hogo4@tshldgs.jp |
| No.110S  |   7  | hogo2@tshldgs.jp |
------------------------------------------------------------


以下の条件を満たすVBAのコードを作成してください。コードの目的は所定の時間になると[HelmetCollection]シートから
[RequestTable]シートに表を作成することです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は、作ったコードの機能を簡潔にして英語で表現してください。
# 条件
- 現在の西暦年とE列の値で年月を作成します。B列の値が2なら"2024/02/01"となります。
- 表の中からB列の値が[2]のものを抜き出し、[RequestTable]に転記してください。
- 転記位置は、[HelmetCollection]シートのA列の値を[RequestTable]シートのB列に。
- B列で同じ行に転記しないようにしてください。また１行目はタイトルなのでそのままにしてください。
- [HelmetCollection]シートのA列の2行目から最終行までを探索し,実行してください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

[HelmetCollection]シートにC列が存在します。そのC列の値によって転記されるシートを変更したいです。
値が[2]と[3]なら[RequestTable]シートを複製し、転記作業を行ってください。
また、シートの名前は[保護帽班_2],[保護帽班_3]のようにしてください。

[Setting]シートのH列と[Hel_SpecSheet]のD列を比較し、合致する行の[Setting]シートのI列の値を[Hel_SpecSheet]のG列に記入するコードを書いてください。

```VB
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"

    ' sourceWorkbookを開く
    Set sourceWorkbook = OpenWorkbook("C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\", "グラフ作成用ファイル.xlsm")

    ' myPathを使用してtargetWorkbookを開く
    Set targetWorkbook = OpenWorkbook(myPath, "")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Sub MigrateData(ByRef sourceWB As Workbook, ByRef targetWB As Workbook)
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    sourceSheets = Array("LOG_Helmet")
    targetSheets = Array("HeLmetTestData", "BaseBallTestData", "biCycleHelmetTestData", "FallArrestTestData")
    IDPrefixes = Array("HBT-", "BBT-", "CHT-", "FAT-")

    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWB.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWB.Sheets(targetSheets(i))

        CopyData sourceSheet, targetSheet, IDPrefixes(i)
    Next i
End Sub

Sub CopyData(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim IDGenRow As Long
    Dim currentID As String

    ' 転記元のシートの最終行と最終列を取得
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).column

    ' 転記先のシートの最終行を取得
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row + 1

    For IDGenRow = 2 To lastRow
        ' 新しいIDを生成して転記先のシートのC列にセット
        currentID = GetNewID(targetSheet, IDPrefix)
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' 転記元から転記先へデータをコピー
        sourceSheet.Range(sourceSheet.Cells(IDGenRow, "C"), sourceSheet.Cells(IDGenRow, "U")).Copy _
            Destination:=targetSheet.Cells(targetLastRow, "D")

        ' D列から最終列までをD列から最終列へコピー
        If lastColumn > 4 Then ' 4列より多い場合のみ実行
            sourceSheet.Range(sourceSheet.Cells(IDGenRow, "D"), sourceSheet.Cells(IDGenRow, lastColumn)).Copy _
                Destination:=targetSheet.Cells(targetLastRow, "E")
        End If

        targetLastRow = targetLastRow + 1
    Next IDGenRow

    ' 転記した範囲を削除（データのみを削除）
    sourceSheet.Range(sourceSheet.Cells(2, "B"), sourceSheet.Cells(lastRow, lastColumn)).ClearContents

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description & " in " & sourceSheet.name, vbCritical
    Application.ScreenUpdating = True
End Sub

Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    If name = "" Then
        fullPath = path
    Else
        fullPath = path & "\" & name
    End If
    Debug.Print "fullPath" & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

Function GetNewID(ByVal targetSheet As Worksheet, ByVal IDPrefix As String) As String
    Dim lastRow As Long
    Dim currentID As String
    Dim currentNumber As Integer

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row
    If lastRow > 1 Then
        currentID = targetSheet.Cells(lastRow, "C").Value
        currentNumber = Val(Mid(currentID, Len(IDPrefix) + 1)) + 1
    Else
        currentNumber = 1
    End If
    GetNewID = IDPrefix & Format(currentNumber, "00000")
End Function
```
# 依頼試験レポートの自動化
依頼試験資料のSpecSheetを記入。
SpecSheetを「グラフ作成用シート.xlsm」にインポート
「グラフ作成用シート.xlsm」で試験結果からResultsSheetを作成
ResultsSheetから試験票を作成
また、ResultsSheetからレポートの下書きを作成
レポートに必要事項を書き込む
試験票とレポートをPDFで生成

グラフ作成用ファイル.xlsm,グラフ作成用ファイル_ヘルメット依頼試験.xlsm,グラフ作成用ファイル_保護帽定期試験用.xlsmの3つのファイルで同じようなマクロを使用しています。ただ、ファイル名などの関係で少しずつカスタムしています。マクロの変更があった場合に3つのコードを変更するのは骨が折れるのでなにかいい方法はありませんか？


試しにグラフ作成用ファイル.xlsmのマクロをグラフ作成用ファイル_ヘルメット依頼試験.xlsmで使用するにはどうすればいいでしょうか？工程を示してください。

グラフ作成用ファイル.xlsmの[LOG_Helmet]シートから試験結果_データベース.xlsmの[HelmetTestData]にデータを転記したいと思います。[LOG_Helmet]シートの行にはA列からCFD列まであり、そのうち必要なのはBN列からAGL列までです。
この場合転記元の列を整理してから転記作業を始めたほうがいいのか、それとも転記してから列を整理したほうがいいのかどちらがいいでしょうか？

```vb
Sub TransferDataUsingArray()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim data As Variant

    ' ソースとターゲットのワークシートを設定
    Set sourceSheet = Workbooks("グラフ作成用ファイル.xlsm").Sheets("LOG_Helmet")
    Set targetSheet = Workbooks("試験結果_データベース.xlsm").Sheets("HelmetTestData")

    ' ソースシートで転記するデータの最終行を見つける
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "BN").End(xlUp).Row

    ' 転記するデータ範囲を設定
    Set dataRange = sourceSheet.Range("BN1:AGL" & lastRow)

    ' データ範囲を配列に読み込む
    data = dataRange.Value

    ' 配列のデータをターゲットシートに転記
    targetSheet.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
End Sub

このコードに以下の機能を追加してください。ブックを閉じるときに"LOG_Helmet","LOG_FallArrest","LOG_Bicycle","LOG_Baseball"のRange("B2:" lastColumn & lastRow)の範囲のデータを消す。

それでは、試験予定テーブル から 試料データテーブル を自動作成するVBAのマクロを作成したいと思います。あなたはプロのVBAコーダーであり、テスト環境下での開発ですのである程度の失敗は許容されます。以下の条件に従うVBAコードを記述してください。
# 条件
- 試験予定テーブルから一意の識別子を作成する。
- 作成した識別子をもとにデータテーブルに識別子を転記する。
- 試料の数は同じ識別子に対し、3個必要。
- 試料データテーブルの設計は前回にあなたが行ったものを参考にしてください。

一意の識別子の形状にもう少し工夫の余地がありそうです。まずは識別子作成のプロシージャを作成し、メインのプロシージャと分離しましょう。

試験予定テーブルでは作成試料が記述されていますが、その欄に試験済み欄を設けたいです。そして試験済み欄にチェックが入っている試料は試料データテーブルに含まれないようにしたい。

Sub CreateSampleDataTable()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRowSource As Long, lastRowTarget As Long, r As Long
    Dim uniqueIdentifier As String
    Dim testCompleted As String ' 試験済み欄をチェックするための変数

    ' ソースとターゲットのワークシートを設定
    Set wsSource = ThisWorkbook.Sheets("試験予定")
    Set wsTarget = ThisWorkbook.Sheets("試料データ")

    ' ソースワークシートの最終行とターゲットワークシートの次の空行を取得
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

    ' 試験予定テーブルの各行をループ
    For r = 2 To lastRowSource ' 1行目はヘッダーと仮定してスキップ
        ' 試験済み欄を確認（H列にあると仮定）
        testCompleted = wsSource.Cells(r, 8).Value

        ' 試験済みでない場合にのみ試料データテーブルに追加
        If testCompleted <> "〇" Then ' 〇または他のマーカーがない場合に処理
            For sampleNumber = 1 To 3
                uniqueIdentifier = CreateUniqueIdentifier(wsSource.Cells(r, 1).Value, _
                                                          wsSource.Cells(r, 3).Value, _
                                                          wsSource.Cells(r, 4).Value, _
                                                          sampleNumber)
                With wsTarget
                    .Cells(lastRowTarget, 1).Value = uniqueIdentifier
                    .Cells(lastRowTarget, 2).Value = "TBD" ' 重量
                    .Cells(lastRowTarget, 3).Value = "TBD" ' 天頂サイズ
                    .Cells(lastRowTarget, 4).Value = "TBD" ' 前処理の種類
                    .Cells(lastRowTarget, 5).Value = wsSource.Cells(r, 4).Value ' 製造ロット
                    .Cells(lastRowTarget, 6).Value = wsSource.Cells(r, 5).Value ' 帽体ロット
                    .Cells(lastRowTarget, 7).Value = wsSource.Cells(r, 6).Value ' 内装ロット
                End With
                lastRowTarget = lastRowTarget + 1
            Next sampleNumber
        End If
    Next r

    MsgBox "試料データテーブルの作成が完了しました。", vbInformation
End Sub
```

- 出荷統計と技術委員会などの関係。

以下の条件を満たすVBAのコードを作成してください。コードの目的は列の値から同値を見つけることです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "LOG_Helmet"のB列の2行目から最終行までの値を選択する
- 選択したエリアを削除する。
- 1行目は削除しない。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

以下の条件を満たすExcelの関数を作成してください。Office365のExcelを使用しています。
# 条件
- 'Setting'シートのH2からI47を参照する。
- 'Hel_SpecSheet'のD列と'Setting'シートのH2からH47列を比較し、同じ値なら'Setting'シートのI列の同じ行の値を'Hel_SpecSheet'のT列に代入する。
- Xlookup関数を使用する。

以下の条件を満たすVBAのコードを作成してください。コードの目的はある列にExcelの関数を代入することです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "Hel_SpecSheet"のD列の2行目から最終行までを探索し、値がある行を探索します。
- その行のT列に'=XLOOKUP(Hel_SpecSheet!D2, Setting!$H$2:$H$47, Setting!$I$2:$I$47, "")'を代入してください。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

ありがとうございます。コードは機能しました。このコードに以下の条件を加えたいです。
# 条件
- K列の値からT列の値を引いた値を再びK列に代入する。
- K列の値が操作されたことを明確にするため、操作した行のU列に'Changed'を代入する。


以下の条件を満たすExcelの関数を作成してください。Office365のExcelを使用しています。
# 条件
- '試料の製作予定'シートのB列の値を参照する。
- 'HelmetCollection'シートのA列と'試料の製作予定'シートのB列を比較し、同じ値なら'HelmetCollection'シートのB列の同じ行の値を'試料の製作予定'シートのC列に代入する。
- Xlookup関数を使用する。

作っていただいた"=DATE(2024, XLOOKUP(B2, HelmetCollection!A:A, HelmetCollection!B:B, ""), 1)"ですが、これとは別に半年後の日数も作りたいです。この関数で作成した半年後の日数を作成する関数を教えて下さい。

以下の条件を満たすVBAのコードを作成してください。コードの目的はシートに一覧表を完成させることです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "試料の製作予定"シートのB列の2行目から"HelmetCollection"シートのA列の2行目から最終行までの値をコピペします。
- さらに"試料の製作予定"シートのB列の最終行までを探索し、その最終行の次の行から"HelmetCollection"シートのA列の2行目から最終行までの値をコピペします。
- "試料の製作予定"シートのB列の2行目から最終行までを探索し、値がある行を探索します。
- "試料の製作予定"シートのB列に'=DATE(2024, XLOOKUP(B2, HelmetCollection!A:A, HelmetCollection!B:B, ""), 1)'を代入してください。
- "試料の製作予定"シートのD列に'=XLOOKUP(B2,HelmetCollection!A:A,HelmetCollection!D:D,"")'を代入してください。
- B"2"の2は各行に対応したい。
- DATE(2024, の2024は今年の年数です。関数を実行したタイミングの年数にしたい。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。

```vb
wsSchedule.Cells(i, "B").Formula = "=DATE(" & yearNow & ", XLOOKUP(B" & i & ", HelmetCollection!A:A, HelmetCollection!B:B, """")" & ", 1)"
ですが、"HelmetCollection"シートのA列からのデータの2回目からのコピペの最終行は以下のように変えてほしいです。
=EDATE(DATE(2024, XLOOKUP(B24, HelmetCollection!A:A, HelmetCollection!B:B, ""), 1), 6)

Sub CompleteSampleScheduleUpdated()
    Dim wsSchedule As Worksheet, wsCollection As Worksheet
    Dim lastRowSchedule As Long, lastRowCollection As Long, i As Long
    Dim yearNow As Integer, monthNum As Integer, targetDate As Date
    Dim lookupResult As Variant

    Set wsSchedule = ThisWorkbook.Sheets("試料の製作予定")
    Set wsCollection = ThisWorkbook.Sheets("HelmetCollection")

    yearNow = Year(Date)
    lastRowCollection = wsCollection.Cells(wsCollection.Rows.Count, "A").End(xlUp).Row
    lastRowSchedule = wsSchedule.Cells(wsSchedule.Rows.Count, "B").End(xlUp).Row + 1

    For i = 2 To lastRowCollection - 1
        ' 簡易的なXLOOKUPの代替。適切な値を見つけるためのロジックが必要です。
        ' ここでは、単純化のために直接月をセットしています。
        ' 実際の使用では、この部分をデータに基づいて適切な値を見つけるロジックに置き換えてください。
        monthNum = wsCollection.Cells(i, "B").Value
        If IsNumeric(monthNum) Then
            ' monthNumが有効な場合に限り、日付計算を実行
            targetDate = DateSerial(yearNow, monthNum, 1) ' DATE関数の代わり
            targetDate = DateAdd("m", 6, targetDate) ' 6ヶ月後の日付を計算
            wsSchedule.Cells(lastRowSchedule + i - 2, "C").Value = targetDate ' C列にセット
        End If
    Next i
End Sub
wsSchedule.Cells(i, "C").Formula = "=EDATE(DATE(" & yearNow & ", XLOOKUP(B" & i & ", HelmetCollection!A:A, HelmetCollection!B:B, ""), 1), 6)"
wsSchedule.Cells(i, "C").Formula = "=EDATE(DATE(" & yearNow & ", XLOOKUP(B" & i & ", HelmetCollection!A:A, HelmetCollection!B:B, """")" & ", 1)" & ", 6)"

wsSchedule.Cells(i, "C").Formula = "=DATE(" & yearNow & ", XLOOKUP(B" & i & ", HelmetCollection!A:A, HelmetCollection!B:B, """")" & ", 1)"
```
以下の条件を満たすVBAのコードを作成してください。コードの目的はシートのある範囲を値を消すことです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "試料の製作予定"シートのB列の2行目から最終行までを探索する。
- 探索した行のZ列までの範囲を消す。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。


myPath = oneDrivePath & "\QC_試験グラフ作成\CSV\"
myPath: C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\CSV\
"C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\image\inkan_pass.png"

myPath = oneDrivePath & "QC_試験グラフ作成\image\"

以下の条件を満たすVBAのコードを作成してください。コードの目的はシートのある範囲を値を消すことです。
あなたはプロのVBAコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
プロシージャの名前は英語で目的に沿うものにしてください。
# 条件
- "試料の製作予定"シートのB列の2行目から最終行までを探索する。
- 探索した行のZ列までの範囲を消す。
- 変数やプロシージャの名前は英語で、コメントは日本語でお願いします。
- シート名は後で変更できるようにしてください。


チャートの上部にはB列の各行の値がタイトルとしてついています。工程としては
1. そのタイトルを削除。
2. 行とワークシートがB列の各業の値で対応しているので、対応したチャートを各シートに画像として貼り付け。
3. 貼り付けた画像の位置を調整
というふうに考えています。懸念点は先のコードで'MatchName'と'wsDestination'のマッチを成立させているのでそれと組み合わせることができるかという点です。

```vb
Dim chartObj As ChartObject
For Each chartObj In wsSource.ChartObjects
    If Not chartObj.Chart.HasTitle Then
        ' タイトルがなければスキップ
    Else
        chartObj.Chart.ChartTitle.Delete
    End If
Next chartObj

For Each wsDestination In ThisWorkbook.Sheets
    dashPosDest = InStr(wsDestination.Name, "-")
    If dashPosDest > 0 Then
        If Left(wsDestination.Name, dashPosDest - 1) = matchName And InStr(wsDestination.Name, temperatureCondition) > 0 Then
            ' 条件に一致するワークシートに対して、適切なチャートをコピー＆ペースト
            ' コピーするチャートオブジェクトを選択（この部分はロジックの追加が必要）
            chartObj.Copy

            ' 貼り付け先のワークシートをアクティブにしてペースト
wsDestination.Activate
wsDestination.Paste
' 貼り付けた後、画像の位置を調整する処理を追加
' この部分はステップ3で詳述
End If
End If
Next wsDestination

With wsDestination.Shapes(wsDestination.Shapes.Count)
    .Left = wsDestination.Cells(1, 1).Left  ' 例えばA1セルの左端に合わせる
    .Top = wsDestination.Cells(1, 1).Top  ' 例えばA1セルの上端に合わせる
    .Width = 400 ' 必要に応じて調整
    .Height = 300 ' 必要に応じて調整
End With
```



画像を貼り付けた後、その位置やサイズを適切に調整する必要があります。Excelの`Shapes`コレクションの最後の要素（最新で貼り付けたオブジェクト）を参照し、そのプロパティを設定して位置を調整します。

コードは途中まで実行され、実行時エラー-2147221040(800401d0)が出た後に止まりました。chartObj.Copyの部分です。
チャートは途中までペーストされていました。

D列の値に"F"が含まれている場合は、IDの先頭に"F"を追加します。
この条件を以下のように変更したい。コードを示してください。

# 条件
- 'Hel_SpecSheet'シートD列と'Setting'シートのH列を照合する。
- 照合すれば、'Setting'シートのJ列を参照し、そこに'x'が含まれていれば、IDの先頭に"F"を追加します。
- 'Setting'シートのJ列を参照し、そこに'x'が含まれていなければ、IDは変更しない。
- 照合しなければエラーを返し、処理を中止する。
```vb
Sub CreateInspectionSheetIDs()
    Dim wsSpecSheet As Worksheet
    Set wsSpecSheet = ThisWorkbook.Sheets("Hel_SpecSheet")

    Dim wsSetting As Worksheet
    Set wsSetting = ThisWorkbook.Sheets("Setting")

    Dim lastRow As Long
    lastRow = wsSpecSheet.Cells(wsSpecSheet.Rows.Count, "D").End(xlUp).Row

    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    For i = 2 To lastRow
        If wsSpecSheet.Cells(i, "D").Value <> "" Then
            wsSpecSheet.Cells(i, "S").Formula = "=IF(INDIRECT(""R" & i & "C9"", FALSE)=""高温"", ""Hot"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""低温"", ""Cold"", IF(INDIRECT(""R" & i & "C9"", FALSE)=""浸せき"", ""Wet"", """")))"
            Dim id As String
            id = wsSpecSheet.Cells(i, "D").Value & "-" & wsSpecSheet.Cells(i, "S").Value & "-" & Left(wsSpecSheet.Cells(i, "E").Value, 1)

            foundMatch = False
            For j = 2 To wsSetting.Cells(wsSetting.Rows.Count, "H").End(xlUp).Row
                If wsSpecSheet.Cells(i, "D").Value = wsSetting.Cells(j, "H").Value Then
                    foundMatch = True
                    If InStr(wsSetting.Cells(j, "J").Value, "x") > 0 Then
                        id = "F" & id
                    End If
                    Exit For
                End If
            Next j

            If Not foundMatch Then
                MsgBox "エラー: D列の値がSettingシートのH列と一致する項目がありません。処理を中止します。"
                Exit Sub
            End If

            wsSpecSheet.Cells(i, "B").Value = id
            wsSpecSheet.Cells(i, "Q").Value = "合格"
            wsSpecSheet.Cells(i, "R").Value = "合格"
        End If
    Next i
End Sub

Sub NameMerge()
    Dim k As Long
    Dim lastRow As Long
    Dim mergeCells As Long
    Dim myRng As Range
    Dim lotVol As String
    Dim lotPrc As String
    Dim lotJoin As String

    Sheets("Print").Activate

    lastRow = 0
    mergeCells = 0
    With Sheets("Print").Cells(1).CurrentRegion
        For lastRow = .rows.count To 2 Step -1


'            ロット入力が完成したときに付け加える
'            lotJoin = Join(.Cells(lastRow, 5) & .Cells(lastRow, 6))
'            lotJoin = .Cells(lastRow, 6)

                If .Cells(lastRow, 4).value = .Cells(lastRow - 1, 4).value Then
                    mergeCells = 1
                    Exit For
                Else
                End If
        Next
    End With

    If mergeCells > 0 Then
        With Sheets("Print").Cells(1).CurrentRegion

'上下のセル(k,6)を比較し、セル内容が同じなら下のセルを消す。
            For k = .rows.count To 2 Step -1
                If .Cells(k, 1).value = .Cells(k - 1, 1).value Then
                    If .Cells(k, 4).value = .Cells(k - 1, 4).value Then
                        .Cells(k, 4).ClearContents
                    End If
                Else
                End If
            Next

'分類番号と商品名のセルをくっつける。

            On Error Resume Next 'エラーが出た行を無視する。
            For Each myRng In .Columns(4).SpecialCells(xlCellTypeBlanks).Areas '←伊藤忠のエラーが起こる
                With myRng.Offset(-1).Resize(myRng.count + 1)
                     .Merge
                    lotVol = Join(Application.Transpose(.Columns(4)), vbLf)
                    lotPrc = Join(Application.Transpose(.Columns(5)), vbLf)
                    .Columns(4).ClearContents
                    .Columns(4).Merge
                    .Columns(4).value = lotVol
                    .Columns(5).ClearContents
                    .Columns(5).Merge
                    .Columns(5).value = lotPrc
                End With
            Next

        End With
    End If

 End Sub
```
ExcelでH列の値を別シートに転記したい。
1. C列の値が天頂、前頭部、後頭部、側頭部
2. D列の値が-、30、45
3. C,D列の値の組み合わせで転記するシートが決まります。
4. E列の値が高温、低温、浸せきで転記する列が決まります。
5. F列の値が1,2,3で転記する行が決まります。

これらの条件を段階的に行うことでH列の値が各シートの所定の位置に決まるように
ロジックを組みたいです。サンプルコードを示してください。

Traceback (most recent call last):
  File "C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\Scripts\importCSV.py", line 14, in <module>
    header = pd.read_csv(os.path.join(directory, files[0]), nrows=0)  # 最初のファイルのヘッダーを読み込み
             ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\QC07\AppData\Local\Programs\Python\Python312\Lib\site-packages\pandas\io\parsers\readers.py", line 1026, in read_csv
    return _read(filepath_or_buffer, kwds)
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\QC07\AppData\Local\Programs\Python\Python312\Lib\site-packages\pandas\io\parsers\readers.py", line 620, in _read
    parser = TextFileReader(filepath_or_buffer, **kwds)
             ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\QC07\AppData\Local\Programs\Python\Python312\Lib\site-packages\pandas\io\parsers\readers.py", line 1620, in __init__
    self._engine = self._make_engine(f, self.engine)
                   ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\QC07\AppData\Local\Programs\Python\Python312\Lib\site-packages\pandas\io\parsers\readers.py", line 1898, in _make_engine
    return mapping[engine](f, **self.options)
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\QC07\AppData\Local\Programs\Python\Python312\Lib\site-packages\pandas\io\parsers\c_parser_wrapper.py", line 93, in __init__
    self._reader = parsers.TextReader(src, **kwds)
                   ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "parsers.pyx", line 574, in pandas._libs.parsers.TextReader.__cinit__
  File "parsers.pyx", line 663, in pandas._libs.parsers.TextReader._get_header
  File "parsers.pyx", line 874, in pandas._libs.parsers.TextReader._tokenize_rows
  File "parsers.pyx", line 891, in pandas._libs.parsers.TextReader._check_tokenize_status
  File "parsers.pyx", line 2053, in pandas._libs.parsers.raise_parser_error
  File "<frozen codecs>", line 322, in decode
UnicodeDecodeError: 'utf-8' codec can't decode byte 0x8b in position 5: invalid start byte

CSVファイルを以下の条件で転記したい。Pythonのコードを示してください。
あなたはプロのPythonコーダーです。バックアップ環境は完全ですので失敗を恐れずチャレンジしてください。
コード全体の名前は英語で目的に沿うものにしてください。
# 条件
- B列以降の列から最終行までを順番に行に変換する。
- 転記元は52行目から最後の行までを転記先の68列目から並べ直す。



import os
import pandas as pd
from datetime import datetime
import shutil

# USBドライブからCSVファイルのパスを指定
directory = 'T:\\HIOKI8847'  # USBドライブのパスを適宜調整
files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

# 日付を取得してファイル名を生成
today = datetime.now().strftime('%Y-%m-%d')
base_filename = f'{today}_TestLOG.CSV'

# 保存先とコピー先の基本パスを設定
save_directory = 'C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\CSV\\'
copy_directory = 'C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\CSV_LOG\\'
上記はPythonコードの一部です。USBドライブからCSVファイルを抜き出すのですが、ドライブ名が S, U, Tの3種類あります。
どのドライブでも対応できるようにシたいのですが可能でしょうか？


input_pathにて入力ファイルを指定しているのですが、このフォルダの中には複数の変換すべきCSVファイルが含まれています。


C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\CSV\\
ファイルパスの指定の仕方ですが、この指定の仕方では他の環境では使えません。いいアプローチは有りませんか？

実行しているディレクトリは
C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\scripts\\
ですが、
save_directoryは
C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\
copy_directoryは
C:\\Users\\QC07\\OneDrive - トーヨーセフティホールディングス株式会社\\QC_試験グラフ作成\\CSV_LOG\\
です。

```Python
input_path = r'C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\CSV\2024-04-19_TestLOG.CSV'
output_path = r'C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\OutputTest.csv'

import os

# スクリプトが実行されているディレクトリの取得
scripts_directory = os.getcwd()

# scriptsディレクトリの親ディレクトリに移動して、基本パスを取得
base_directory = os.path.abspath(os.path.join(scripts_directory, os.pardir))

# 保存先ディレクトリとコピー先ディレクトリのパスを設定
save_directory = os.path.join(base_directory, 'CSV')
copy_directory = base_directory
```
input_pathにて入力ファイルを指定しているのですが、このフォルダの中には複数の変換すべきCSVファイルが含まれています。

input_filepath = os.path.join(base_directory, 'CSV', '2024-04-19_TestLOG.CSV')の1行目はヘッダーになっています。
それをoutput_path = r'C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\OutputTest.csv'の
B列に転記したい。コードの修正をお願いできますか？

# 条件
- Y軸の最大値は maxVal の値を10の位で、50単位で繰り上げてください。123 なら150、173なら200となるように

以上のコードの"ID = ID & Mid(ws.Cells(rowIndex, 4).value, 4, 3)"の部分を改変したい。まずは変更部分のロジックを説明してください。
# 条件
- このセルの値は"No.396", "No.3220F”のようになっています。まずは数字部分を抜き出し、値にFがついている場合はID生成の際に数字の前と後ろにFをつけたい。

```vb
' SpecSheetに転記するプロシージャの本体。アイコンに紐づけ。
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
                    ' 数値に変換
                    If IsNumeric(CDbl(cell.Value)) Then
                        cell.Value = CDbl(cell.Value)
                    Else
                        cell.Value = 0
                    End If
                    cell.Interior.ColorIndex = 6 ' 黄色に色付け
                    errorMsg = errorMsg & "数値に変換したセル: " & cell.Address(False, False) & vbNewLine
                End If
            End If

            ' 列N、O、Pで文字列の確認
            If j = Columns("N").column Or j = Columns("O").column Or j = Columns("P").column Then
                If Not VarType(cell.Value) = vbString Then
                    ' 文字列に変換
                    cell.Value = CStr(cell.Value)
                    cell.Interior.ColorIndex = 6 ' 黄色に色付け
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


LOG_HelmetのC列の値に01-396-天Hot-W,04-396F-天Cold-Wなどの値が入っています。これを
396-天Hot-W
F396F-天Cold-W
に変更するようなVBAコードがほしいです。396など3桁~5桁の数字が連続して並んでいる値のあとにFがつく場合のみ頭にFをつけたIDを生成するようにしてください。


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
End Sub




' クラスモジュール Record
Public ID As String
Public Name As String
Public Value As Double

' クラスモジュール CellManager
Public Cell As Range

Public Sub SetValue(rec As Record)
    ' Recordオブジェクトのデータをセルに設定
    Cell.Offset(0, 0).Value = rec.ID
    Cell.Offset(0, 1).Value = rec.Name
    Cell.Offset(0, 2).Value = rec.Value
End Sub

Public Function GetValue() As Record
    ' セルのデータをRecordオブジェクトに格納
    Dim rec As Record
    Set rec = New Record
    rec.ID = Cell.Offset(0, 0).Value
    rec.Name = Cell.Offset(0, 1).Value
    rec.Value = Cell.Offset(0, 2).Value
    Set GetValue = rec
End Function


Sub ExampleWithStringSplitCondition()
    ' CellManagerオブジェクトの作成
    Dim cellManager As CellManager
    Set cellManager = New CellManager

    ' 対象のセルを設定
    Set cellManager.Cell = ThisWorkbook.Sheets("Sheet2").Cells(1, 1)

    ' GetValue関数を呼び出してRecordオブジェクトを取得
    Dim rec As Record
    Set rec = cellManager.GetValue

    ' rec.IDを分割
    Dim idParts() As String
    idParts = Split(rec.ID, ",")

    ' 分割した各部分に基づいて条件分岐
    If UBound(idParts) >= 0 Then
        If idParts(0) = "Part1" Then
            MsgBox "First part is 'Part1'. Name: " & rec.Name & ", Value: " & rec.Value
        ElseIf idParts(0) = "Part2" Then
            MsgBox "First part is 'Part2'. Name: " & rec.Name & ", Value: " & rec.Value
        Else
            MsgBox "First part is " & idParts(0) & ". Name: " & rec.Name & ", Value: " & rec.Value
        End If
    End If

    ' 追加の条件分岐（必要に応じて）
    If UBound(idParts) >= 1 Then
        If idParts(1) = "SubPart1" Then
            MsgBox "Second part is 'SubPart1'. Name: " & rec.Name & ", Value: " & rec.Value
        ElseIf idParts(1) = "SubPart2" Then
            MsgBox "Second part is 'SubPart2'. Name: " & rec.Name & ", Value: " & rec.Value
        Else
            MsgBox "Second part is " & idParts(1) & ". Name: " & rec.Name & ", Value: " & rec.Value
        End If
    End If
End Sub

Sub SetupTestData()
    ' テストデータを設定するサブルーチン
    Dim cellManager As CellManager
    Set cellManager = New CellManager

    Dim rec As Record
    Set rec = New Record

    rec.ID = "Part1,SubPart1"
    rec.Name = "Test Record"
    rec.Value = 123.45

    ' 対象のセルを設定
    Set cellManager.Cell = ThisWorkbook.Sheets("Sheet2").Cells(1, 1)

    ' セルにRecordのデータを設定
    cellManager.SetValue rec
End Sub


Sub ExampleUsage()
    ' CellManagerオブジェクトの作成
    Dim cellManager As CellManager
    Set cellManager = New CellManager

    ' 対象のセルを設定（例：Sheet2のA1セル）
    Set cellManager.Cell = ThisWorkbook.Sheets("Sheet2").Cells(1, 1)

    ' Recordオブジェクトを作成し、データを設定
    Dim rec As Record
    Set rec = New Record
    rec.ID = "Part1,SubPart1"
    rec.Name = "Test Record"
    rec.Value = 123.45

    ' Recordオブジェクトのデータをセルに書き込む
    cellManager.SetValue rec

    ' セルからRecordオブジェクトのデータを読み込む
    Dim newRec As Record
    Set newRec = cellManager.GetValue

    ' 読み込んだデータを確認する
    Debug.Print newRec.ID
    Debug.Print newRec.Name
    Debug.Print newRec.Value
End Sub



 F110F-Hot-天	110F	天頂	2024/5/17	26.0 ℃	3.07 kN
 F110F-Cold-天	110F	天頂	2024/5/17	26.0 ℃	4.91 kN
 F110F-Wet-天	110F	天頂	2024/5/17	26.0 ℃	2.89 kN
 F110F-Hot-前	110F	前頭部	2024/5/17	26.0 ℃	5.25 kN
 F110F-Cold-前	110F	前頭部	2024/5/17	26.0 ℃	5.06 kN
 F110F-Wet-前	110F	前頭部	2024/5/17	26.0 ℃	5.64 kN
 F110F-Hot-後	110F	後頭部	2024/5/17	26.0 ℃	5.12 kN
 F110F-Cold-後	110F	後頭部	2024/5/17	26.0 ℃	5.77 kN
 F110F-Wet-後	110F	後頭部	2024/5/17	26.0 ℃	5.19 kN


' Record クラスモジュール
Public Class Record
    Public ID As String
    Public Temperature As String
    Public Location As String
    Public DateValue As Date
    Public TemperatureValue As String
    Public Force As Double
    Public Values As Collection ' 複数行の値を保持するコレクション

    ' 初期化メソッド（単独行）
    Public Sub InitSingle(ByVal ID As String, ByVal Temperature As String, ByVal Location As String, ByVal DateValue As Date, ByVal TemperatureValue As String, ByVal Force As Double)
        Me.ID = ID
        Me.Temperature = Temperature
        Me.Location = Location
        Me.DateValue = DateValue
        Me.TemperatureValue = TemperatureValue
        Me.Force = Force
    End Sub

    ' 初期化メソッド（複数行）
    Public Sub InitMultiple(ByVal ID As String, ByVal Values As Collection)
        Me.ID = ID
        Set Me.Values = Values
    End Sub
End Class

' DataSetManager クラスモジュール
Public Class DataSetManager
    Public records As Collection

    ' 初期化メソッド
    Public Sub Init()
        Set records = New Collection
    End Sub

    ' データをシートから読み込み、レコードインスタンスを作成
    Public Sub LoadData(ByVal ws As Worksheet)
        Dim lastRow As Long
        Dim i As Long
        Dim currentID As String
        Dim record As Record
        Dim rowData As Collection

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        i = 2
        Do While i <= lastRow
            currentID = ws.Cells(i, 1).Value

            ' 単独行のチェック
            If ws.Cells(i, 1).Value Like "*-*" Then
                Set record = New Record
                record.InitSingle ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value
                records.Add record
                i = i + 1
            Else
                ' 複数行のチェック
                Set rowData = New Collection
                Do While i <= lastRow And ws.Cells(i, 1).Value = currentID
                    rowData.Add Array(ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value)
                    i = i + 1
                Loop
                Set record = New Record
                record.InitMultiple currentID, rowData
                records.Add record
            End If
        Loop
    End Sub

    ' 選択基準に基づいて値をシートに書き込む
    Public Sub WriteSelectedValuesToSheet(sheetName As String, criteria As String)
        Dim ws As Worksheet
        Dim record As Record
        Dim row As Integer
        Dim i As Integer
        Dim valueArray As Variant

        Set ws = ThisWorkbook.Sheets(sheetName)
        row = 1 ' データを開始する行

        For Each record In records
            If record.ID Like "*" & criteria & "*" Then
                If Not record.Values Is Nothing Then
                    For i = 1 To record.Values.Count
                        valueArray = record.Values(i)
                        ws.Cells(row, 1).Value = record.ID
                        ws.Cells(row, 2).Value = valueArray(0) ' Temperature
                        ws.Cells(row, 3).Value = valueArray(1) ' Location
                        ws.Cells(row, 4).Value = valueArray(2) ' DateValue
                        ws.Cells(row, 5).Value = valueArray(3) ' TemperatureValue
                        ws.Cells(row, 6).Value = valueArray(4) ' Force
                        row = row + 1
                    Next i
                Else
                    ws.Cells(row, 1).Value = record.ID
                    ws.Cells(row, 2).Value = record.Temperature
                    ws.Cells(row, 3).Value = record.Location
                    ws.Cells(row, 4).Value = record.DateValue
                    ws.Cells(row, 5).Value = record.TemperatureValue
                    ws.Cells(row, 6).Value = record.Force
                    row = row + 1
                End If
            End If
        Next record
    End Sub
End Class

Sub Main()
    Dim ws As Worksheet
    Dim dataSetManager As DataSetManager
    Dim criteria As String

    ' データシートの設定
    Set ws = ThisWorkbook.Sheets("DataSheet")

    ' DataSetManagerの初期化とデータロード
    Set dataSetManager = New DataSetManager
    dataSetManager.Init
    dataSetManager.LoadData ws

    ' 選択基準を設定（例: "天"）
    criteria = "天"

    ' 選択基準に基づいてデータを別のシートに書き込む
    dataSetManager.WriteSelectedValuesToSheet "OutputSheet", criteria
End Sub

| A                | B    | C     | D          | E        | F      |
|------------------|------|-------|------------|----------|--------|
| ID               | Temp | Loc   | Date       | TempVal  | Force  |
| F110F-Hot-天     | 110F | 天頂  | 2024/5/17  | 26.0 ℃   | 3.07 kN|
| F110F-Cold-天    | 110F | 天頂  | 2024/5/17  | 26.0 ℃   | 4.91 kN|
| F110F-Wet-天     | 110F | 天頂  | 2024/5/17  | 26.0 ℃   | 2.89 kN|
| F110F-Hot-前     | 110F | 前頭部| 2024/5/17  | 26.0 ℃   | 5.25 kN|
| F110F-Cold-前    | 110F | 前頭部| 2024/5/17  | 26.0 ℃   | 5.06 kN|
| F110F-Wet-前     | 110F | 前頭部| 2024/5/17  | 26.0 ℃   | 5.64 kN|
| F110F-Hot-後     | 110F | 後頭部| 2024/5/17  | 26.0 ℃   | 5.12 kN|
| F110F-Cold-後    | 110F | 後頭部| 2024/5/17  | 26.0 ℃   | 5.77 kN|
| F110F-Wet-後     | 110F | 後頭部| 2024/5/17  | 26.0 ℃   | 5.19 kN|



No.340, 341 の必要試料

No.340(穴なし) 帽体白：ハンモック長さ310mm
 社内データ取り試験用：9個 (高温、低温、浸せき3個ずつ)
 型式申請用：８個(申請1件につき4個ずつ、今回は飛来・落下用と墜落時用の2件)

No.341(穴あり) 帽体白：ハンモック長さ310mm
 社内データ取り試験用：9個 (高温、低温、浸せき3個ずつ)
 型式申請用：８個(申請1件につき4個ずつ、今回は飛来・落下用と墜落時用の2件)


' Record クラスモジュール
Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As String
Public Force As Double
Public Values As Collection ' 複数行の値を保持するコレクション

' 初期化メソッド（単独行）
Public Sub InitSingle(ByVal ID As String, ByVal Temperature As String, ByVal Location As String, ByVal DateValue As Date, ByVal TemperatureValue As String, ByVal Force As Double)
    Me.ID = ID
    Me.Temperature = Temperature
    Me.Location = Location
    Me.DateValue = DateValue
    Me.TemperatureValue = TemperatureValue
    Me.Force = Force
End Sub

' 初期化メソッド（複数行）
Public Sub InitMultiple(ByVal ID As String, ByVal Values As Collection)
    Me.ID = ID
    Set Me.Values = Values
End Sub

' DataSetManager クラスモジュール
Public records As Collection

' 初期化メソッド
Public Sub Init()
    Set records = New Collection
End Sub

' データをシートから読み込み、レコードインスタンスを作成
Public Sub LoadData(ByVal ws As Worksheet)
    On Error GOTO LoadDataError
    Dim lastRow As Long
    Dim i As Long
    Dim currentID As String
    Dim record As Record
    Dim rowData As Collection

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    i = 2
    Do While i <= lastRow
        currentID = ws.Cells(i, 1).Value

        ' 単独行のチェック
        If ws.Cells(i, 1).Value Like "*-*" Then
            Set record = New Record
            record.InitSingle ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value
            records.Add record
            i = i + 1
        Else
            ' 複数行のチェック
            Set rowData = New Collection
            Do While i <= lastRow And ws.Cells(i, 1).Value = currentID
                rowData.Add Array(ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value)
                i = i + 1
            Loop
            Set record = New Record
            record.InitMultiple currentID, rowData
            records.Add record
        End If
    Loop
    Exit Sub
LoadDataError:
    MsgBox "データの読み込み中にエラーが発生しました: " & Err.Description
End Sub

' 選択基準に基づいて値をシートに書き込む
Public Sub WriteSelectedValuesToSheet(sheetName As String, criteria As String)
    On Error GoTo WriteDataError

    Dim ws As Worksheet
    Dim record As Record
    Dim row As Integer
    Dim i As Integer
    Dim valueArray As Variant

    ' シートの存在確認
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo WriteDataError
    If ws Is Nothing Then
        MsgBox "指定されたシートが存在しません: " & sheetName
        Exit Sub
    End If

    row = 2 ' データを開始する行

    For Each record In records
        If record.ID Like "*" & criteria & "*" Then
            If Not record.Values Is Nothing Then
                For i = 1 To record.Values.Count
                    valueArray = record.Values(i)
                    ws.Cells(row, 1).Value = record.ID
                    ws.Cells(row, 2).Value = valueArray(0) ' Temperature
                    ws.Cells(row, 3).Value = valueArray(1) ' Location
                    ws.Cells(row, 4).Value = valueArray(2) ' DateValue
                    ws.Cells(row, 5).Value = valueArray(3) ' TemperatureValue
                    ws.Cells(row, 6).Value = valueArray(4) ' Force
                    row = row + 1
                Next i
            Else
                ws.Cells(row, 1).Value = record.ID
                ws.Cells(row, 2).Value = record.Temperature
                ws.Cells(row, 3).Value = record.Location
                ws.Cells(row, 4).Value = record.DateValue
                ws.Cells(row, 5).Value = record.TemperatureValue
                ws.Cells(row, 6).Value = record.Force
                row = row + 1
            End If
        End If
    Next record
    Exit Sub

WriteDataError:
    MsgBox "データの書き込み中にエラーが発生しました: " & Err.Description
End Sub


Sub Main()
    Dim ws As Worksheet
    Dim dataSetManager As DataSetManager
    Dim criteria As String

    ' データシートの設定
    Set ws = ThisWorkbook.Sheets("DataSheet")

    ' DataSetManagerの初期化とデータロード
    Set dataSetManager = New DataSetManager
    dataSetManager.Init
    dataSetManager.LoadData ws

    ' 選択基準を設定（例: "天"）
    criteria = "天"

    ' 選択基準に基づいてデータを別のシートに書き込む
    dataSetManager.WriteSelectedValuesToSheet "OutputSheet", criteria
End Sub

以上はExcelVBAでクラスモジュールを使用した一連のコードです。このコードの簡単な解説をお願いします。
またエラーチェックをお願いします。実際にコードの修正をする必要はありません。


record.IDは'01-F110F-Hot-前'のような形式になっています。
これを利用して場合分けを行いたい。
#条件
- 先頭の'01'の数字は試料No.です。これが同一の場合MultiRecordとして同じレコードに入ります。
- また1個目の-のあとの文字列'F110F'のようにFが含まれている場合はMultiRecordで、それ以外はSingleRecordです。

レコードクラスは上記のようになっています。
ws.Range("B" & i).Value,_
以下の一連の処理はレコードクラスのプロパティに順番に値を格納しているという理解でよろしいでしょうか？
つまり
ws.Range("B" & i).Value = Me.ID
ws.Range("C" & i).Value = Me.Temperature...


' Record クラスモジュール
Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As String
Public Force As Double
Public Values As Collection ' 複数行の値を保持するコレクション

' 初期化メソッド（単独行）
Public Sub InitSingle(ByVal ID As String, ByVal Temperature As String, ByVal Location As String, ByVal DateValue As Date, ByVal TemperatureValue As String, ByVal Force As Double)
    Me.ID = ID
    Me.Temperature = Temperature
    Me.Location = Location
    Me.DateValue = DateValue
    Me.TemperatureValue = TemperatureValue
    Me.Force = Force
End Sub

' 初期化メソッド（複数行）
Public Sub InitMultiple(ByVal ID As String, ByVal Values As Collection)
    Me.ID = ID
    Set Me.Values = Values
End Sub

' 特定の行と列の値を取得するメソッド
Public Function GetValue(ByVal rowIndex As Integer, ByVal colIndex As Integer) As Variant
    If rowIndex > 0 And rowIndex <= Me.Values.Count Then
        GetValue = Me.Values(rowIndex)(colIndex - 1) ' colIndexは1から始まるが、Arrayは0から始まるため調整
    Else
        GetValue = CVErr(xlErrValue) ' 無効なインデックスの場合はエラーを返す
    End If
End Function

Sub Main()
    Dim ws As Worksheet
    Dim dataSetManager As DataSetManager
    Dim criteriaList As Variant
    Dim criteria As Variant
    Dim sheetName As String
    Dim specificValue As Variant
    Dim record As Record

    ' データシートの設定
    Set ws = ThisWorkbook.Sheets("DataSheet")

    ' DataSetManagerの初期化とデータロード
    Set dataSetManager = New DataSetManager
    dataSetManager.Init
    dataSetManager.LoadData ws

    ' 選択基準のリストを設定
    criteriaList = Array("SingleRecord", "MultiRecord") ' 条件に基づいた基準を設定

    ' 選択基準に基づいてデータを別のシートに書き込む
    For Each criteria In criteriaList
        If criteria = "SingleRecord" Then
            sheetName = "OutputSingle"
        Else
            sheetName = "OutputSheet"
        End If
        dataSetManager.WriteSelectedValuesToSheet sheetName, CStr(criteria)
    Next criteria

    ' 特定のレコードから特定の値を取得する例
    For Each record In dataSetManager.records
        If InStr(record.ID, "F") > 0 Then ' 複数行レコード
            specificValue = record.GetValue(2, 4) ' 2行目の4列目の値を取得
            MsgBox "The value in the 2nd row and 4th column is: " & specificValue
        End If
    Next record
End Sub


' ----------------------------------------------------------------------------------------------------------------
' 2024年5月28日のコード
' Record クラスモジュール
Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As String
Public Force As Double
Public Values As Collection ' 複数行の値を保持するコレクション

' 初期化メソッド（単独行）
Public Sub InitSingle(ByVal ID As String, ByVal Temperature As String, ByVal Location As String, ByVal DateValue As Date, ByVal TemperatureValue As String, ByVal Force As Double)
    Me.ID = ID
    Me.Temperature = Temperature
    Me.Location = Location
    Me.DateValue = DateValue
    Me.TemperatureValue = TemperatureValue
    Me.Force = Force
End Sub

' 初期化メソッド（複数行）
Public Sub InitMultiple(ByVal ID As String, ByVal Values As Collection)
    Me.ID = ID
    Set Me.Values = Values
End Sub

' 複数行レコードから特定の値を取得するメソッド
Public Function GetSpecificValues() As Collection
    Dim result As New Collection
    Dim firstRow As Variant
    Dim secondRow As Variant
    Dim thirdRow As Variant

    If Me.Values.Count = 3 Then
        ' 1つ目のレコードの全ての値
        firstRow = Me.Values(1)
        result.Add firstRow

        ' 2つ目のレコードの3, 4, 5列目の値
        secondRow = Array( _
            Me.Values(2)(2), _  ' 3列目の値
            Me.Values(2)(3), _  ' 4列目の値
            Me.Values(2)(4)     ' 5列目の値
        )
        result.Add secondRow

        ' 3つ目のレコードの3, 4, 5列目の値
        thirdRow = Array( _
            Me.Values(3)(2), _  ' 3列目の値
            Me.Values(3)(3), _  ' 4列目の値
            Me.Values(3)(4)     ' 5列目の値
        )
        result.Add thirdRow
    Else
        ' エラーハンドリング（適切なメッセージや処理を追加）
        MsgBox "Values collection does not contain exactly 3 records."
    End If

    Set GetSpecificValues = result
End Function


' DataSetManager クラスモジュール
Public records As Collection

' 初期化メソッド
Public Sub Init()
    Set records = New Collection
End Sub

' データをシートから読み込み、レコードインスタンスを作成
Public Sub LoadData(ByVal ws As Worksheet)
    On Error GoTo LoadDataError
    Dim lastRow As Long
    Dim i As Long
    Dim currentID As String
    Dim record As Record
    Dim rowData As Collection
    Dim sampleNo As String
    Dim identifier As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    i = 2
    Do While i <= lastRow
        currentID = ws.Range("A" & i).Value
        sampleNo = Split(currentID, "-")(0) ' 01の部分
        identifier = Split(currentID, "-")(1) ' F110Fの部分

        ' 'F'を含むかどうかのチェック
        If InStr(identifier, "F") > 0 Then
            ' 複数行のチェック（同じ試料No.の場合）
            Set rowData = New Collection
            Do While i <= lastRow And Split(ws.Range("A" & i).Value, "-")(0) = sampleNo
                rowData.Add Array( _
                    ws.Range("B" & i).Value, _
                    ws.Range("C" & i).Value, _
                    ws.Range("D" & i).Value, _
                    ws.Range("E" & i).Value, _
                    ws.Range("F" & i).Value)
                i = i + 1
            Loop
            Set record = New Record
            record.InitMultiple currentID, rowData
            records.Add record
        Else
            ' 単独行のチェック
            Set record = New Record
            record.InitSingle _
                ws.Range("A" & i).Value, _
                ws.Range("B" & i).Value, _
                ws.Range("C" & i).Value, _
                ws.Range("D" & i).Value, _
                ws.Range("E" & i).Value, _
                ws.Range("F" & i).Value
            records.Add record
            i = i + 1
        End If
    Loop
    Exit Sub
LoadDataError:
    MsgBox "データの読み込み中にエラーが発生しました: " & Err.Description
End Sub

' 選択基準に基づいて値をシートに書き込む
Public Sub WriteSelectedValuesToSheet(sheetName As String, criteria As String)
    On Error GoTo WriteDataError

    Dim ws As Worksheet
    Dim record As Record
    Dim row As Integer
    Dim i As Integer
    Dim valueArray As Variant
    Dim singleSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetIndex As Integer

    If criteria = "SingleRecord" Then
        ' 基準がSingleRecordの場合
        Set singleSheet = ThisWorkbook.Sheets(sheetName)
        sheetIndex = 1
        For Each record In records
            If InStr(record.ID, "F") = 0 Then ' identifierに'F'を含まないレコードを処理
                ' OutputSingleシートを複製
                singleSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                newSheet.Name = sheetName & "_" & sheetIndex
                sheetIndex = sheetIndex + 1

                ' 1レコードずつ出力
                newSheet.Range("A2").Value = record.ID
                newSheet.Range("B2").Value = record.Temperature
                newSheet.Range("C2").Value = record.Location
                newSheet.Range("D2").Value = record.DateValue
                newSheet.Range("E2").Value = record.TemperatureValue
                newSheet.Range("F2").Value = record.Force
            End If
        Next record
    Else
        ' 基準がMultiRecordの場合
        Set ws = ThisWorkbook.Sheets(sheetName)
        row = 2 ' データを開始する行

        For Each record In records
            If InStr(record.ID, "F") > 0 Then ' identifierに'F'を含むレコードを処理
                Set specificValues = record.GetSpecificValues
                For Each value In specificValues
                    ws.Range("A" & row).Value = record.ID
                    ws.Range("B" & row).Value = value(0) ' Temperature
                    ws.Range("C" & row).Value = value(1) ' Location
                    ws.Range("D" & row).Value = value(2) ' DateValue
                    ws.Range("E" & row).Value = value(3) ' TemperatureValue
                    ws.Range("F" & row).Value = value(4) ' Force
                    row = row + 1
                Next value
            End If
        Next record
    End If

    Exit Sub

WriteDataError:
    MsgBox "データの書き込み中にエラーが発生しました: " & Err.Description
End Sub

Sub Main()
    Dim ws As Worksheet
    Dim dataSetManager As DataSetManager
    Dim criteriaList As Variant
    Dim criteria As Variant
    Dim sheetName As String
    Dim specificValues As Collection
    Dim record As Record
    Dim value As Variant

    ' データシートの設定
    Set ws = ThisWorkbook.Sheets("DataSheet")

    ' DataSetManagerの初期化とデータロード
    Set dataSetManager = New DataSetManager
    dataSetManager.Init
    dataSetManager.LoadData ws

    ' 選択基準のリストを設定
    criteriaList = Array("SingleRecord", "MultiRecord") ' 条件に基づいた基準を設定

    ' 選択基準に基づいてデータを別のシートに書き込む
    For Each criteria In criteriaList
        If criteria = "SingleRecord" Then
            sheetName = "OutputSingle"
        Else
            sheetName = "OutputSheet"
        End If
        dataSetManager.WriteSelectedValuesToSheet sheetName, CStr(criteria)
    Next criteria

    ' 特定のレコードから特定の値を取得する例
    For Each record In dataSetManager.records
        If InStr(record.ID, "F") > 0 Then ' 複数行レコード
            Set specificValues = record.GetSpecificValues
            For Each value In specificValues
                Debug.Print Join(value, ", ")
            Next value
        End If
    Next record
End Sub
Public Sub WriteSelectedValuesToSheet(sheetName As String, criteria As String)
    On Error GoTo WriteDataError

    Dim ws As Worksheet
    Dim record As Record
    Dim row As Integer
    Dim i As Integer
    Dim valueArray As Variant
    Dim singleSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetIndex As Integer

    If criteria = "SingleRecord" Then
        ' 基準がSingleRecordの場合
        Set singleSheet = ThisWorkbook.Sheets(sheetName)
        sheetIndex = 1
        For Each record In records
            If InStr(record.ID, "F") = 0 Then ' identifierに'F'を含まないレコードを処理
                ' OutputSingleシートを複製
                singleSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                newSheet.Name = sheetName & "_" & sheetIndex
                sheetIndex = sheetIndex + 1

                ' 1レコードずつ出力
                newSheet.Range("A2").Value = record.ID
                newSheet.Range("B2").Value = record.Temperature
                newSheet.Range("C2").Value = record.Location
                newSheet.Range("D2").Value = record.DateValue
                newSheet.Range("E2").Value = record.TemperatureValue
                newSheet.Range("F2").Value = record.Force
            End If
        Next record
    Else
        ' 基準がMultiRecordの場合
        Set ws = ThisWorkbook.Sheets(sheetName)
        row = 2 ' データを開始する行

        For Each record In records
            If InStr(record.ID, "F") > 0 Then ' identifierに'F'を含むレコードを処理
                Set specificValues = record.GetSpecificValues
                For Each value In specificValues
                    ws.Range("A" & row).Value = record.ID
                    ws.Range("B" & row).Value = value(0) ' Temperature
                    ws.Range("C" & row).Value = value(1) ' Location
                    ws.Range("D" & row).Value = value(2) ' DateValue
                    ws.Range("E" & row).Value = value(3) ' TemperatureValue
                    ws.Range("F" & row).Value = value(4) ' Force
                    row = row + 1
                Next value
            End If
        Next record
    End If

    Exit Sub

WriteDataError:
    MsgBox "データの書き込み中にエラーが発生しました: " & Err.Description
End Sub



Processing row: 1, ID: 01-F110F-Hot-天, SampleNo: 01
Adding to rowData: 1, SampleNo: 01
Adding to rowData: 4, SampleNo: 01
Adding to rowData: 6, SampleNo: 01
Initialized Multiple Record with ID: 01-F110F-Hot-天, Values count: 3
Initialized Multiple Record with ID: 01-F110F-Hot-天, Values count: 3
Recorded ID: 01-F110F-Hot-天, SampleNo: 01
    110F, 天頂, 2024/05/17, 26, 3.07
    110F, 前頭部, 2024/05/17, 26, 5.25
    110F, 後頭部, 2024/05/17, 26, 5.12
Processing row: 2, ID: 02-110-Cold-天, SampleNo: 02
Initializing Single Record: 2
Recorded ID: 02-110-Cold-天, SampleNo: 02
    110, 天頂, 2024/05/17, 26, 4.91
Processing row: 3, ID: 03-F110F-Wet-天, SampleNo: 03
Adding to rowData: 3, SampleNo: 03
Adding to rowData: 5, SampleNo: 03
Adding to rowData: 7, SampleNo: 03
Initialized Multiple Record with ID: 03-F110F-Wet-天, Values count: 3
Initialized Multiple Record with ID: 03-F110F-Wet-天, Values count: 3
Recorded ID: 03-F110F-Wet-天, SampleNo: 03
    110F, 天頂, 2024/05/17, 26, 2.89
    110F, 前頭部, 2024/05/17, 26, 5.64
    110F, 後頭部, 2024/05/17, 26, 5.19
Values collection count: 3
Values collection count: 3
110F, 天頂, 2024/05/17, 26, 3.07
2024/05/17, 26, 5.25
2024/05/17, 26, 5.12
Values collection count: 3
110F, 天頂, 2024/05/17, 26, 2.89
2024/05/17, 26, 5.64
2024/05/17, 26, 5.19

ID	Temperature	Location	DateValue	TemperatureVallue	Force
01-F110F-Hot-天	110F	天頂	2024/5/17	29	3.07
02-110-Cold-天	110	天頂	2024/5/17	26	4.91
03-F110F-Wet-天	110F	天頂	2024/5/17	26	2.89
01-F110F-Hot-前	110F	前頭部	2024/5/17	26	5.25
03-F110F-Wet-前	110F	前頭部	2024/5/17	29	5.64
01-F110F-Hot-後	110F	後頭部	2024/5/17	26	5.12
03-F110F-Wet-後	110F	後頭部	2024/5/17	29	5.19

理解しました。私のもう一つの目的は'OutputSingle/OutputSheet'のペアと似たような構造のシートのコピーを作ることです。
つまり先ほど示したシートの内容を以下のシートに振り分けることが目的です。
- 'OutputSingle/OutputSheet'のペア
- 'Rstl_Single/Rstl_Triple'のペア
- 'Resutl_Temp'のペア
これらのシート名が被らないような方法を考えたいです。
- 'OutputSingle/OutputSheet'のペア
- 'Rstl_Single/Rstl_Triple'のペア
のシート名作成方法は先に示した方法と同じで構いません。もう一度総合的なロジックをステップバイステップで考えてください。
何か足りない情報がありましたら聞いてください。


OutputSingle/OutputSheet シートの作成とデータの書き込みとRstl_Single/Rstl_Triple シートの作成とデータの書き込みはシート名が異なるだけでロジックは同じなはずです。まとめることはできないのでしょうか？
まとめた場合にどのようなロジックになるかを言葉で説明してください。コードの提示は私がそれに納得してからお願いします。

Sub CopyAndPopulateSheet( _
    sourceSheetName As String, _
    prefix As String, _
    index As Integer, _
    customPropertyName As String, _
    customPropertyValue As String, _
    dataCollection As Collection, _
    writeMethod As String)

    Dim sheetName As String
    Dim ws As Worksheet

    sheetName = GenerateSheetName(prefix, index)
    Debug.Print "Generated sheet name: " & sheetName

    ' シートの存在確認と作成
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        ' コピーするソースシートが存在するか確認
        If Not SheetExists(sourceSheetName) Then
            Debug.Print "Source sheet not found: " & sourceSheetName
            Exit Sub
        End If

        Sheets(sourceSheetName).Copy After:=Sheets(Sheets.Count)
        Set ws = ActiveSheet
        ws.Name = sheetName

        ' シート名の変更が成功したか確認
        If ws.Name <> sheetName Then
            Debug.Print "Failed to rename the sheet correctly."
            Exit Sub
        End If
    End If
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' カスタムプロパティの設定
        ws.CustomProperties.Add Name:=customPropertyName, Value:=customPropertyValue

        ' データの転記
        Debug.Print "Using write method: " & writeMethod
        Select Case writeMethod
            Case "WriteSelectedValuesToOutputSheet"
                DataSetManager.WriteSelectedValuesToOutputSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToRstlSheet"
                DataSetManager.WriteSelectedValuesToRstlSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToResultTempSheet"
                DataSetManager.WriteSelectedValuesToResultTempSheet ws.Name, dataCollection
            Case Else
                Debug.Print "Unknown write method: " & writeMethod
        End Select
        Debug.Print "Data written successfully to " & ws.Name
    Else
        Debug.Print "Failed to create or find the sheet: " & sheetName
    End If

End Sub



2024年6月3日やること
TestSheetCreationAndDataWriting()で、レコードを作成してテストをしているが、これをRecordクラスモジュールから読み取って行う。
DataSetManagerクラスモジュールで書き込みの場合分けを実装する。IDから飛来・落下試験と墜落試験を分ける。


Public Sub LoadData(ByVal ws As Worksheet, ByVal row As Integer)
    ID = ws.Cells(row, 1).Value
    Temperature = ws.Cells(row, 2).Value
    Location = ws.Cells(row, 3).Value
    DateValue = ws.Cells(row, 4).Value
    TemperatureValue = ws.Cells(row, 5).Value
    Force = ws.Cells(row, 6).Value

    ' IDを分析してカテゴリを決定
    Dim parts() As String
    parts = Split(ID, "-")

    ' 中間部分でのカテゴリ判定
    If InStr(parts(1), "F") > 0 Then
        Category = "MultiValue"
    Else
        Category = "SingleValue"
    End If

    ' 末尾部分でのグループ判定
    Select Case parts(3)
        Case "側"
            Group = "SideValue." & parts(3) & "." & parts(2)
        Case Else
            Group = Category & "." & parts(3) & "." & parts(2)
    End Select

    ' TemperatureValue の値に基づいたグループコードの追加
    If TemperatureValue = 26 Then
        Group = Group & "." & "Spot"
    Else
        Group = Group & "." & "Regular"
    End If
End Sub
```

全体のコードの整理をしたいと思います。
RecordクラスモジュールのRecordプロシージャでRecordの原型を作成します。
RecordクラスモジュールのLoadDataプロシージャでRecordのインスタンスを作成します。
また、LoadDataプロシージャでレコードごとにキーを作成します。
PopulateGroupedSheetsを用いてそのキーと各シートを紐付けます。
DataSetManagerクラスモジュールにある各WriteSelectedValuesTo〜モジュールに合わせてシートに値を書き込みます。

```vb
Sub PopulateGroupedSheets(groupedRecords As Object)
    Dim key As Variant
    Dim recordGroup As Collection
    Dim templateSheetName As String
    Dim newSheetName As String
    Dim targetSheet As Worksheet

    For Each key In groupedRecords.Keys
        Set recordGroup = groupedRecords(key)

        ' Determine the appropriate template sheet based on the group key
        Select Case key
            Case "申請_飛来"
                templateSheetName = "申請_飛来"
            Case "申請_墜落"
                templateSheetName = "申請_墜落"
            Case "定期_飛来"
                templateSheetName = "定期_飛来"
            Case "定期_墜落"
                templateSheetName = "定期_墜落"
            Case "側面"
                templateSheetName = "側面"
            Case "依頼試験"
                templateSheetName = "依頼試験"
            Case Else
                Debug.Print "No matching template found for: " & key
                GoTo NextGroup
        End Select

        ' Check if the sheet already exists and copy it if necessary
        If Not SheetExists(templateSheetName) Then
            Debug.Print "Template sheet does not exist: " & templateSheetName
            GoTo NextGroup
        Else
            Sheets(templateSheetName).Copy After:=Sheets(Sheets.Count)
            Set targetSheet = ActiveSheet
            targetSheet.Name = templateSheetName & "_" & Format(Now, "yyyymmdd_hhmmss")
        End If

        ' Populate the new sheet with data

key:MultiValue.後.Wetがキーの内容。この内容でシートの振り分けができるかを考える。
        Dim newName as string
        ' Template sheet determination based on group key
        If InStr(key, "SingleValue") > 0 Then
            templateName = "申請_飛来"
            templateName = "定期_飛来"
        ElseIf InStr(key, "SideValue") > 0 Then
            templateName = "側面試験"
        Else
            templateName = "申請_墜落"
            templateName = "定期_墜落"
        End If

このように
各値にそって1枚~2枚のシートを複製し、それをwrite~メソッドに従って値を振り分けたい。

EXCELVBAで条件に従い、行全体を非表示にするコードを書いてください。
# 条件
- 範囲はA列の値があるところまで
- C列の値がない行を非表示にする。
- 非表示にしている行を表示する。

newSheetName = key & "_" & templateName & "_" & sheetIndex
keyの部分をtemplateNames.ADD = "申請_飛来" の前の2文字にできますか?
提示している例なら"申請"のようになります。

このコードを実行したときにキーの値はデバッグ出力するとこうなります。
key:SingleValue.110.天.Cold
key:SingleValue.110.天.Cold
key:MultiValue.110F.天.Wet
key:MultiValue.110F.天.Wet
key:MultiValue.110F.前.Hot
key:MultiValue.110F.前.Hot
key:MultiValue.110F.前.Wet
key:MultiValue.110F.前.Wet
key:MultiValue.110F.後.Hot
key:MultiValue.110F.後.Hot
key:MultiValue.110F.後.Wet
key:MultiValue.110F.後.Wet
key:SingleValue.110.天.Cold
key:SingleValue.110.天.Cold
key:MultiValue.170F.天.Wet
key:MultiValue.170F.天.Wet
key:MultiValue.170F.前.Hot
key:MultiValue.170F.前.Hot
key:MultiValue.170F.前.Wet
key:MultiValue.170F.前.Wet
key:MultiValue.170F.後.Hot
key:MultiValue.170F.後.Hot
key:MultiValue.170F.後.Wet
key:MultiValue.170F.後.Wet
このキーの値にそってシートを複製したいです。同じ項目をコピーしている部分もあるのでそれにも気を配ってください。

まずはグループの例を示すので振り分けのロジックを考えてください。
Group1;key:SingleValue.110.天.Cold ;申請_飛来
Group2;MultiValue.110F.天.Wet, MultiValue.110F.前.Wet, MultiValue.110F.後.Wet  ;申請_墜落シート１枚目
Group3;MultiValue.110F.天.Hot, MultiValue.110F.前.Hot, MultiValue.110F.後.Hot  ;申請_飛来シート2枚目


Group 1
key: SingleValue.110.天.Cold
key: SingleValue.110.天.Cold
key: SingleValue.110.天.Cold
key: SingleValue.110.天.Cold
Group 2
key: MultiValue.110F.前.Hot
key: MultiValue.110F.後.Hot
Group 3
key: MultiValue.110F.前.Hot
key: MultiValue.110F.後.Hot
Group 4
key: MultiValue.110F.天.Wet
key: MultiValue.110F.前.Wet
key: MultiValue.110F.後.Wet
Group 5
key: MultiValue.110F.天.Wet
key: MultiValue.110F.前.Wet
key: MultiValue.110F.後.Wet
Group 6
key: MultiValue.1７0F.前.Hot
key: MultiValue.1７0F.後.Hot
Group 7
key: MultiValue.1７0F.天.Wet
key: MultiValue.1７0F.前.Wet
key: MultiValue.1７0F.後.Wet
Group 8
key: MultiValue.1７0F.前.Hot
key: MultiValue.1７0F.後.Hot
Group 9
key: MultiValue.1７0F.天.Wet
key: MultiValue.1７0F.前.Wet
key: MultiValue.1７0F.後.Wet



key:SingleValue.110.天.Cold
key:SingleValue.110.天.Cold
key:MultiValue.110F.天.Wet
key:MultiValue.110F.天.Wet
key:MultiValue.110F.前.Hot
上記のキーを持ったレコードを適切に分類するコードをVBAで作成してください。
コードは段階的に条件を絞る形式で作成してください。
キーは以下のようにピリオド（.）で区切られた複数のセグメントから構成されます
<タイプ>.<番号>.<位置>.<条件>
1. まずタイプ{Single, Multi}ごとにグループを作成します。
2. 次に番号{3~4桁の数字}が同じグループを作成します。
3. 次に条件{Hot, Cold ,Wet}が共通するグループを作成します。
4. 最後に位置{天、前、後、側}のうち{天、前、後}は同グループでは重複しません。また、{側}が入っているグループは{天、前、後}のどれも含みません。


SingleValue Groups:
Group 110-Cold:
  天: SingleValue.110.天.Cold
  天: SingleValue.110.天.Cold
MultiValue Groups:
Group 110F-Wet:
  天: MultiValue.110F.天.Wet
  天: MultiValue.110F.天.Wet
Group 110F-Hot:
  前: MultiValue.110F.前.Hot


keys = Array( _
    "SingleValue.110.天.Cold", _
    "SingleValue.110.天.Cold", _
    "MultiValue.110F.天.Wet", _
    "MultiValue.110F.天.Wet", _
    "MultiValue.110F.前.Hot", _
    "SingleValue.210.前.Cold", _
    "SingleValue.210.後.Cold", _
    "MultiValue.215.前.Hot", _
    "MultiValue.215.天.Hot", _
    "SingleValue.310.側.Wet", _
    "SingleValue.320.天.Cold", _
    "SingleValue.320.前.Hot", _
    "MultiValue.320F.天.Cold", _
    "MultiValue.320F.前.Wet", _
    "MultiValue.325F.後.Hot", _
    "MultiValue.325F.側.Hot", _
    "SingleValue.330F.後.Cold", _
    "SingleValue.330F.側.Wet", _
    "MultiValue.340F.天.Wet", _
    "MultiValue.340F.前.Wet" _
)


SingleValue Groups:
Group 110-Cold:
  天:
    SingleValue.110.天.Cold
    SingleValue.110.天.Cold
Group 210-Cold:
  前:
    SingleValue.210.前.Cold
  後:
    SingleValue.210.後.Cold
Group 310-Wet-側:
  側:
    SingleValue.310.側.Wet
Group 320-Cold:
  天:
    SingleValue.320.天.Cold
Group 320-Hot:
  前:
    SingleValue.320.前.Hot
Group 330F-Cold:
  後:
    SingleValue.330F.後.Cold
Group 330F-Wet-側:
  側:
    SingleValue.330F.側.Wet
MultiValue Groups:
Group 110F-Wet:
  天:
    MultiValue.110F.天.Wet
    MultiValue.110F.天.Wet
Group 110F-Hot:
  前:
    MultiValue.110F.前.Hot
Group 215-Hot:
  前:
    MultiValue.215.前.Hot
  天:
    MultiValue.215.天.Hot
Group 320F-Cold:
  天:
    MultiValue.320F.天.Cold
Group 320F-Wet:
  前:
    MultiValue.320F.前.Wet
Group 325F-Hot:
  後:
    MultiValue.325F.後.Hot
Group 325F-Hot-側:
  側:
    MultiValue.325F.側.Hot
Group 340F-Wet:
  天:
    MultiValue.340F.天.Wet
  前:
    MultiValue.340F.前.Wet

ID
試料ID
品番
試験内容
検査日
温度
最大値
最大値を記録した時間
49kNの継続時間
73kNの継続時間
前処理
重量
天頂隙間
色
ロットNumber
帽体ロット
内装ロット
構造
貫通
試験区分

ID
sampleID
itemNum
testPart
Date
testTemp
maxValue
timeOfMax
duration49kN
duration73kN
preProcess
sampleWeight
sampleTop
sampleColor
sampleLotNum
sampleHelLot
sampleBandLot
structureResult
assemblyResult
penettationResult
testSection

Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As Double
Public Force As Double
Public Values As Collection
Public Category As String
Public Group As String
このような形で以下の変数名を宣言するコードの断片を書いてください。
変数の方はお任せします。
Public ID As String 'ID
Public sampleID As String '試料ID
Public itemNum As String '品番
Public testPart As String '試験位置
Public testDate As Date '検査日
Public testTemp As Double '温度
Public maxValue As Double '最大値
Public timeOfMax As Double '最大値を記録した時間
Public duration49kN As Double '4.9kNの継続時間
Public duration73kN As Double '7.3kNの継続時間
Public preProcess As String '前処理
Public sampleWeight As Double '重量
Public sampleTop As Double '天頂隙間
Public sampleColor As String '帽体色
Public sampleLotNum As String '製品ロット
Public sampleHelLot As String '帽体ロット
Public sampleBandLot As String '内装ロット
Public structureResult As String '構造検査
Public penetrationResult As String '貫通検査
Public testSection As String '試験区分


Added new entry to sheetTypeIndex: Single_01 = 申請_飛来_01
newSheetName: 申請_飛来_01
Created new sheet: 申請_飛来_01
Record added to sheet: 申請_飛来_01 for groupID: 01
Added new entry to sheetTypeIndex: Single_02 = 申請_飛来_02
newSheetName: 申請_飛来_02
Created new sheet: 申請_飛来_02
Record added to sheet: 申請_飛来_02 for groupID: 02
Added new entry to sheetTypeIndex: Multi_03 = 申請_墜落_03
newSheetName: 申請_墜落_03
Created new sheet: 申請_墜落_03
Record added to sheet: 申請_墜落_03 for groupID: 03
Added new entry to sheetTypeIndex: Multi_04 = 申請_墜落_04
newSheetName: 申請_墜落_04
Created new sheet: 申請_墜落_04
Record added to sheet: 申請_墜落_04 for groupID: 04
Added new entry to sheetTypeIndex: Multi_01 = 申請_墜落_01
newSheetName: 申請_墜落_01
Created new sheet: 申請_墜落_01
Record added to sheet: 申請_墜落_01 for groupID: 01
Added new entry to sheetTypeIndex: Multi_02 = 申請_墜落_02
newSheetName: 申請_墜落_02
Created new sheet: 申請_墜落_02
Record added to sheet: 申請_墜落_02 for groupID: 02
newSheetName: 申請_墜落_01
Record added to sheet: 申請_墜落_01 for groupID: 01
newSheetName: 申請_墜落_02
Record added to sheet: 申請_墜落_02 for groupID: 02
newSheetName: 申請_墜落_01
Record added to sheet: 申請_墜落_01 for groupID: 01
newSheetName: 申請_墜落_02
Record added to sheet: 申請_墜落_02 for groupID: 02
newSheetName: 申請_墜落_03
Record added to sheet: 申請_墜落_03 for groupID: 03
Total unique records: 11
既存のシートは{LOG_Helmet, DataSheet, "申請_飛来","申請_墜落","定期_飛来","定期_墜落","側面試験","依頼試験",}
です。

次のステップです。
    Dim baseTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "申請_飛来"
        Case "Multi"
            baseTemplateName = "申請_墜落"
        Case Else
            baseTemplateName = "その他"
    End Select
でsheetTypeの値を利用してbaseTemplateNameを決めていますが、これにいくつかの条件を追加したい。
- Caseが"Single"の場合 baseTemplateName = "定期_飛来"を加える。
- Caseが"Multi"の場合 baseTemplateName = "定期_墜落"を加える。
sheetType"Single"が1個見つかった場合、"申請_飛来"のコピーシートが1枚、
"定期_飛来"シートが1枚、計2枚増えることになります。どのようなロジックになるかを日本語で説明してください。


Sub ClassifyKeys(sheetType As String, groupID As String)
    ' レコードごとにシートネームを作成する
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")

    ' グループIDを作成
    groupID = Left(groupID, 2)

    Dim baseTemplateName As String
    Dim additionalTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "申請_飛来"
            additionalTemplateName = "定期_飛来"
        Case "Multi"
            baseTemplateName = "申請_墜落"
            additionalTemplateName = "定期_墜落"
        Case Else
            baseTemplateName = "その他"
            additionalTemplateName = ""
    End Select

    ' 基本テンプレートと追加テンプレートのシート処理
    Call ProcessTemplateSheet(baseTemplateName, sheetType, groupID, sheetTypeIndex)
    If additionalTemplateName <> "" Then
        Call ProcessTemplateSheet(additionalTemplateName, sheetType, groupID, sheetTypeIndex)
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object)
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID

    ' シート名の決定
    If Not sheetTypeIndex.Exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = templateName & "_" & groupID
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If

    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName

    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        Select Case templateName
            Case "申請_飛来", "申請_墜落", "定期_飛来", "定期_墜落", "側面試験", "依頼試験", "LOG_Helmet", "DataSheet"
                Worksheets(templateName).Copy After:=Worksheets(Worksheets.Count)
                Set newSheet = Worksheets(Worksheets.Count)
                newSheet.Name = newSheetName
                ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).Name = "Temp_" & newSheetName
                Debug.Print "Copied sheet from template: " & templateName & " to new sheet: " & newSheet.Name
            Case Else
                Debug.Print "No template found for templateName: " & templateName
        End Select
    Else
        Set newSheet = Worksheets(newSheetName)
    End If

    Debug.Print "Record added to sheet: " & newSheet.Name & " for groupID: " & groupID
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

' 修正させたコード
Sub ClassifyKeys(sheetType As String, groupID As String)
    ' レコードごとにシートネームを作成する
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")

    ' グループIDを作成
    groupID = Left(groupID, 2)

    Dim baseTemplateName As String
    Dim additionalTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "申請_飛来"
            additionalTemplateName = "定期_飛来"
        Case "Multi"
            baseTemplateName = "申請_墜落"
            additionalTemplateName = "定期_墜落"
        Case Else
            baseTemplateName = "その他"
            additionalTemplateName = ""
    End Select

    ' 基本テンプレートと追加テンプレートのシート処理
    ProcessTemplateSheet baseTemplateName, groupID, sheetTypeIndex
    If additionalTemplateName <> "" Then
        ProcessTemplateSheet additionalTemplateName, groupID, sheetTypeIndex
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, groupID As String, ByRef sheetTypeIndex As Object)
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID

    ' シート名の決定
    If Not sheetTypeIndex.Exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = templateName & "_" & groupID
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If

    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName

    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        If templateName <> "" Then
            Worksheets(templateName).Copy After:=Worksheets(Worksheets.Count)
            Set newSheet = Worksheets(Worksheets.Count)
            newSheet.Name = newSheetName
            ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).Name = "Temp_" & newSheetName
            Debug.Print "Copied sheet from template: " & templateName & " to new sheet: " & newSheet.Name
        Else
            Debug.Print "No template found for templateName: " & templateName
        End If
    Else
        Set newSheet = Worksheets(newSheetName)
    End If

    Debug.Print "Record added to sheet: " & newSheet.Name & " for groupID: " & groupID
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

Sub TransferDataBasedOnID()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim maxValue As Double, duration49kN As Double, duration73kN As Double
    Dim data As Collection
    Dim dataItem As Variant
    Dim nextRow As Long

    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割
        idParts = Split(wsSource.Cells(i, 3).Value, "-")
        If UBound(idParts) >= 2 Then
            ' グループ名（部位）を取得
            group = idParts(2)

            ' グループ名に基づいてシート名を設定
            Select Case group
                Case "天"
                    targetSheetName = "Impact_Top"
                Case "前"
                    targetSheetName = "Impact_Front"
                Case "後"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' 対応するグループがない場合はスキップ
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).Value
                    GoTo NextIteration
            End Select

            groupName = "Group:" & idParts(0) & group

            ' 列をアルファベットで指定して値を取得
            maxValue = wsSource.Range("H" & i).Value
            duration49kN = wsSource.Range("J" & i).Value
            duration73kN = wsSource.Range("K" & i).Value

            ' グループ名とシート名の対応を確認
            Debug.Print "Group: " & groupName & "; Sheet: " & targetSheetName
            Debug.Print "Max Value: " & Format(maxValue, "0.00") & " 49kN Duration: " & Format(duration49kN, "0.00") & " 73kN Duration: " & Format(duration73kN, "0.00")

            ' データをコレクションに追加
            data.Add Array(groupName, targetSheetName, Format(maxValue, "0.00"), Format(duration49kN, "0.00"), Format(duration73kN, "0.00"))
        End If
NextIteration:
    Next i

    ' コレクションから各シートにデータを転記
    For Each dataItem In data
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)

        ' 目的のシートを設定（存在しない場合は作成）
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsDest.Name = targetSheetName
        End If
        On Error GoTo 0

        ' 次にデータを入力する行を取得
        nextRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1

        ' データを転記
        wsDest.Cells(nextRow, 1).Value = groupName
        wsDest.Cells(nextRow, 2).Value = maxValue
        wsDest.Cells(nextRow, 3).Value = duration49kN
        wsDest.Cells(nextRow, 4).Value = duration73kN
    Next dataItem

    ' リソースを解放
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub


groupedRecords is initialilzed.
Added new group for sheetType: Single
Current count for Single:1
Record added to sheet:申請_飛来_01for groupID:01
Record added to sheet:定期_飛来_01for groupID:01
groupedRecords is initialilzed.
Current count for Single:4
Record added to sheet:申請_飛来_02for groupID:02
Record added to sheet:定期_飛来_02for groupID:02
groupedRecords is initialilzed.
Added new group for sheetType: Multi
Current count for Multi:1
Record added to sheet:申請_墜落_03for groupID:03
Record added to sheet:定期_墜落_03for groupID:03
groupedRecords is initialilzed.
Current count for Multi:4
Record added to sheet:申請_墜落_04for groupID:04
Record added to sheet:定期_墜落_04for groupID:04
groupedRecords is initialilzed.
Current count for Multi:7
Record added to sheet:申請_墜落_01for groupID:01
Record added to sheet:定期_墜落_01for groupID:01
groupedRecords is initialilzed.
Current count for Multi:10
Record added to sheet:申請_墜落_02for groupID:02
Record added to sheet:定期_墜落_02for groupID:02
groupedRecords is initialilzed.
Current count for Multi:13
Record added to sheet:申請_墜落_01for groupID:01
Record added to sheet:定期_墜落_01for groupID:01
groupedRecords is initialilzed.
Current count for Multi:16
Record added to sheet:申請_墜落_02for groupID:02
Record added to sheet:定期_墜落_02for groupID:02
groupedRecords is initialilzed.
Current count for Multi:19
Record added to sheet:申請_墜落_01for groupID:01
Record added to sheet:定期_墜落_01for groupID:01
groupedRecords is initialilzed.
Current count for Multi:22
Record added to sheet:申請_墜落_02for groupID:02
Record added to sheet:定期_墜落_02for groupID:02
groupedRecords is initialilzed.
Current count for Multi:25
Record added to sheet:申請_墜落_03for groupID:03
Record added to sheet:定期_墜落_03for groupID:03
key: Single, count: 6
key: Multi, count: 27
key: Single, count: 6
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
key: Multi, count: 27
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
  Record ID: 03-350F-後-Wet-White, Group ID: 03.350F.Wet.後, Sheet Type: Multi
Total unique records: 11


{LOG_Helmet, DataSheet, 申請_飛来, 申請_墜落, 定期_飛来, 定期_墜落, 側面試験, 依頼試験}

396-天Hot	01-396-天-Hot-White	396	天頂	2024/2/21	23 	4
396-天Cold	02-396-天-Cold-White	396	天頂	2024/2/21	23 	4
F396F-天Hot	03-396F-天-Hot-OthClr	396F	天頂	2024/2/21	23 	4
F396F-天Cold	04-396F-天-Cold-OthClr	396F	天頂	2024/2/21	23 	4
397-天Hot	01-350F-天-Hot-White	350F	天頂	2024/4/10	23 	3
397-天Cold	02-350F-天-Cold-White	350F	天頂	2024/4/10	23 	5
F397F-天Hot	01-350F-前-Hot-White	350F	前頭部	2024/4/10	23 	5
F397F-天Cold	02-350F-前-Cold-White	350F	前頭部	2024/4/10	23 	5
397-天Hot	01-350F-後-Hot-White	350F	後頭部	2024/4/10	23 	7
397-天Cold	02-350F-後-Cold-White	350F	後頭部	2024/4/10	23 	7
F397F-天Hot	03-350F-後-Wet-White	350F	後頭部	2024/4/10	23 	7
レコードは上記のテーブルの内容を読み込んでいます。問題は


お世話になっております。https://www.sg-mark.org/mailmagazine/no76/ のページを拝見して質問があります。

- CEマークは自己適合宣言ですので、 とありますが、これは業者が独自に試験を行い適合したかどうかはその独自試験の結果から判断している。 ということでしょうか？

- CE認証(認証番号：BS EN1078:2012+A1:2012(Exclude clause 4.1/6/7) という表記の製品が販売されています。これはEN1078に完全に適合したという意味でしょうか？それとも4.1条、6条、7条以外の項目に適合したという意味でしょうか？

突然の質問すみません。お手数をおかけしますがご回答よろしく。
