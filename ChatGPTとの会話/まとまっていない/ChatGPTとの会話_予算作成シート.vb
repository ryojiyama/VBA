以下のExcelの関数を条件に従うように修正してください。
# Excelの関数
=IF(E2="-","",EDATE('社外校正(Edit) (2)'!$F2,'社外校正(Edit) (2)'!$E2))

# 条件
- F2に値がなければ"-"を表示するようにしたい。
- E2に値がなければ"-"を表示するようにしたい。

=IF(OR(ISBLANK(G2), ISBLANK('社外校正(Edit) (2)'!$H2)), "-", TEXT(DATEVALUE(TEXT(EDATE('社外校正(Edit) (2)'!$H2, G2), "yyyy-mm-dd")) + 1, "yyyy-mm-dd"))

以下の条件を満たすExcelの関数を作成してください。
# 条件
→が含まれるセルのみを処理する。
→の右側の文字列をすべてO列に転記する。
→は削除する。

以下の条件を満たすExcelのマクロを作成してください。
# 条件
-「Sheet1」に表の内容を転記する。
-Q列の文字列に「社外」もしくは「どちらも」の記述がある行が対象。
-「Sheet1」のB列に転記元のO列を転記する。
-「Sheet1」のF列に転記元のB列を転記する。
- E列の数値は12で割り、その商を「1/"商"年」と「Sheet1」のE列に記載する。
- 「Sheet1」のC、D,G,列にはそれぞれ定数として「校正」、「定期」、「1」と記載する。
- 転記元の行数を数えてその行数の数だけ処理を上から行う。
- 転記先と転記元のシート名をあとから変更しやすくしてください。
- 日本語のコメントを付けてください。

以下の条件を満たすExcelのマクロを作成してください。
# 条件
- シート「Sheet1」から「まとめ_山中」へと値を転機する。
- 転記元のC列の行数を数えてその行数の数だけ処理を上から行う。
- 転記先のC列を探索し、値が「審査費用」の次の行から「部材」の前の行まで必要な行を挿入する。
- 転記先と転記元のシート名をあとから変更しやすくしてください。
- 日本語のコメントを付けてください。

おおむねうまくいきました。しかし改善点があります。以下のようにコードを修正してください。
# 条件
- 転記するデータはB列からG列です。
- 転記先のどの部分から行が挿入されたのかを確認したいので挿入した行に淡いグレーの色をつけてください。

修正点があります。
- 挿入される行が必要な行より多くなっています。プロセスを見直してください。
- 薄いグレーに塗る行が挿入された後の下の行からになっています。


以下のコードを参考に条件を満たすExcelのマクロを作成してください。
# 条件
- M列で2024年7月1日から2025年6月30日の範囲を含む列を抜き出す。
- 日付は後で変えやすくしてください。
- 日本語のコメントを付けてください。
# コード

以下の条件を満たすExcelのマクロを作成してください。
# 条件
- シート「まとめ_山中」の2行目から最終行までのA列からI列までの範囲のフォントを「游ゴシック」にしてください。。
- シートの2行目から最終行までのA列からI列までの範囲のセル色を偶数行は白、奇数行はRGB(242,242,242)で塗ってください。
- シートに存在する罫線をすべて消去してください。
- 消去した処理のあとにシート「まとめ_山中」の2行目から最終行までのA列からI列までの範囲に罫線「xlHairline」を引いてください。
- 日本語のコメントを付けてください。

以下のコードを参考に条件を満たすExcelのマクロを作成してください。
# 条件

- calibrationCycleという変数を作成し、「Quotient & "年に1回"」の値を代入してください。
- TargetSheet.Cells(j, "G").Value = SourceSheet.Cells(i, "O").Valueの行に以下の文字を付け加えてください。
「に「calibrationCycle」との周期で校正を依頼する。」


以下の条件を満たすExcelのマクロを作成してください。
# 条件
- シート「確認用シート」から「予算案_提出用」へと値を転記する。
-「確認用シート」のC列の行数を数えてその行数の数だけ処理を上から行う。
- 「予算案_提出用」に必要な行数を新しく挿入する。
- スクリーンの更新と自動計算をオン・オフし、コード実行の速度を改善する。
- 「予算案_提出用」のA列を探索し、最終行の位置から挿入を開始してください
- 転記先と転記元のシート名をあとから変更しやすくしてください。
- 日本語のコメントを付けてください。

以下の条件を満たすようにコードを修正してください。
# 条件
- 「確認用シート」のB列の値を「予算案_提出用」のB列に転記する。
- 「確認用シート」のC列の値を「予算案_提出用」のC列に転記する。
- 「確認用シート」のD列の値を「予算案_提出用」のD列に転記する。
- 「確認用シート」のE列の値を「予算案_提出用」のE列に転記する。
- 「確認用シート」のF列の値を「予算案_提出用」のF列に転記する。
- 「確認用シート」のG列の値を「予算案_提出用」のG列に転記する。
- 新しく挿入した行のA列に「=ROW()-1」の関数を入力してください。
- 新しく挿入した行のH列に「=IFERROR(E2*F2,0)」の関数を入力してください。
- 「=IFERROR(G2*F2,0)」の行番号は行数によって変更してください。
- シートの構成が変わる場合もあるので1列ごとの処理を記述してください。


以下の条件を満たすExcelのマクロを作成してください。
# 条件
- シート「予算案_提出用」のフォーマットを規定します。
- A列の2行目から最終行までの範囲からB列、F列のフォントを「游ゴシック Medium」にしてください。大きさは11
- A列の2行目から最終行までの範囲からG列のフォントを「游ゴシック Light」にしてください。大きさは10
- A列の2行目から最終行までの範囲からA列とC~E列のフォントを「Meiryo UI」にしてください。
- A列の2行目から最終行までの範囲からA列に記載してある数字が奇数の場合はOdd、偶数の場合はEvenとしてください。
- Oddの行でA列、E列のセル色をRGB(242,242,242)、B~D、F~G列のセル色を白にしてください。
- Evenの行でA列、E列のセル色をRGB(198,224,180)、B~D、F~G列のセル色をRGB(226,239,218)にしてください。
- 変数名などは英語で記述してください。
- 日本語のコメントを付けてください。


読み解いたコードを参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- ファイル「社内校正予定一覧表.xlsm」のシート「確認用シート」からファイル「予算策定の提出シート.xlsm」のシート「予算案_提出用」へと値を転記する。
- 両方のファイルは同じディレクトリ「C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\品質管理部_予算策定の書類」にあります。
- 転記するデータは添付ファイルに従ってください。
- ファイル名やシート名は後で変更できるようにしてください。
- 変数名などは英語で記述してください。
- 日本語のコメントを付けてください。
- 転記するデータは添付ファイルに従ってください。



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
上記の表を参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- 型式名ごとに更新費用を計算します。
- 更新費用-1の計算式はB列からD列の「空白以外の値」の個数*「I列の値」です。(下記の表ならNo.100から順に、1, 2, 2, 3, 3になります)J列に記載します。
- 更新費用-2の計算式は 4 * 「I列の値」です。J列に記載します。
- 計算は一番下の行から行います。
- 更新費用-2が当てはまるのは、1列目の下と上の文字を比較し、4文字目から6文字目までが一致し、かつ
H列の値も一致し、かつともにD列に「T」から始まる値がある場合です。
- 更新費用-2が当てはまった行の上のJ列の行は空白にしてください。また、この処理はすべてのJ列に値を記入してから行ってください。
- 計算は1行毎に行い、J列に記載します。
- Range("A" & i)のようにどのセルを参照しているかがわかりやすい表記でお願いします。
- 日本語のコメントを付けてください。
- 変数名などは英語で記述してください。

下記のコードに以下の条件を加えてマクロを修正してください。


コード1の更新費用-2の条件にコード2の条件をあてはめてコードを完成させてください。
# 条件
- はじめに更新費用-1の条件でJ列を埋めてください。
- それから更新費用-2の条件で当てはまる行のデータを削除してください。
- コード1の該当する条件はすべて無視してください。

# コード1
Sub CalculateRenewalFees()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' 1つめのシートを操作対象とする

    Dim lastRow As Long
    Dim i As Long
    Dim cnt As Long
    Dim renewalFee1 As Double
    Dim renewalFee2 As Double
    Dim shouldSkip As Boolean

    ' 最後の行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 一番下の行から計算を始める
    For i = lastRow To 2 Step -1
        ' 初期値をセット
        cnt = 0
        shouldSkip = False
        renewalFee1 = 0
        renewalFee2 = 0

        ' B列からD列までの空白以外の値の数を計算（型式名ごと）
        If ws.Range("B" & i).Value <> "" Then cnt = cnt + 1
        If ws.Range("C" & i).Value <> "" Then cnt = cnt + 1
        If ws.Range("D" & i).Value <> "" Then cnt = cnt + 1

        ' 更新費用-1を計算
        renewalFee1 = cnt * ws.Range("I" & i).Value

        ' 更新費用-2の条件をチェック
        If i > 2 Then
            If Mid(ws.Range("A" & i).Value, 4, 3) = Mid(ws.Range("A" & (i - 1)).Value, 4, 3) And _
               ws.Range("H" & i).Value = ws.Range("H" & (i - 1)).Value And _
               Left(ws.Range("D" & i).Value, 1) = "T" And Left(ws.Range("D" & (i - 1)).Value, 1) = "T" Then

                ' 更新費用-2を計算
                renewalFee2 = 4 * ws.Range("I" & i).Value

                ' 上の行のJ列を空白にする
                shouldSkip = True
            End If
        End If

        ' J列に更新費用-1を記入
        ws.Range("J" & i).Value = renewalFee1

        ' J列に更新費用-2を記入（条件が当てはまる場合）
        If renewalFee2 > 0 Then
            ws.Range("J" & i).Value = renewalFee2
        End If

        ' J列を空白にする（条件が当てはまる場合）
        If shouldSkip Then
            ws.Range("J" & (i - 1)).Value = ""
        End If
    Next i
End Sub

# コード2
Sub CheckMatchingDRows()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    ' シートの初期設定
    Set ws = ThisWorkbook.Sheets("Sheet1")  ' 適切なシート名に変更してください

    ' 最後の行を探す
    lastRow = ws.Cells(Rows.Count, "D").End(xlUp).Row

    ' ループで各行をチェック
    For i = 2 To lastRow ' 2行目から開始して最後の行まで（1行目はヘッダーと仮定）
        ' D列のi行目とi-1行目が同じかどうか、かつ、空白でないかをチェック
        If ws.Range("D" & i).Value = ws.Range("D" & (i - 1)).Value And _
           ws.Range("D" & i).Value <> "" And ws.Range("D" & (i - 1)).Value <> "" Then
            ' 条件に合致した場合、下の行（i）の行番号をImmediate Windowに出力
            Debug.Print "Matching non-empty values found at row: " & i
        End If
    Next i
End Sub

下記のコードを参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- F列の11文字目までを比較し、同一なら同一型式とする。
- 3行目と4行目が同一型式の場合、3行目と4行目のA列の値を結合し、TargetSheetのB列に転記する。
- 3行目と4行目が同一型式の場合、4行目の処理はスキップする。

Sub MaskFees()
    ' "マスク型式一覧"から"確認用シート"にデータを移動
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim Quotient As Double
    Dim calibrationCycle As String
    Dim StartDate As Date ' 開始日を格納する変数
    Dim EndDate As Date   ' 終了日を格納する変数

    ' 日付範囲の設定
    StartDate = DateValue("2024-07-01")
    EndDate = DateValue("2025-06-30")

    ' 転記元と転記先のシート名を設定
    Set SourceSheet = ThisWorkbook.Sheets("マスク型式一覧")
    Set TargetSheet = ThisWorkbook.Sheets("確認用シート")

    ' 転記元シートの最終行を取得
    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, "A").End(xlUp).Row
    j = 2 ' 転記先の行番号の初期値を設定

    ' 転記元の行を上から順に確認
    For i = 1 To lastRow

        ' M列が日付かどうかを確認
        If IsDate(SourceSheet.Cells(i, "H").Value) Then
            ' M列の日付が指定範囲内であるか確認
            If SourceSheet.Cells(i, "H").Value >= StartDate And SourceSheet.Cells(i, "H").Value <= EndDate Then

                ' 登録期間の表示
                If IsNumeric(SourceSheet.Cells(i, "G").Value) Then
                    Quotient = SourceSheet.Cells(i, "G").Value / 12
                    calibrationCycle = Round(Quotient, 2) & "年に1回" ' 周期を設定
                    TargetSheet.Cells(j, "J").Value = "1/" & Quotient & "年" '計算して周期を表示
                End If

                ' B列とO列の値を転記
                If SourceSheet.Cells(i, "i").Value <> "" Then
                    TargetSheet.Cells(j, "G").Value = "検定番号" & SourceSheet.Cells(i, "I").Value & "として「" & calibrationCycle & "」の型式検定を行う。"
                Else
                    TargetSheet.Cells(j, "G").Value = ""
                End If

                TargetSheet.Cells(j, "B").Value = SourceSheet.Cells(i, "A").Value '機器の名称
                'TargetSheet.Cells(j, "A").Value = SourceSheet.Cells(i, "A").Value '管理番号

                ' M列の値をJ列に転記
                TargetSheet.Cells(j, "i").Value = SourceSheet.Cells(i, "H").Value '登録予定日
                TargetSheet.Cells(j, "D").Value = SourceSheet.Cells(i, "K").Value '単価
                ' 定数を設定
                TargetSheet.Cells(j, "C").Value = "1"
                TargetSheet.Cells(j, "E").Formula = "=IFERROR(D" & j & "*(1+F" & j & "),"""")"
                TargetSheet.Cells(j, "F").Value = 0.1
                TargetSheet.Cells(j, "H").Value = TargetSheet.Cells(j, "D").Value * TargetSheet.Cells(j, "F").Value
                TargetSheet.Cells(j, "K").Value = "マスク"
                ' 転記先の次の行に移動
                j = j + 1
            End If
        End If

    Next i
End Sub


以下の条件を満たすExcelのマクロを作成してください。
# 条件
- J列の値があるグループとないグループにわける。例：AとB
- J列の値があるグループのうち値が「I列の値＊4」のグループとそれ以外のグループにわける。例：1と2
- J列を上から見ていき、当てはまるものに対して、K列に「A1」、「A2」、「B1」、「B2」のように記載する。

以下のコードを参考に条件を満たすExcelのマクロを作成してください。
# 条件
- BからDまでの列を探索し、値が入っているセルの数を数える。
- K列がA2ならB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。
- K列がA2ならCells(i, "A")の値を、シート「確認用シート」のB列に転記する。
- K列がA1ならB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。
- K列がA1ならCells(i, "A")とCells(i+1, "A")の値を連結し、シート「確認用シート」のB列に転記する。
- K列がB1ならスキップする。
- 変数名などは英語で記述してください。
- コメントは日本語でお願いします。

このコードに以下の条件を満たすExcelのマクロを作成してください。
# 条件
- K列=A2かつA列の1文字目から6文字目までが同一かつH列の値を比較して同一ならば、以下の処理を行う。
- 検出した行のうち、上の行のK列をA3、下の行のK列をA4に変更する。
- K列がA3の行はCells(i, "A")とCells(i+1, "A")の値を連結し、シート「確認用シート」のB列に転記する。
- K列がA3の行はK列がA3とA4の行のB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。
- K列がA4の行はスキップする。
# コード

HelmetFeesをもとにProcessSheetsの内容を反映させてコードを完成してください。

Sub SetKColumnValues()
' K列に値をセットするためのプロシージャ
    Dim mainWs As Worksheet
    Dim lastRow As Long, i As Long

    ' シート設定
    Set mainWs = ThisWorkbook.Sheets("保護帽型式一覧")

    ' 最後の行を取得
    lastRow = mainWs.Cells(mainWs.Rows.count, "A").End(xlUp).Row

    ' J列の値に基づいてK列に値を設定
    For i = 1 To lastRow
        If IsEmpty(mainWs.Cells(i, "J")) Then
            mainWs.Cells(i, "K").Value = "B1"
        ElseIf IsNumeric(mainWs.Cells(i, "J").Value) And IsNumeric(mainWs.Cells(i, "I").Value) Then
            If mainWs.Cells(i, "J").Value = mainWs.Cells(i, "I").Value * 4 Then
                mainWs.Cells(i, "K").Value = "A1"
            Else
                mainWs.Cells(i, "K").Value = "A2"
            End If
        Else
            mainWs.Cells(i, "K").Value = "Error"
        End If
    Next i
End Sub

' K列の値に基づいて確認用シートに転記するためのプロシージャ
Sub ProcessSheets()
    ' 既存の処理（省略）

    ' A3, A4の処理
    For i = 2 To lastRow - 1
        If mainWs.Cells(i, "K").Value = "A2" And _
           Left(mainWs.Cells(i, "A").Value, 6) = Left(mainWs.Cells(i + 1, "A").Value, 6) And _
           mainWs.Cells(i, "H").Value = mainWs.Cells(i + 1, "H").Value Then

            mainWs.Cells(i, "K").Value = "A3"
            mainWs.Cells(i + 1, "K").Value = "A4"

        End If
    Next i
End Sub


以下の条件を満たすExcelのマクロを作成してください。
# 条件
- シート「保護帽型式一覧」から「確認用シート」へと値を転機する。
- BからDまでの列を探索し、値が入っているセルの数を数える。
- K列がA2ならB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。各値の間は「,」で区切り、最後に「,」はつけない。
- K列がA2ならCells(i, "A")の値を、シート「確認用シート」のB列に転記する。
- K列がA1ならB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。
- K列がA1ならCells(i, "A")とCells(i+1, "A")の値を連結し、シート「確認用シート」のB列に転記する。
- K列がA3の行はCells(i, "A")とCells(i+1, "A")の値を連結し、シート「確認用シート」のB列に転記する。
- K列がA3の行はK列がA3とA4の行のB~Dまでの値を変数に格納し、その値をシート「確認用シート」のG列に転記する。
- K列がA4、B1の行はスキップする。
- 変数名などは英語で記述してください。
- コメントは日本語でお願いします。

Sub CopyValues()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rowCount As Long, i As Long, targetRow As Long
    Dim cellValue As String, concatenatedValue As String, transposedValue As String

    ' シートを設定する
    Set wsSource = ThisWorkbook.Sheets("保護帽型式一覧")
    Set wsTarget = ThisWorkbook.Sheets("確認用シート")

    ' 確認用シートの次の空行を特定する
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row + 1

    ' 保護帽型式一覧シートの行数を取得する
    rowCount = wsSource.Cells(wsSource.Rows.Count, "K").End(xlUp).Row

    ' K列を走査する
    For i = 1 To rowCount
        ' K列の値を取得する
        cellValue = wsSource.Cells(i, "K").Value

        ' BからDまでの列の値を結合する
        transposedValue = wsSource.Cells(i, "B").Value & "," & wsSource.Cells(i, "C").Value & "," & wsSource.Cells(i, "D").Value
        transposedValue = Left(transposedValue, Len(transposedValue) - 1) ' 最後のカンマを削除

        ' K列の値に応じて処理をする
        Select Case cellValue
            Case "A1"
                ' A1の場合
                concatenatedValue = wsSource.Cells(i, "A").Value & wsSource.Cells(i + 1, "A").Value
                wsTarget.Cells(targetRow, "B").Value = concatenatedValue
                wsTarget.Cells(targetRow, "G").Value = transposedValue
                targetRow = targetRow + 1

            Case "A2"
                ' A2の場合
                wsTarget.Cells(targetRow, "B").Value = wsSource.Cells(i, "A").Value
                wsTarget.Cells(targetRow, "G").Value = transposedValue
                targetRow = targetRow + 1

            Case "A3"
                ' A3の場合
                concatenatedValue = wsSource.Cells(i, "A").Value & wsSource.Cells(i + 1, "A").Value
                wsTarget.Cells(targetRow, "B").Value = concatenatedValue

                ' A3とA4のBからDまでの値を結合する
                transposedValue = wsSource.Cells(i, "B").Value & "," & wsSource.Cells(i, "C").Value & "," & wsSource.Cells(i, "D").Value & ","
                transposedValue = transposedValue & wsSource.Cells(i + 1, "B").Value & "," & wsSource.Cells(i + 1, "C").Value & "," & wsSource.Cells(i + 1, "D").Value
                wsTarget.Cells(targetRow, "G").Value = transposedValue

                targetRow = targetRow + 1

            Case "A4", "B1"
                ' スキップ
        End Select
    Next i
End Sub


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
上記の表を参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- 2022年2月15日から2022年12月31日の範囲の登録日を持つ行を抜き出す。
- その行に対し、以下のコードの処理を行う。

# コード


以上のコードで以下の条件に従ってコードを修正してください。
# 条件
- "保護帽型式一覧"シートにおいてB列に新しく「販売」という列を設けました。それに伴い、B列以降が1列右にズレています。その変更に対応してください。
- "保護帽型式一覧"シートにおいてB列に「継続中」という値がある行だけを処理したいです。


以上のコードで以下の条件に従ってコードを修正してください。
# 条件
- "マスク型式一覧"シートのB列に「継続中」という値がある行だけを処理したいです。


|  A       |  B    |  C   |     D      |     H     |     I      |
|----------|-------|------|------------|-----------|------------|
| 型式名    | 100ms | 101ms| 102ms      | 103ms     | 104ms      |
|----------|-------|------|------------|-----------|------------|
| No.100   | 10    | 09   |  11        | 14        |  16        |
| No.105   | 08    | 06   |  11        | 14        |  16        |
| No.110   | 10    | 10   |  11        | 14        |  16        |
| No.110F  | 10    | 08   |  11        | 14        |  16        |
| No.110S  | 12    | 09   |  11        | 14        |  16        |
-----------------------------------------------------------------------


ExcelVBAで以下の要素を備えるユーザフォームで条件を満たすコードを作成してください。
# 要素
- ユーザフォームの名前は「frmDateRange」にする。
- 6つのテキストボックスがあり、名前はそれぞれ、YearStartBox,MouthStartBox,DayStartbox,YearEndBox,MouthEndBox,DayEndBoxにする。
- 2つのボタンがあり、それぞれOkButton,CancelButtonにする。
- YearStartBox,MouthStartBox,DayStartboxにはそれぞれ「2022」「1」「1」など西暦の年、月、日が入る。
- YearEndBox,MouthEndBox,DayEndBoxにはそれぞれ「2022」「12」「31」など西暦の年、月、日が入る。
- テキストボックスには初期値が設定されている。初期値は上記の例で構わない。
- テキストボックスの修正はリストと直接入力の両方が可能である。
- 全てのテキストボックスに数値が入ったときにOkButtonを押すと、変数StartDateには「2022-01-01」、変数EndDateには「2022-12-31」のように日付が入る。
- 変数に数値が入ると、ユーザーフォームが閉じる。


以上のコードを参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- 変数StartDate,EndDateを以下のプロシージャで使用したい。
- プロシージャの名前はそれぞれ、TransferAndFormatHelmetData, MaskFees, BicycleFeesです。
- それぞれのシートはあるシートからシートに値を転記するコードで、変数StartDate,EndDateで範囲を指定している。
- TransferAndFormatHelmetData, MaskFees, BicycleFeesで共通した変数を使用する工程を教えて下さい。
- なにか足りない情報があれば、お知らせください。


VBAProject内のユーザフォーム"frmDateRange"と標準モジュール内のUpdateSafeteyEquipmentFee内の
TransferAndFormatHelmetData, MaskFees, BicycleFeesのプロシージャ間でデータをやり取りしたいのです。



Sub GenerateCalibrationBudgetReport()
    ' "確認用シート"から"予算案_提出用"にデータを移動するマクロ

    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim lastRow As Long
    Dim StartRow As Long
    Dim InsertRowCount As Long
    Dim i As Long

    ' 転記元と転記先のシート名を変数に設定
    Set SourceSheet = ThisWorkbook.Sheets("確認用シート")
    Set TargetSheet = ThisWorkbook.Sheets("予算案_提出用")

    ' 転記元のC列の最終行を取得
    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, "C").End(xlUp).Row

    ' 転記先のA列の最終行の次の行を取得
    StartRow = TargetSheet.Cells(TargetSheet.Rows.Count, "A").End(xlUp).Row + 1

    ' 転記するデータの行数を計算
    InsertRowCount = lastRow - 1 ' 2行目から開始するため

    ' 必要な行数を転記先に挿入
    TargetSheet.Rows(StartRow & ":" & StartRow + InsertRowCount - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' スクリーンの更新と自動計算をオフにする
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 転記元のデータを転記先にコピー
    For i = 2 To 11 ' A列からK列まで
        SourceSheet.Cells(2, i).Resize(lastRow - 1, 1).Copy TargetSheet.Cells(StartRow, i)
    Next i

    ' スクリーンの更新と自動計算をオンにする
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
以上のコードを参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- TargetSheet = ThisWorkbook.Sheets("予算案_提出用")のシートにおいて、予算策定の提出Bookの"予算案_提出用"のシートに変更する。


Sub GenerateCalibrationBudgetReport()
    ' "確認用シート"から"予算案_提出用"にデータを移動するマクロ

    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim TargetWorkbook As Workbook
    Dim lastRow As Long
    Dim StartRow As Long
    Dim InsertRowCount As Long
    Dim i As Long
    Dim wbookName As String

    wbookName = "予算策定の提出Book.xlsx"

    ' 転記元のシート名を変数に設定
    Set SourceSheet = ThisWorkbook.Sheets("確認用シート")

    ' 転記先のワークブックが開いているか確認
    On Error Resume Next
    Set TargetWorkbook = Workbooks(wbookName)
    On Error GoTo 0

    ' もしワークブックが開いていなければ、ここで開く（指定したパスにファイルが存在する必要がある）
    If TargetWorkbook Is Nothing Then
        Set TargetWorkbook = Workbooks.Open("C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\品質管理部_予算策定の書類\" & "予算策定の提出シート.xlsm")
    End If

    ' 転記先のシート名を変数に設定
    Set TargetSheet = TargetWorkbook.Sheets("予算案_提出用")

    ' 転記元のC列の最終行を取得
    lastRow = SourceSheet.Cells(SourceSheet.Rows.Count, "C").End(xlUp).Row

    ' 転記先のA列の最終行の次の行を取得
    StartRow = TargetSheet.Cells(TargetSheet.Rows.Count, "A").End(xlUp).Row + 1

    ' 転記するデータの行数を計算
    InsertRowCount = lastRow - 1 ' 2行目から開始するため

    ' 必要な行数を転記先に挿入
    TargetSheet.Rows(StartRow & ":" & StartRow + InsertRowCount - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' スクリーンの更新と自動計算をオフにする
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 転記元のデータを転記先にコピー
    For i = 2 To 11 ' A列からK列まで
        SourceSheet.Cells(2, i).Resize(lastRow - 1, 1).Copy TargetSheet.Cells(StartRow, i)
    Next i

    ' スクリーンの更新と自動計算をオンにする
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

以上のコードを参考に以下の条件を満たすExcelのマクロを作成してください。
# 条件
- 関数'=IFERROR(sum(E2:E*),"")をC列に"小計"という文字列が入った行のE列に挿入する。
- 関数'=IFERROR(sum(E2:E*)*0.1,"")をC列に"消費税"という文字列が入った行のE列に挿入する。
- 関数'=IFERROR(SUM(E40:E41),"")をC列に"総計"という文字列が入った行のE列に挿入する。
- 予算案_提出用シートのF列を探索し、値が入っているセルの数を数える。
- sum(E2:E*)の*はそのセルの数である。

以上のコードを以下の条件に従い修正してください。
# 条件
- A列、C列、D列、E列、F列の幅をそれぞれ「60ピクセル」「75ピクセル」「90ピクセル」「160ピクセル」「92ピクセル」にする。
- B列、G列は自動調整する。
- E列の最終行と、最後から2番めの行は太字にする。
- F列の書式設定を[％]にする。


当社における製品の呼称として「BPB480 S, BPB480 M」, 「3123A691 S, 3123A691 M」,「BPB441 S, BPB441 M」, 「3123A692 S, 3123A692 M」, 「BPB442 S, BPB442 M」, 「3123A693 S, 3123A693 M」,「BPB540 S, BPB540 M」, 「3123A694 S, 3123A694 M」を用います。

以上のコードを以下の条件に従い修正してください。
# 条件
- H列で値がダブる場合はその行をスキップする。
- H列で値がダブる行のA列の値をつなげ、繋げた値を転記する

以上のコードを以下の条件に従い修正してください。
# 条件
- TargetSheet.Cells(j, "D").Value = SourceSheet.Cells(i, "W").Valueが発動する条件をSourceSheetのX列の値が空白の場合にする。
- 全ての転記作業が終わった後に、TargetSheetに次の1行を付け足す。
- A列"-", B列"重量計の校正", D列"54000", G列"服部に「1年に1回」の周期で校正を依頼する。"

以上のコードを以下の条件に従い修正してください。
# 条件
- rebateValueが3の場合に別の処理を加えたい。
- I列の値がStartDataとEndDataの間にある場合のみ、その行を転記する。
- I列の値が同じ場合はH列の数字が大きい行を残す。
- その他の処理はrevateValueが2や3以外の場合と同じにする。


Dim sumFeesTransferred As Boolean

Sub CalculateAndTransferJQAFees()
    Dim ws As Worksheet, wsTarget As Worksheet
    Dim targetRow As Long
    Dim StartDate As Date, EndDate As Date

    StartDate = DateValue("2024-07-01")
    EndDate = DateValue("2025-06-30")

    sumFeesTransferred = False

    Set ws = ThisWorkbook.Sheets("JIS・SGなど各種協会")
    Set wsTarget = ThisWorkbook.Sheets("確認用シート")

    targetRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row + 1

    TransferSumFees ws, wsTarget, targetRow
    TransferRows ws, wsTarget, StartDate, EndDate, targetRow
End Sub

Sub TransferSumFees(ByRef ws As Worksheet, ByRef wsTarget As Worksheet, ByRef targetRow As Long)
    Dim sumFees As Long
    If Not sumFeesTransferred Then
        sumFees = CalculateJISRegistrationFee(ws)
        With wsTarget
            .Cells(targetRow, "B").Value = "日本品質保証機構 認証登録維持料"
            .Cells(targetRow, "D").Value = sumFees
            .Cells(targetRow, "C").Value = 3
            .Cells(targetRow, "E").Formula = "=IFERROR(D" & targetRow & "*(1+F" & targetRow & "),"""")"
            .Cells(targetRow, "F").Value = 0.1
            .Cells(targetRow, "G").Value = "JIS保護メガネ、ベルトスリング、墜落制止用器具の維持費"
            .Cells(targetRow, "H").Value = .Cells(targetRow, "D").Value * .Cells(targetRow, "F").Value
            SetColumnsIJK ws, wsTarget, 0, targetRow
        End With
        sumFeesTransferred = True
        targetRow = targetRow + 1
    End If
End Sub

Function CalculateJISRegistrationFee(ByRef ws As Worksheet, ByVal StartDate As Date, ByVal EndDate As Date) As Long
    Dim i As Long, lastRow As Long, sumFees As Long, rebateValue As Long
    Dim dateValue As Date
    Dim dictDates As Object

    Set dictDates = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    sumFees = 0

    For i = 1 To lastRow
        If IsDate(ws.Cells(i, "I").Value) Then
            dateValue = ws.Cells(i, "I").Value
            ' Only consider rows with I column between StartDate and EndDate
            If dateValue >= StartDate And dateValue <= EndDate Then
                ' For rows with the same date in column I, retain the one with the higher value in column H
                If Not dictDates.Exists(dateValue) Or ws.Cells(i, "H").Value > dictDates(dateValue) Then
                    dictDates(dateValue) = ws.Cells(i, "H").Value

                    ' Different processing for rebateValue
                    rebateValue = ws.Cells(i, "K").Value
                    Select Case rebateValue
                        Case 2
                            sumFees = sumFees + ws.Cells(i, "J").Value
                        Case 3
                            sumFees = ws.Cells(i, "J").Value
                        Case Else
                            ' Processing for other values
                            sumFees = sumFees + ws.Cells(i, "J").Value
                    End Select
                End If
            End If
        End If
    Next i

    CalculateJISRegistrationFee = sumFees
    Set dictDates = Nothing
End Function

Sub TransferRows(ByRef ws As Worksheet, ByRef wsTarget As Worksheet, ByVal StartDate As Date, ByVal EndDate As Date, ByRef targetRow As Long)
    Dim lastRow As Long, i As Long

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 1 To lastRow
        If IsValidDate(ws.Cells(i, "I").Value, StartDate, EndDate) Then
            With wsTarget
                .Cells(targetRow, "B").Value = ws.Cells(i, "A").Value
                .Cells(targetRow, "D").Value = ws.Cells(i, "J").Value
                .Cells(targetRow, "C").Value = 1
                .Cells(targetRow, "E").Formula = "=IFERROR(D" & targetRow & "*(1+F" & targetRow & "),"""")"
                .Cells(targetRow, "F").Value = 0.1
                .Cells(targetRow, "G").Value = ws.Cells(i, "C").Value & "の費用"
                .Cells(targetRow, "H").Value = .Cells(targetRow, "D").Value * .Cells(targetRow, "F").Value
                SetColumnsIJK ws, wsTarget, i, targetRow
            End With
            targetRow = targetRow + 1
        End If
    Next i
End Sub

Function IsValidDate(ByVal cellValue As Variant, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    IsValidDate = False
    If IsDate(cellValue) Then
        If cellValue >= StartDate And cellValue <= EndDate Then
            IsValidDate = True
        End If
    End If
End Function

Sub SetColumnsIJK(ByRef SourceSheet As Worksheet, ByRef TargetSheet As Worksheet, ByVal sourceRow As Long, ByVal targetRow As Long)
    If sourceRow = 0 Then
        TargetSheet.Cells(targetRow, "I").Value = "?"
        TargetSheet.Cells(targetRow, "J").Value = "1/1年"
        TargetSheet.Cells(targetRow, "K").Value = "日本品質保証機構"
    Else
        TargetSheet.Cells(targetRow, "I").Value = SourceSheet.Cells(sourceRow, "I").Value
        TargetSheet.Cells(targetRow, "J").Value = "1/1年"
        TargetSheet.Cells(targetRow, "K").Value = SourceSheet.Cells(sourceRow, "C").Value
    End If
End Sub
