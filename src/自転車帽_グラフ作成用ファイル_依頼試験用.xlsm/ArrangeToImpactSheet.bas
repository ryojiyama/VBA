Attribute VB_Name = "ArrangeToImpactSheet"
Option Explicit




' 出来上がった"レポートグラフ"シートに各値を配置する
Sub ArrangeDataByGroup()
    Dim wsName As String: wsName = "レポートグラフ" ' シート名に含まれる部分文字列
    Dim wsResult As Worksheet
    Dim lastRow As Long

    ' "Impact"を含むすべてのワークシートをループ
    For Each wsResult In ThisWorkbook.Worksheets
        If InStr(wsResult.Name, wsName) > 0 Then
            ' ワークシートの最終使用行を取得
            lastRow = wsResult.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

            ' 列Iに基づいてグループを処理
            ProcessGroupsInColumnI wsResult, lastRow
        End If
    Next wsResult
End Sub

Private Sub ProcessGroupsInColumnI(ws As Worksheet, lastRow As Long)
    'ArrangeDataByGroupのサブプロシージャ。I列の値でグループを作成
    Dim firstRow As Long: firstRow = 2
    Dim groupInsert As Variant
    Dim i As Long
    Dim groupRange As Range

    ' 列Iのグループを特定するために各行をループ
    Do While firstRow <= lastRow
        groupInsert = GetGroupNumber(ws.Cells(firstRow, "I").value)

        ' グループ番号が空白や無効の場合は次の行に進む
        If Not IsNull(groupInsert) And groupInsert <> "" Then

            ' 現在のグループの最終行を見つける
            For i = firstRow + 1 To lastRow
                ' I列の値が空白の場合、ループを終了
                If ws.Cells(i, "I").value = "" Then Exit For
                ' I列の値が次のグループに変わったらループを終了
                If GetGroupNumber(ws.Cells(i, "I").value) <> groupInsert Then Exit For
            Next i

            ' デバッグ: グループの範囲を出力
            ' Debug.Print "グループ範囲: A" & firstRow & ":G" & i - 1

            ' 現在のグループの範囲を設定
            Set groupRange = ws.Range("A" & firstRow & ":G" & i - 1)
            ' 列Aに基づいてグループを処理
            ProcessGroupsInColumnA ws, groupInsert, groupRange

            ' 次のグループへ
            firstRow = i
        Else
            ' groupInsertがNullまたは空の場合、次の行へ
            firstRow = firstRow + 1
        End If
    Loop
End Sub


Private Sub ProcessGroupsInColumnA(ws As Worksheet, groupInsert As Variant, groupRange As Range)
    ' ArrangeDataByGroupのサブプロシージャ。列Iに基づいてグループを処理
    Dim groupFirstRow As Long: groupFirstRow = 2
    Dim groupResults As Variant
    Dim j As Long
    Dim lastRowA As Long

    ' 列Aの最終使用行を取得
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' 列Aのグループを特定するために各行をループ
    Do While groupFirstRow <= lastRowA
        groupResults = GetGroupNumber(ws.Cells(groupFirstRow, "A").value)

        ' グループ番号が空白でない場合に処理を行う
        If Not IsNull(groupResults) And groupResults <> "" Then
            ' 現在のグループの最終行を見つける
            For j = groupFirstRow + 1 To lastRowA + 1
                If j > lastRowA Or GetGroupNumber(ws.Cells(j, "A").value) <> groupResults Then Exit For
            Next j

            ' グループのサイズを計算
            Dim groupSize As Long
            groupSize = j - groupFirstRow
'            Debug.Print "FirstRow:" & groupFirstRow
'            Debug.Print "Size:" & groupSize

            ' グループサイズに応じてデータを配置
            If groupResults = groupInsert Then
                Select Case groupSize
                ' groupRange.Cells:挿入した表内, ws.Cells:Group以下の結果一覧表
                    Case 4 ' グループに2つのレコードがある場合
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value ' A列のGroupの値
                        groupRange.Cells(3, 1).value = ws.Cells(groupFirstRow, "G").value ' 前処理
                        groupRange.Cells(1, 3).value = ws.Cells(groupFirstRow, "D").value
                        groupRange.Cells(2, 3).value = ws.Cells(groupFirstRow, "F").value
                        groupRange.Cells(1, 4).value = ws.Cells(groupFirstRow, "B").value ' 左上の衝撃値
                        groupRange.Cells(1, 6).value = ws.Cells(groupFirstRow + 1, "D").value
                        groupRange.Cells(2, 6).value = ws.Cells(groupFirstRow + 1, "F").value
                        groupRange.Cells(1, 7).value = ws.Cells(groupFirstRow + 1, "B").value ' 右上の衝撃値
                        groupRange.Cells(4, 3).value = ws.Cells(groupFirstRow + 2, "D").value
                        groupRange.Cells(5, 3).value = ws.Cells(groupFirstRow + 2, "F").value
                        groupRange.Cells(4, 4).value = ws.Cells(groupFirstRow + 2, "B").value ' 左下の衝撃値
                        groupRange.Cells(4, 6).value = ws.Cells(j - 1, "D").value
                        groupRange.Cells(5, 6).value = ws.Cells(j - 1, "F").value
                        groupRange.Cells(4, 7).value = ws.Cells(j - 1, "B").value ' 右下の衝撃値
                        
                        'groupRange.Cells(2, 3).value = ws.Cells(j - 1, "C").value & ws.Cells(j - 1, "D").value 'j基点の位置合わせ。念の為保留
                    Case 3 ' 場合分けがあるときのために一時保留
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
                End Select

                ' セルの書式を設定
                Call FormatGroupRange(groupRange)
            End If

            ' 次のグループへ
            groupFirstRow = j
        Else
            ' groupResultsがNullまたは空白の場合、次の行へ
            groupFirstRow = groupFirstRow + 1
        End If
    Loop
End Sub


Private Sub FormatGroupRange(groupRange As Range)
    ' groupRangeの開始行をデバッグ出力
    Debug.Print "groupRangeの開始行: " & groupRange.row
    
    Dim ws As Worksheet
    Set ws = groupRange.Worksheet

    Dim rowIndex As Range
    Dim headerRange1 As Range
    Dim headerRange2 As Range
    Dim headerInput1 As Range
    Dim headerInput2 As Range
    Dim impactValue As Range
    Dim fontRange As Range

    ' ワークシート上の絶対的なセル範囲を取得
    With ws
        ' groupRangeの開始行の1列目から7列目までの範囲
        Set rowIndex = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 5, 1))
        Set headerRange1 = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 2, 7))
        Set headerRange2 = .Range(.Cells(groupRange.row + 3, 1), Cells(groupRange.row + 5, 7))
    
        ' headerInput1: 2列目から6列目の列全体
        Set headerInput1 = .Columns("B:F")
        ' headerInput2: 2列目と5列目の単独の列全体
        Set headerInput2 = Union(.Columns("B"), .Columns("E"))
        Set impactValue = Union(.Cells(groupRange.row, 4), .Cells(groupRange.row, 7), .Cells(groupRange.row + 3, 4), .Cells(groupRange.row + 3, 7))
    End With


    ' 範囲作成サンプル_保留しておく。
'    Set fontRange = Union(headerRange, leftColumnRange1)

    ' fontRange1 に対して書式設定
    With headerRange1.Font
        .Name = "UDEV Gothic"
        .Color = RGB(60, 60, 60) ' フォントの色を白に設定
    End With
    With headerRange2.Font
        .Name = "UDEV Gothic"
        .Color = RGB(60, 60, 60) ' フォントの色を白に設定
    End With
    
    With headerInput1
        .HorizontalAlignment = xlCenter ' 水平方向の中央揃え
        .VerticalAlignment = xlCenter ' 垂直方向の中央揃え
    End With
    With headerInput2
        .HorizontalAlignment = xlLeft ' 水平方向の中央揃え
        .VerticalAlignment = xlCenter ' 垂直方向の中央揃え
    End With
    
    With impactValue
        .NumberFormat = "0"" G""" ' 数値フォーマット 0.0"G" を追加
    End With
    
    With rowIndex.Font
        .Color = RGB(230, 230, 230)
    End With
    With rowIndex.Interior
        .Color = RGB(48, 84, 150) ' 背景色を青に設定
    End With

End Sub


Private Function GetGroupNumber(cellValue As String) As Variant
    ' ArrangeDataByGroupのサブプロシージャ。I列の値が次のグループに変わったらループを終了
    Dim regex As Object, matches As Object
    Dim result As String

    ' 空白セルやGroupの処理
    If cellValue = "" Or cellValue = "Group" Then
        GetGroupNumber = Null
        Exit Function
    End If

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "\D" ' 数字以外の文字にマッチするパターン
    End With

    ' 数字以外の文字を空文字に置換
    result = regex.Replace(cellValue, "")

    ' 結果が空文字でない場合、数字を返す
    If result <> "" Then
        GetGroupNumber = result
    Else
        GetGroupNumber = Null ' 数字が見つからない場合はNullを返す
    End If
End Function



' "LOG_Bicycel"シートのチャートを"レポートグラフ"シートに移動する。
' チャートの出現位置はサブルーチンで設定している。
Sub MoveChartsFromLOGToReport()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim chartObj As ChartObject
    Dim groupCell As Range
    Dim targetTop As Double
    Dim targetLeft As Double
    Dim offsetX As Double
    Dim offsetY As Double
    Dim i As Integer
    Dim recordID As String
    Dim idNumber As Integer ' IDの最初の数値部分
    Dim chartHeight As Double
    Dim chartWidth As Double
    Dim previousLeft As Double
    previousLeft = 0

    ' シートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Set wsTarget = ThisWorkbook.Sheets("レポートグラフ")

    ' A列で"Group"という値を探す
    Set groupCell = wsTarget.Columns("A").Find(What:="Group", LookIn:=xlValues, LookAt:=xlWhole)
    If groupCell Is Nothing Then
        MsgBox "レポートグラフシートのA列に'Group'が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' チャートの設置基準のオフセット（ピクセル単位）
    offsetY = 30 ' 上方向に30ピクセル
    offsetX = 10 ' 各チャートを右方向に10ピクセルずつずらす

    ' 縦横比1:2に設定するためのサイズ
    chartHeight = 200 ' 高さを200ピクセルに設定
    chartWidth = chartHeight * 2 ' 幅は高さの2倍に設定

    ' チャートの移動
    i = 0
    For Each chartObj In wsSource.ChartObjects
        ' チャートのタイトルからIDを取得
        If chartObj.chart.HasTitle Then
            recordID = Replace(chartObj.chart.chartTitle.Text, "ID: ", "") ' "ID: "を除去してIDのみ取得
            Debug.Print "recordID: " & recordID ' イミディエイトウィンドウにIDを出力

            ' IDの最初の数値部分を抽出
            idNumber = CInt(Split(recordID, "-")(0)) ' recordIDの最初の部分を数値化

            ' チャートの位置をサブプロシージャで設定
            Call SetChartPosition(idNumber, i, groupCell.Left, targetTop, targetLeft, previousLeft)
            ' チャートをコピー
            chartObj.Copy

            ' コピーのタイムラグを作成
            WaitHalfASecond

            ' レポートグラフシートをアクティブにして、チャートを貼り付け
            wsTarget.Activate
            wsTarget.Paste

            ' 貼り付けられたチャートのオブジェクトを取得
            With wsTarget.ChartObjects(wsTarget.ChartObjects.Count)
                ' チャートの位置を設定
                .Top = targetTop
                .Left = targetLeft

                ' チャートのサイズを設定 (縦横比 1:2)
                .Height = chartHeight
                .Width = chartWidth
                ' チャートの位置を設定
                Call SetChartPosition(idNumber, i, groupCell.Left, targetTop, targetLeft, previousLeft)
                previousLeft = targetLeft
            End With

            ' 次のチャート位置を右にずらす
            i = i + 1
        End If
    Next chartObj

    MsgBox "チャートの移動が完了しました。", vbInformation

    ' メモリ解放
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set chartObj = Nothing
    Set groupCell = Nothing
End Sub

Sub SetChartPosition(ByVal idNumber As Integer, ByVal chartIndex As Integer, ByVal groupLeft As Double, ByRef targetTop As Double, ByRef targetLeft As Double, ByVal previousLeft As Double)
    ' idNumber に基づいて高さを動的に計算
    targetTop = 100 + (idNumber - 1) * 200 ' idNumberが増えるごとに200ピクセルずつ高さを変える
    
    ' 横方向の位置をchartIndex に基づいて調整
    If targetLeft <= previousLeft Then
        targetLeft = previousLeft + 15 ' 同じ高さの場合でも徐々に横方向にずらす
    Else
        targetLeft = groupLeft + 400 + (chartIndex Mod 9) * 5 ' 9つのチャートごとに横にずらす
    End If
    
    Debug.Print "Index:"; chartIndex & " targetTop:"; targetTop & " targetLeft:"; targetLeft
End Sub

Sub WaitHalfASecond()
    Dim start As Single
    start = Timer
    Do While Timer < start + 0.4
        DoEvents ' リソース解放
    Loop
End Sub




' データ振り分けの確認をデバッグウインドウで行う。移行が中途半端で終わっている。
Sub ConsolidateData()

  ' シート名の変数を定義
  Dim wsName As String: wsName = "レポートグラフ" ' シート名の一部を指定
  Dim wsResult As Worksheet
  Dim i As Long, j As Long
  Dim groupInsert As Variant, groupResults As Variant
  Dim insertNum As Long, resultNum As Long
  Dim dict As Object

  ' "レポートグラフ"を含むシートを順に処理
  For Each wsResult In ThisWorkbook.Worksheets
    If InStr(wsResult.Name, wsName) > 0 Then

      ' 最終行を取得 (A列とI列の両方で値が入力されている最後の行を取得)
      Dim lastRow As Long: lastRow = wsResult.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

      ' グループごとにデータを格納する辞書を初期化
      Set dict = CreateObject("Scripting.Dictionary")

      ' I列とA列の値に基づいてグループ化し、辞書に格納
      For i = 2 To lastRow
        groupInsert = GetGroupNumber(wsResult.Cells(i, "I").value)
        groupResults = GetGroupNumber(wsResult.Cells(i, "A").value)

        ' デバッグ: I列とA列の値、グループ番号を表示
        Debug.Print "Sheet: " & wsResult.Name & ", Row: " & i & ", I Column: " & wsResult.Cells(i, "I").value & ", GroupInsert: " & groupInsert
        Debug.Print "Sheet: " & wsResult.Name & ", Row: " & i & ", A Column: " & wsResult.Cells(i, "A").value & ", GroupResults: " & groupResults

        ' グループが一致し、かつGroupInsertとGroupResultsが両方ともNullでない場合、辞書にデータを追加
        If Not IsNull(groupInsert) And Not IsNull(groupResults) And groupInsert = groupResults Then
          If Not dict.Exists(groupInsert) Then
            dict.Add groupInsert, New Collection
          End If
          
          ' デバッグ: 辞書に追加されるデータを確認
          Debug.Print "Adding value to group: " & groupInsert & ", Value: " & wsResult.Cells(i, "C").value
          
          dict(groupInsert).Add wsResult.Cells(i, "C").value
        End If
      Next i

      ' デバッグ: グループごとにD列の値を表示
      Debug.Print "Sheet: " & wsResult.Name & ", Grouped Data:"
      For Each groupInsert In dict.Keys
        Debug.Print "Group: " & groupInsert & ", Values: ";
        ' Collectionの要素をループ処理するための変数を定義
        Dim item As Variant
        For Each item In dict(groupInsert)
          Debug.Print item & " ";
        Next item
        Debug.Print
      Next groupInsert

    End If
  Next wsResult

End Sub
