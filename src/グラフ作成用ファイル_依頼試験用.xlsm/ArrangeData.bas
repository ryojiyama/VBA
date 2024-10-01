Attribute VB_Name = "ArrangeData"
Option Explicit


' データ振り分けの確認をデバッグウインドウで行う
Sub ConsolidateData()

  ' シート名の変数を定義
  Dim wsName As String: wsName = "Impact" ' シート名の一部を指定
  Dim wsResult As Worksheet
  Dim i As Long, j As Long
  Dim groupInsert As Variant, groupResults As Variant
  Dim insertNum As Long, resultNum As Long
  Dim dict As Object

  ' "Impact"を含むシートを順に処理
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
          dict(groupInsert).Add wsResult.Cells(i, "D").value
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

' 出来上がった"impact"シートに各値を配置する
Sub ArrangeDataByGroup()
    Dim wsName As String: wsName = "Impact" ' シート名に含まれる部分文字列
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
    Call InsertTextInMergedCells
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
    lastRowA = ws.Cells(ws.Rows.count, "A").End(xlUp).row

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

            ' グループサイズに応じてデータを配置
            If groupResults = groupInsert Then
                Select Case groupSize
                    Case 2 ' グループに2つのレコードがある場合
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
                        groupRange.Cells(3, 2).value = ws.Cells(groupFirstRow, "E").value
                        groupRange.Cells(3, 5).value = ws.Cells(j - 1, "E").value
                        groupRange.Cells(1, 2).value = ws.Cells(groupFirstRow, "B").value & ws.Cells(groupFirstRow, "C").value
                        groupRange.Cells(1, 5).value = ws.Cells(j - 1, "B").value & ws.Cells(j - 1, "C").value
                    Case 3 ' グループに3つのレコードがある場合
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
                        groupRange.Cells(3, 2).value = ws.Cells(groupFirstRow, "E").value
                        groupRange.Cells(3, 4).value = ws.Cells(groupFirstRow + 1, "E").value
                        groupRange.Cells(3, 6).value = ws.Cells(j - 1, "E").value
                        groupRange.Cells(1, 2).value = ws.Cells(groupFirstRow, "B").value & ws.Cells(groupFirstRow, "C").value
                        groupRange.Cells(1, 4).value = ws.Cells(groupFirstRow + 1, "B").value & ws.Cells(groupFirstRow + 1, "C").value
                        groupRange.Cells(1, 6).value = ws.Cells(j - 1, "B").value & ws.Cells(j - 1, "C").value
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
    ' セルの書式設定を行うProcessGroupsInColumnAのサブルーチン
    Dim ws As Worksheet
    Set ws = groupRange.Worksheet

    Dim headerRange As Range
    Dim leftColumnRange As Range
    Dim fontRange As Range

    ' ワークシート上の絶対的なセル範囲を取得
    With ws
        ' groupRangeの開始行の1列目から7列目までの範囲
        Set headerRange = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row, 7))
        ' groupRangeの開始行から2行下までの1列目の範囲
        Set leftColumnRange = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 2, 1))
    End With

    ' 上記の範囲を結合
    Set fontRange = Union(headerRange, leftColumnRange)

    ' セル色をRGB(48,84,150)に設定
    With headerRange.Interior
        .color = RGB(48, 84, 150)
    End With

    With leftColumnRange.Interior
        .color = RGB(48, 84, 150)
    End With

    ' フォントを"UDEV Gothic"に設定し、フォントの色を白に設定
    With fontRange.Font
        .Name = "UDEV Gothic"
        .color = RGB(255, 255, 255) ' フォントの色を白に設定
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




' 各シートに表題をつける
Sub InsertTextInMergedCells()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim textToInsert As String
    Dim sheetDict As Object

    ' シート名と対応するテキストの辞書を作成
    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "Impact_Top", "天頂部衝撃試験"
    sheetDict.Add "Impact_Front", "前頭部衝撃試験"
    sheetDict.Add "Impact_Back", "後頭部衝撃試験"
    sheetDict.Add "Impact_Side", "側頭部衝撃試験"

    ' 各シートをループ
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        
        ' シート名に"Impact"が含まれている場合のみ処理
        If InStr(sheetName, "Impact") > 0 Then
            ' シート名に基づいて挿入するテキストを決定
            If sheetDict.Exists(sheetName) Then
                textToInsert = sheetDict(sheetName)
                
                ' Cells(1,2)~Cells(1,7)を結合
                With ws.Range(ws.Cells(1, 2), ws.Cells(1, 7))
                    .Merge
                    .value = textToInsert ' シート名に対応するテキストを挿入
                    .HorizontalAlignment = xlCenter ' テキストを中央揃え
                    .VerticalAlignment = xlCenter   ' テキストを縦中央揃え
                    .Font.Name = "游ゴシック"         ' フォントを"游ゴシック"に設定
                    .Font.size = 20                   ' フォントサイズを20に設定
                    .Font.Bold = True
                End With
                
                ' 行の高さを50に設定
                ws.Rows(1).RowHeight = 50
            End If
        End If
    Next ws
End Sub

' チャートを各シートに分配する。
Sub DistributeChartsToRequestedSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
    Dim key As Variant
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"シートを対象にする
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"シートのチャートオブジェクトをグループ分け
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.Text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartNameを"-"で分割し、sheetNameを取得
        parts = Split(chartObj.Name, "-")
        If UBound(parts) >= 2 Then
            ' sheetNameを実際のシート名に変換
            Select Case parts(2)
                Case "天"
                    sheetName = "Impact_Top"
                Case "前"
                    sheetName = "Impact_Front"
                Case "後"
                    sheetName = "Impact_Back"
                Case Else
                    sheetName = parts(2) ' それ以外の場合はそのまま
            End Select
        Else
            sheetName = parts(0)
        End If
        
        ' 実際のシート名をキーとしてディクショナリに追加
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' グループごとにチャートを対応するシートにコピー
    For Each key In groups.Keys
        ' シートの存在を確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' シートが存在しない場合、チャートをコピーしない
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' チャートのコピー
            Dim chart As ChartObject
            Dim newChart As ChartObject
            For Each chart In groups(key)
                ' チャートをコピー
                chart.Copy
                WaitHalfASecond ' 0.5秒待機
                ' コピーしたチャートを貼り付け、戻り値として新しいチャートオブジェクトを取得
                targetSheet.Paste
                Set newChart = targetSheet.ChartObjects(targetSheet.ChartObjects.count)
                
                ' 元のチャートの位置に基づいて、右下に相対的に移動
                With newChart
                    .Top = chart.Top + 50  ' 元のチャートの位置から50ポイント下に移動
                    .Left = chart.Left + 100 ' 元のチャートの位置から100ポイント右に移動
                End With
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not copied."
        End If
    Next key
End Sub

Sub WaitHalfASecond()
    Dim start As Single
    start = Timer
    Do While Timer < start + 0.4
        DoEvents ' リソース解放
    Loop
End Sub

