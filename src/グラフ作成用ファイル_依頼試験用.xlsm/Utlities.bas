Attribute VB_Name = "Utlities"
' Impactシートとレポート本文内の表を削除する。
Sub CleanUpSheetsByName()
    Call DeleteImpactSheets
    Call DeleteInsertedRows
End Sub

' Impactシートのグラフ付きのテーブルを削除する。
Sub DeleteInsertRows()

  ' シート名に"Impact"を含むシートをループ処理
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
    If InStr(ws.Name, "Impact") > 0 Then
      ' シートをアクティブにする
      ws.Activate

      ' I列に"Insert" + 数字が入っている行を逆順にループ処理
      Dim lastRow As Long
      lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
      Dim i As Long
      For i = lastRow To 1 Step -1
        ' I列の値が "Insert" で始まり、その後に数字が続く場合に削除
        If ws.Cells(i, "I").value Like "Insert[0-9]*" Then
          ' 行全体を削除
          ws.Rows(i).Delete
        End If
      Next i
    End If
  Next ws

End Sub

Sub DeleteRowsAfterGroup_ImpactSheets()

    ' "Impact"を含むシート内をループ処理
    Dim ws As Worksheet
    Dim groupRowNumber As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' 全ワークシートをループ処理
    For Each ws In ThisWorkbook.Worksheets
        ' シート名に "Impact" が含まれるシートを対象にする
        If InStr(ws.Name, "Impact") > 0 Then
            ' A列に "Group" と書いている行を見つける
            lastRow = ws.UsedRange.Rows.count
            groupRowNumber = 0
            For i = 1 To lastRow
                If ws.Cells(i, "A").value = "Group" Then
                    groupRowNumber = i
                    Exit For
                End If
            Next i
            
            ' "Group"が見つかった場合
            If groupRowNumber > 0 Then
                ' 削除する範囲が1行以上あることを確認
                If groupRowNumber + 1 <= lastRow Then
                    ws.Rows(groupRowNumber + 1 & ":" & lastRow).EntireRow.Delete
                End If
            Else
                ' "Group"が見つからなかった場合の処理 (例: メッセージボックスを表示)
                MsgBox "シート '" & ws.Name & "' に 'Group' が見つかりませんでした。", vbExclamation
            End If
        End If
    Next ws

End Sub


Sub PrintedReportSheets()
    Call PrintImpactSheet
    Call PrintSideImpactSheet
End Sub

Sub PrintImpactSheet()
    Dim ws As Worksheet
    Dim sheetNames1 As Variant
    Dim sheetFound As Boolean
    Dim i As Long
    Dim sheetName As String

    ' 条件1: 特定のシートを印刷 ("Impact_Top", "Impact_Front", "Impact_Back", "レポート本文"を含む)
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back", "レポート本文")

    ' 各シートを検索して印刷
    For i = LBound(sheetNames1) To UBound(sheetNames1)
        sheetName = sheetNames1(i)
        sheetFound = False
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = sheetName Then
                ws.PrintOut From:=1, To:=1
                sheetFound = True
                Exit For
            End If
        Next ws
        ' シートが見つからない場合はメッセージを表示
        If Not sheetFound Then
            MsgBox "シート '" & sheetName & "' が見つかりません。", vbExclamation
        End If
    Next i
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    Dim sheetFound As Boolean

    ' 条件2: "Impact_Side"を名前に含むシートを印刷
    sheetFound = False
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "Impact_Side") > 0 Then
            ws.PrintOut From:=1, To:=1
            sheetFound = True
        End If
    Next ws
    
    ' "Impact_Side"シートが見つからなかった場合のメッセージ
    If Not sheetFound Then
        MsgBox "シート名に 'Impact_Side' を含むシートが見つかりません。", vbExclamation
    End If

    ' メモリ解放
    Set ws = Nothing
End Sub


' Impactを含むシート名の調整
Sub DeleteRowsBelowHeader()
    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim sheetName As String

    ' ワークシートをループ
    For Each ws In ThisWorkbook.Worksheets
        ' シート名に"Impact"が含まれているかチェック
        If InStr(ws.Name, "Impact") > 0 Then
            ' ヘッダーの下の行から最終行までを削除
            ws.Rows("15:" & ws.Rows.count).Delete
        End If
    Next ws
End Sub


Sub PrintChartIDs()
    Dim ws As Worksheet
    Dim chtObj As ChartObject

    ' 各ワークシートをループ
    For Each ws In ThisWorkbook.Worksheets
        ' 各ワークシート内のチャートオブジェクトをループ
        For Each chtObj In ws.ChartObjects
            ' CreateChartID関数を使用して "Chart ID" を生成
            Dim chartID As String
            chartID = CreateChartID(chtObj.chart.ChartArea.TopLeftCell)

            ' イミディエイトウィンドウに出力
            Debug.Print "Chart Name: " & chtObj.Name & ", Chart ID: " & chartID
        Next chtObj
    Next ws
End Sub

' アクティブシートのI列に "Insert" と印がついている行を削除する
Private Sub DeleteInsertedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' "レポート本文"シートを取得
    Set ws = ThisWorkbook.Sheets("レポート本文")
    
    ' I列の最終行を取得
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    
    ' 最後の行から1行ずつ上に向かって削除を確認
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "I").value, 6) = "Insert" Then
            ws.Rows(currentRow).Delete
        End If
    Next currentRow
End Sub

' シート名に"Impact"とついているシートを削除する。
Sub DeleteImpactSheets()
    Dim ws As Worksheet
    Dim sheetNamesToDelete As Collection
    Dim sheetName As String
    Dim i As Long
    
    ' 削除対象のシート名を一時的に保持するコレクションを作成
    Set sheetNamesToDelete = New Collection
    
    ' ワークシートをループ
    For Each ws In ThisWorkbook.Worksheets
        ' シート名に"Impact"が含まれているかチェック
        If InStr(ws.Name, "Impact") > 0 Then
            ' 削除対象のシート名をコレクションに追加
            sheetNamesToDelete.Add ws.Name
        End If
    Next ws
    
    ' コレクション内のシートを削除
    For i = sheetNamesToDelete.count To 1 Step -1
        ThisWorkbook.Sheets(sheetNamesToDelete(i)).Delete
    Next i
End Sub

