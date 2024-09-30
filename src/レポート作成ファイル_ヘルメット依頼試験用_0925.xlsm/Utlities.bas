Attribute VB_Name = "Utlities"


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




Sub PrintImpactSheet()
    Dim ws As Worksheet
    
    ' 条件1: 特定のシートを印刷
    Dim sheetNames1 As Variant
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    For Each ws In ThisWorkbook.Sheets
        If foundSheetName(ws.Name, sheetNames1) Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    
    ' 条件2: "Impact_Side"を名前に含むシートを印刷
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "Impact_Side") > 0 Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Function foundSheetName(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            foundSheetName = True
            Exit Function
        End If
    Next i
    foundSheetName = False
End Function
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


Private Sub DeleteInsertedRows()
    ' アクティブシートのI列に "Insert" と印がついている行を削除する
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' アクティブシートを取得
    Set ws = ActiveSheet
    
    ' I列の最終行を取得
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    
    ' 最後の行から1行ずつ上に向かって削除を確認
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "I").value, 6) = "Insert" Then
            ws.Rows(currentRow).Delete
        End If
    Next currentRow
End Sub

