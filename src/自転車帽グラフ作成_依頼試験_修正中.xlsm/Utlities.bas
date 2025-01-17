Attribute VB_Name = "Utlities"

Public Sub DeleteReportGraphSheets()
    Dim ws As Worksheet
    Dim i As Long
    
    ' 後ろから前にループしてシートを削除
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        
        ' シート名に「レポートグラフ」が含まれているシートを削除
        If InStr(ws.Name, "レポートグラフ") > 0 Then
            Application.DisplayAlerts = False  ' 削除確認メッセージを表示しない
            ws.Delete
            Application.DisplayAlerts = True   ' 削除確認メッセージの表示を元に戻す
        End If
    Next i
End Sub

' "レポート本文"シートのL列に "Insert" と印がついている行を削除する
Public Sub DeleteInsertedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' "レポート本文"シートを取得
    Set ws = ThisWorkbook.Sheets("レポート本文")
    
    ' I列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).row
    
    ' 最後の行から1行ずつ上に向かって削除を確認
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "L").value, 6) = "Insert" Then
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
    For i = sheetNamesToDelete.Count To 1 Step -1
        ThisWorkbook.Sheets(sheetNamesToDelete(i)).Delete
    Next i
End Sub
