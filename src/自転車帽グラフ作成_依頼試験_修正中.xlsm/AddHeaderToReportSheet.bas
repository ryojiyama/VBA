Attribute VB_Name = "AddHeaderToReportSheet"
Option Explicit

Public Sub InsertHeaderRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCell As Range
    Dim previousValue As String
    Dim currentValue As String
    Dim insertRow As Long
    Dim newValue As String
    
    ' "レポートグラフ"シートを設定
    Set ws = ThisWorkbook.Worksheets("レポートグラフ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    
    previousValue = ""
    insertRow = 0
    
    ' I列を最終行から上に向かって探索
    For i = lastRow To 1 Step -1
        Set currentCell = ws.Cells(i, "I")
        
        ' セルが空でなく、Insertで始まる場合を処理
        If Not IsEmpty(currentCell) Then
            If Left(currentCell.value, 6) = "Insert" Then
                ' 数値が続いているか確認（例：Insert1, Insert2など）
                If IsNumeric(Mid(currentCell.value, 7)) Then
                    currentValue = currentCell.value
                    
                    ' 前の値と異なる場合（新しいグループの開始）
                    If currentValue <> previousValue And previousValue <> "" Then
                        ' 数字部分を抽出して新しい値を作成
                        newValue = "NewColumn" & Mid(previousValue, 7)
                        
                        ' 現在の行の上に新しい行を挿入
                        ws.Rows(insertRow).Insert Shift:=xlDown
                        ' 挿入した行にNewColumn+Numを設定
                        ws.Cells(insertRow, "I").value = newValue
                        Debug.Print "Inserted row at " & insertRow & " with value " & newValue
                    End If
                    
                    ' 現在の値を記録
                    previousValue = currentValue
                    ' 次の挿入位置を現在の行に設定
                    insertRow = i
                End If
            End If
        End If
    Next i
    
    ' 最初のグループのための行挿入
    If insertRow > 0 Then
        ' 最後のグループの数字を使用して新しい値を作成
        newValue = "NewColumn" & Mid(previousValue, 7)
        
        ws.Rows(insertRow).Insert Shift:=xlDown
        ws.Cells(insertRow, "I").value = newValue
        Debug.Print "Inserted row at first group " & insertRow & " with value " & newValue
    End If
    
    MsgBox "ヘッダー行の挿入が完了しました。", vbInformation
End Sub

Public Sub FormatNewColumnRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCell As Range
    Dim aCell As Range
    Dim hasNewColumn As Boolean
    Dim excludeList As Variant
    
    ' 除外する文字列のリストを定義
    excludeList = Array("SampleText")
    
    ' "レポートグラフ"シートを設定
    Set ws = ThisWorkbook.Worksheets("レポートグラフ")
    
    ' 最終行を取得（I列とA列の大きい方を使用）
    lastRow = WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "I").End(xlUp).row, _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    
    ' I列に"NewColumn"が存在するかチェック
    hasNewColumn = False
    For i = 1 To lastRow
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                hasNewColumn = True
                Exit For
            End If
        End If
    Next i
    
    ' "NewColumn"が見つからない場合は処理を中止
    If Not hasNewColumn Then
        MsgBox "I列に'NewColumn'を含む値が見つかりません。" & vbCrLf & _
               "処理を中止します。", vbExclamation
        Exit Sub
    End If
    
    ' メイン処理
    For i = 1 To lastRow
        Set currentCell = ws.Cells(i, "I")
        Set aCell = ws.Cells(i, "A")
        
        ' A列の値を確認（3文字以上かつ除外リストに含まれない場合）
        If Not IsEmpty(aCell) Then
            If Len(aCell.value) >= 3 Then
                ' 除外リストに含まれていないかチェック
                Dim isExcluded As Boolean
                Dim excludeWord As Variant
                isExcluded = False
                
                For Each excludeWord In excludeList
                    If aCell.value = excludeWord Then
                        isExcluded = True
                        Exit For
                    End If
                Next excludeWord
                
                ' 除外リストに含まれていない場合のみ処理
                If Not isExcluded Then
                    ' 前の行のB-G列に値を転記
                    If i > 1 Then  ' 1行目より下の場合のみ
                        With ws.Range(ws.Cells(i - 1, "B"), ws.Cells(i - 1, "G"))
                            .Merge
                            .value = aCell.value
                            .HorizontalAlignment = xlLeft
                        End With
                    End If
                End If
            End If
        End If
        
        ' I列のNewColumnの処理
        If Not IsEmpty(currentCell) Then
            If Left(currentCell.value, 9) = "NewColumn" Then
                ' 行の高さを設定
                ws.Rows(i).RowHeight = 18
                
                ' B列からG列を結合
                With ws.Range(ws.Cells(i, "B"), ws.Cells(i, "G"))
                    .Merge
                    .HorizontalAlignment = xlLeft
                End With
                
                ' 背景色とフォント色を設定
                With ws.Range(ws.Cells(i, "A"), ws.Cells(i, "G"))
                    .Interior.Color = RGB(48, 84, 150)
                    .Font.Color = RGB(242, 242, 242)
                End With
                
                Debug.Print "Formatted NewColumn row at " & i
            End If
        End If
    Next i
    
    MsgBox "フォーマットが完了しました。", vbInformation
End Sub

Public Sub SetReportHeader()
    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim headerRange As Range
    Dim headerExists As Boolean
    
    ' シートの存在確認
    If Not WorksheetExists("レポートグラフ") Or Not WorksheetExists("レポート本文") Then
        MsgBox "必要なシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' シートの設定
    Set ws = ThisWorkbook.Worksheets("レポートグラフ")
    Set wsSource = ThisWorkbook.Worksheets("レポート本文")
    
    ' HeaderColumnの存在確認
    headerExists = False
    If Not IsEmpty(ws.Range("I1")) Then
        If ws.Range("I1").value = "HeaderColumn" Or ws.Range("I2").value = "HeaderColumn" Then
            headerExists = True
        End If
    End If
    
    ' ヘッダーが既に存在する場合は処理を終了
    If headerExists Then
        Debug.Print "ヘッダーは既に存在します。"
        Exit Sub
    End If
    
    ' 既存のヘッダー行を挿入
    ws.Rows("1:2").Insert Shift:=xlDown
    
    Application.ScreenUpdating = False
    
    ' A1:B2 の結合とコピー
    With ws.Range("A1:B2")
        .Merge
        .value = wsSource.Range("A1").value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    ' C1:D2 の結合とコピー
    With ws.Range("C1:E2")
        .Merge
        .value = wsSource.Range("C1").value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' F列の設定
    With ws.Range("F1:F2")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    ws.Range("F1").value = wsSource.Range("G1").value
    ws.Range("F2").value = wsSource.Range("G2").value
    
    ' G列の設定
    With ws.Range("G1:G2")
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("G1").value = wsSource.Range("H1").value
    ws.Range("G2").value = wsSource.Range("H2").value
    
    ' HeaderColumnの設定
    ws.Range("I1:I2").value = "HeaderColumn"
    
    ' 全体の書式設定
    With ws.Range("A1:G2")
        .Font.Name = "游ゴシック"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 行の高さを設定
    ws.Rows(1).RowHeight = 20
    ws.Rows(2).RowHeight = 20
    
    Application.ScreenUpdating = True
    
    Debug.Print "レポートヘッダーを設定しました。"
End Sub
Public Sub ClearColoredCellsInColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim clearedCount As Long
    
    ' シートの存在確認
    If Not WorksheetExists("レポートグラフ") Then
        MsgBox "レポートグラフシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' シートの設定
    Set ws = ThisWorkbook.Worksheets("レポートグラフ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Application.ScreenUpdating = False
    
    ' カウンターの初期化
    clearedCount = 0
    
    ' A列の各セルをチェック
    For Each cell In ws.Range("A1:A" & lastRow)
        ' セルに背景色がある場合
        If cell.Interior.colorIndex <> xlNone Then
            ' セルの内容をクリア
            cell.ClearContents
            clearedCount = clearedCount + 1
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    ' 結果を表示
    If clearedCount > 0 Then
        Debug.Print clearedCount & "個のセルの内容を消去しました。"
    Else
        Debug.Print "背景色のついているセルは見つかりませんでした。"
        MsgBox "背景色のついているセルは見つかりませんでした。", vbInformation
    End If
End Sub
' シートの存在確認関数
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    WorksheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function
