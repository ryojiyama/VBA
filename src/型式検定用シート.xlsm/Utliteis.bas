Attribute VB_Name = "Utliteis"


' 転記した15行目以下を削除する。
Sub ClearTransferData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim sheetNames As Variant
    Dim i As Long
    
    ' クリアするシート名をリスト化
    sheetNames = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    ' 各シートをループ処理
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        ' シートが存在する場合、データをクリア
        If Not ws Is Nothing Then
            startRow = 16 ' データの開始行（ヘッダーの次の行）
            lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
            
            ' データ範囲が存在する場合、クリア
            If lastRow >= startRow Then
                ws.Range("B" & startRow & ":Z" & lastRow).ClearContents
            End If
            
            ' wsをリセット
            Set ws = Nothing
        End If
    Next i
End Sub

