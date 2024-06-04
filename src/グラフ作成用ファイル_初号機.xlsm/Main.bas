Attribute VB_Name = "Main"

Public Sub ShowForm()
    Call Form_Helmet.Show
End Sub



Sub MultiplyValues()
    ' 1行目をmsの値に変換する。(1000をかけるだけ)
    ' ワークシートを設定
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 開始列と終了列を設定
    Dim startColumn As Integer: startColumn = 22 ' V列
    Dim endColumn As Integer: endColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column ' 最終列

    ' 列ごとにループ
    Dim i As Integer
    For i = startColumn To endColumn
        ' セルの値に1000を掛ける
        ws.Cells(1, i).Value = ws.Cells(1, i).Value * 1000
    Next i
End Sub



' 各列に書式設定をする
Sub FormatCells()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range
    
    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")

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
            End If
        Next cell
    Next sheet
End Sub



Sub ClickIconAttheTop()
    ' 右のシートに移動する
    On Error Resume Next
    ActiveSheet.Next.Select
    If Err.number <> 0 Then
        MsgBox "This is the last sheet."
    End If
    On Error GoTo 0
End Sub

Sub ClickUSBIcon()
    'USBのアイコンをクリックする
    UserForm1.Show
End Sub


Sub ClickGraphIcon()
    'グラフのアイコンをクリックする
    UserForm1.Show
End Sub


Sub ClickPhotoIcon()
    '画像のアイコンをクリックする
    UserForm1.Show
End Sub


Sub ClicIconAttheBottom()
    ' 左のシートに移動する
    On Error Resume Next
    ActiveSheet.Previous.Select
    If Err.number <> 0 Then
        MsgBox "This is the first sheet."
    End If
    On Error GoTo 0
End Sub





