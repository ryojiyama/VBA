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
    Dim endColumn As Integer: endColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column ' 最終列

    ' 列ごとにループ
    Dim i As Integer
    For i = startColumn To endColumn
        ' セルの値に1000を掛ける
        ws.Cells(1, i).value = ws.Cells(1, i).value * 1000
    Next i
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





