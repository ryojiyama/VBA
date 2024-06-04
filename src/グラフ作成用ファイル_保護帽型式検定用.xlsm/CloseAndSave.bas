Attribute VB_Name = "CloseAndSave"
Sub CloseAndSave()
    ' 確認のメッセージボックスを表示
    Dim Response As VbMsgBoxResult
    Response = MsgBox("読み込んだシートとすべてのグラフが消去されます。", vbOKCancel + vbQuestion, "確認")

    ' OKが押された場合に処理を実行
    If Response = vbOK Then
        Call DeleteAllChartsAndSheets
        Call SetRowHeightAndColumnWidth
        MsgBox "処理が完了しました。", vbInformation, "操作完了"
    End If
End Sub


Sub DeleteAllChartsAndSheets()
    ' シート中のグラフと余計なシートを削除
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String

    ' 保護するシートのリスト
    Dim protectSheets As Variant
    protectSheets = Array("LOG_Helmet", "Setting", "Hel_SpecSheet", "Penetration", "Impact_Top", "Impact_Front", "Impact_Back", "Impact_Side")

    ' 警告表示を無効化
    Application.DisplayAlerts = False

    ' 各シートに対して処理を実行
    For Each sheet In ThisWorkbook.Sheets
        sheetName = sheet.name
        ' 保護するシート以外の場合、シートを削除
        If Not IsInArray(sheetName, protectSheets) Then
            sheet.Delete
        End If
    Next sheet

    ' 警告表示を元に戻す
    Application.DisplayAlerts = True

    ' ブックを保存
    ThisWorkbook.Save
End Sub



' DeleteAllChartsAndSheets_配列内に特定の値が存在するかチェックする関数
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub SetRowHeightAndColumnWidth()
    ' A1の幅と高さを20にする。
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant

    ' 設定を適用するシート名のリストを定義する
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    ' シート名の配列をループする
    For Each sheetName In sheetNames
        ' シート名がこのワークブックに存在する場合、行の高さと列の幅を設定する
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Range("A1").RowHeight = 20
            ws.Range("A1").ColumnWidth = 20
            Set ws = Nothing
        End If
    Next sheetName
End Sub

