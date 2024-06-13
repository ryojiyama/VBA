Attribute VB_Name = "Utlities"
' DeleteAllChartsAndSheets_シート中のグラフと余計なシートを削除する
Sub DeleteAllChartsAndSheets()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String
    Dim proceed As Integer

    ' シートのリスト
    Dim sheetList() As Variant
    sheetList = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")

    Application.DisplayAlerts = False

    ' 各シートに対して処理を実行
    For Each sheet In ThisWorkbook.Sheets
        sheetName = sheet.name
        ' グラフの削除とデータの警告表示
        If IsInArray(sheetName, sheetList) Then
            For Each chart In sheet.ChartObjects
                chart.Delete
            Next chart
            ' B2セルからZZ15までのデータの有無をチェックし、有れば警告を表示
            If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
                Application.DisplayAlerts = True
                proceed = MsgBox("Sheet '" & sheetName & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
                Application.DisplayAlerts = False
                If proceed = vbNo Then Exit Sub
            End If
        ' シートの削除
        ElseIf sheetName <> "Setting" And sheetName <> "Hel_SpecSheet" Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True


End Sub

' DeleteAllChartsAndSheets_配列内に特定の値が存在するかチェックする関数
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


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
