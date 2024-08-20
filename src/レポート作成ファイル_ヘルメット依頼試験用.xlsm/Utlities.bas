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

Sub PrintImpactSheet()
    Dim ws As Worksheet
    
    ' 条件1: 特定のシートを印刷
    Dim sheetNames1 As Variant
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    For Each ws In ThisWorkbook.Sheets
        If foundSheetName(ws.name, sheetNames1) Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    
    ' 条件2: "Impact_Side"を名前に含むシートを印刷
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.name, "Impact_Side") > 0 Then
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
        If InStr(ws.name, "Impact") > 0 Then
            ' ヘッダーの下の行から最終行までを削除
            ws.Rows("15:" & ws.Rows.Count).Delete
        End If
    Next ws
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
