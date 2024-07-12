Attribute VB_Name = "Utliteis"
'CopiedSheetNamesで記されているシートを削除する。
Sub DeleteCopiedSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "CopiedSheetNamesシートが見つかりません。"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    Dim i As Long
    Application.DisplayAlerts = False
    For i = 1 To lastRow
        Dim sheetName As String
        sheetName = ws.Cells(i, 1).value
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    ClearCopiedSheetNames
End Sub
'CopiedSheetNamesをクリアする。
Sub ClearCopiedSheetNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Cells.ClearContents
    End If
End Sub
' "LOG_Helmet上のグラフを削除する
Public Sub DeleteAllChartsOnLOG_Helmet()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    ' "LOG_Helmet"シートを取得
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' シート上のすべてのグラフオブジェクトをループ
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Sub PrintMatchingSheetsFirstPage_SUb()
    Dim ws As Worksheet
    Dim copiedSheetNames As Worksheet
    Dim sheetName As String
    Dim cell As Range
    Dim foundSheet As Worksheet
    
    ' CopiedSheetNamesシートを設定
    Set copiedSheetNames = ThisWorkbook.Sheets("CopiedSheetNames")
    
    ' A列の値をループ
    For Each cell In copiedSheetNames.Range("A1:A" & copiedSheetNames.Cells(copiedSheetNames.Rows.count, "A").End(xlUp).row)
        sheetName = cell.value
        
        ' 一致するシートを検索
        On Error Resume Next
        Set foundSheet = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' シートが存在する場合、1ページ目を印刷
        If Not foundSheet Is Nothing Then
            With foundSheet
                ' 印刷領域を設定
                .PageSetup.PrintArea = ""
                ' シートを1ページ目のみ印刷
                .PrintOut Preview:=False
            End With
            ' foundSheetをクリア
            Set foundSheet = Nothing
        End If
    Next cell
End Sub
Sub PrintFirstPageOfUniqueListedSheets()
    ' 指定された検査票の1ページ目を、重複なく1回ずつ印刷するプロシージャ
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim printedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' CopiedSheetNames シートを設定
    Set wsSource = ThisWorkbook.Sheets("CopiedSheetNames")
    Set printedSheets = New Collection ' 印刷されたシート名を追跡するコレクション

    ' A列の最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row

    ' A列の値をループ
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value

        On Error Resume Next
        ' コレクションに同じ名前が既に存在するかチェック
        printedSheets.Add sheetName, sheetName
        If Err.number = 0 Then ' 追加が成功した場合、シートはまだ印刷されていない
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' シートの1ページ目のみを印刷
            End If
        End If
        On Error GoTo 0 ' エラーハンドリングをリセット

        Set wsTarget = Nothing
    Next i
End Sub




'Sub DeleteAllChartsAndSheets()
'    ' シート中のグラフと余計なシートを削除
'    Dim sheet As Worksheet
'    Dim chart As ChartObject
'    Dim sheetName As String
'    Dim proceed As Integer
'
'    ' シートのリスト
'    Dim sheetList() As Variant
'    sheetList = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
'
'    Application.DisplayAlerts = False
'
'    ' 各シートに対して処理を実行
'    For Each sheet In ThisWorkbook.Sheets
'        sheetName = sheet.name
'        ' グラフの削除とデータの警告表示
'        If IsInArray(sheetName, sheetList) Then
'            For Each chart In sheet.ChartObjects
'                chart.Delete
'            Next chart
'            ' B2セルからZZ15までのデータの有無をチェックし、有れば警告を表示
'            If Application.WorksheetFunction.CountA(sheet.Range("B2:ZZ15")) <> 0 Then
'                Application.DisplayAlerts = True
'                proceed = MsgBox("Sheet '" & sheetName & "' contains data. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
'                Application.DisplayAlerts = False
'                If proceed = vbNo Then Exit Sub
'            End If
'        ' シートの削除
'        ElseIf sheetName <> "Setting" And sheetName <> "Hel_SpecSheet" And sheetName <> "InspectionSheet" Then
'            sheet.Delete
'        End If
'    Next sheet
'
'    Application.DisplayAlerts = True
'
'    ' ブックを保存
'    ThisWorkbook.Save
'
'End Sub
