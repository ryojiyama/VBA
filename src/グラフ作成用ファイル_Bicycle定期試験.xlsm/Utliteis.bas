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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

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
    ' CopiedSheetNamesに記載されているシートのチャートを削除するプロシージャ
Public Sub DeleteChartsOnListedSheets()

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim processedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    Dim chartObj As ChartObject
    
    ' CopiedSheetNames シートを設定
    Set wsSource = ThisWorkbook.Sheets("CopiedSheetNames")
    Set processedSheets = New Collection ' 処理済みシート名を追跡するコレクション
    
    ' A列の最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' A列の値をループ
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value
        
        On Error Resume Next
        ' コレクションに同じ名前が既に存在するかチェック
        processedSheets.Add sheetName, sheetName
        
        If Err.Number = 0 Then ' 追加が成功した場合、シートはまだ処理されていない
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                ' シート上のすべてのグラフオブジェクトを削除
                For Each chartObj In wsTarget.ChartObjects
                    chartObj.Delete
                Next chartObj
            End If
        End If
        
        On Error GoTo 0 ' エラーハンドリングをリセット
        Set wsTarget = Nothing
    Next i
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
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    ' A列の値をループ
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value

        On Error Resume Next
        ' コレクションに同じ名前が既に存在するかチェック
        printedSheets.Add sheetName, sheetName
        If Err.Number = 0 Then ' 追加が成功した場合、シートはまだ印刷されていない
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' シートの1ページ目のみを印刷
            End If
        End If
        On Error GoTo 0 ' エラーハンドリングをリセット

        Set wsTarget = Nothing
    Next i
End Sub



' Settingの"B2"セルにフォーカス
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    If SheetExists("Setting") Then
        Application.GoTo ActiveWorkbook.Sheets("Setting").Range("B2")
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
