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

' 右クリックカスタムメニュー：グラフのY軸の値調整
Sub UniformizeLineGraphAxes()
    On Error GoTo ErrorHandler
    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                ' Split the chart name using "-"
                Dim Parts() As String
                Parts = Split(chartObj.name, "-")
                
                ' Check the third part of the name
                If UBound(Parts) >= 2 Then
                    With chartObj.chart.Axes(xlValue)
                        If Parts(2) = "天" Then
                            .MaximumScale = 5
                            .MajorUnit = 1# ' 1.0刻み
                        ElseIf Parts(2) = "前" Or Parts(2) = "後" Or Parts(2) = "側面" Then
                            .MaximumScale = 10
                            .MajorUnit = 2# ' 2.0刻み
                        End If
                    End With
                End If
            Next chartObj
        End If
    Next ws

    MsgBox "すべてのシートのグラフのY軸の最大値と目盛り単位を設定しました。", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical

End Sub

' LOG_Helmetシートのアイコンを消す。
Sub DeleteIconsKeepCharts()
    Dim ws As Worksheet
    Dim shp As Shape

    ' LOG_Helmetシートを指定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' シート内のすべてのシェイプをループ処理
    For Each shp In ws.Shapes
        ' シェイプがグラフオブジェクトでない場合、削除
        If shp.Type <> msoChart Then
            shp.Delete
        End If
    Next shp
End Sub

