Attribute VB_Name = "Utlities"

Option Explicit

'*******************************************************************************
' 定数定義
'*******************************************************************************
Private Const FIRST_DATA_ROW As Long = 2
Private Const FIRST_DATA_COL As String = "B"

'ヘッダー情報の定義
Private Type SheetHeaders
    sheetName As String
    headers() As String
End Type

'*******************************************************************************
' メインプロシージャ
' 機能：指定されたLOGシートのデータをクリアしヘッダーを設定
' 引数：なし
'*******************************************************************************
Public Sub ResetLogSheets()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim targetSheets As Variant
    targetSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    
    Dim i As Long
    Dim sheetName As String
    Dim missingSheets As String
    Dim processedSheets As String
    Dim processedCount As Long
    Dim ws As Worksheet
    
    processedCount = 0
    
    ' 各シートを処理
    For i = LBound(targetSheets) To UBound(targetSheets)
        sheetName = CStr(targetSheets(i))
        
        ' シートの存在確認と取得を一度に行う
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo ErrorHandler
        
        If Not ws Is Nothing Then
            ' シートが存在する場合は処理を実行
            ClearSheetData ws
            SetSheetHeaders ws
            
            ' 処理済みシートを記録
            If processedSheets = "" Then
                processedSheets = sheetName
            Else
                processedSheets = processedSheets & ", " & sheetName
            End If
            processedCount = processedCount + 1
        Else
            ' 存在しないシートを記録
            If missingSheets = "" Then
                missingSheets = sheetName
            Else
                missingSheets = missingSheets & ", " & sheetName
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True

    ' 結果報告
    If processedCount = 0 Then
        Debug.Print "以下のシートが見つかりません: " & missingSheets
    ElseIf missingSheets <> "" Then
        Debug.Print "処理完了: " & processedSheets
        Debug.Print "未処理（シート無し）: " & missingSheets
    Else
        Debug.Print "全シート処理完了"
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "エラー発生: " & Err.Description & " (エラー番号: " & Err.Number & ")"
End Sub

' シートの存在をチェックする関数
Private Function sheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    sheetExists = Not ws Is Nothing
End Function

'*******************************************************************************
' サブプロシージャ
' 機能：指定シートのデータをクリア（B列以降のデータのみ）
' 引数：ws - クリア対象のシート名
' 前提：FIRST_DATA_ROW は定数として定義済み
'*******************************************************************************
Private Sub ClearSheetData(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Debug.Print "処理開始 - シート名: " & ws.Name
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim clearRange As Range
    
    With ws
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        lastCol = .Cells(1, .columns.Count).End(xlToLeft).column
        
        If lastRow >= FIRST_DATA_ROW And lastCol >= 2 Then
            Set clearRange = .Range(.Cells(FIRST_DATA_ROW, "B"), .Cells(lastRow, lastCol))
            With clearRange
                .ClearContents
                .Interior.colorIndex = xlNone
                .Borders.LineStyle = xlNone
            End With
            Debug.Print "データクリア完了 - 行数: " & (lastRow - FIRST_DATA_ROW + 1)
        Else
            Debug.Print "クリア対象データなし"
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "エラー発生 - シート '" & ws.Name & "': " & Err.Description
End Sub

'*******************************************************************************
' サブプロシージャ
' 機能：シートのヘッダーを設定
' 引数：ws - ヘッダーを設定するシート名
' 依存：GetSheetHeaders関数によるヘッダー情報の取得
'*******************************************************************************
Private Sub SetSheetHeaders(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' ヘッダー情報の取得
    Dim headers As Variant
    headers = GetSheetHeaders(ws.Name)
    
    If Not IsEmpty(headers) Then
        Debug.Print "ヘッダー設定開始 - シート名: " & ws.Name
        
        ' A列からZ列までのヘッダーをクリア
        ws.Range("A1:Z1").ClearContents
        
        ' 各列にヘッダーを設定
        Dim headerRange As Range
        Set headerRange = ws.Range("A1:Z1")
        
        ' A列は固定で "順番"
        ws.Range("A1").value = ""
        
        ' B列以降に配列の内容を設定
        Dim i As Long
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 2).value = headers(i)
        Next i
        
        ' ヘッダー行の書式設定
        With headerRange
            .Font.Name = "游ゴシック"
            .Font.size = 12
            .Font.Bold = True
            .Font.Color = RGB(217, 217, 217)
            .Interior.Color = RGB(48, 84, 150)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(190, 190, 190)
            End With
            
        ' 行の高さを設定
        headerRange.EntireRow.RowHeight = 20
        ' 列幅を設定（新しいサブルーチンを呼び出し）
        SetColumnWidths ws

        End With
    Else
        Debug.Print "警告 - ヘッダー情報が空です: " & ws.Name
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "エラー発生 - シート '" & ws.Name & "': " & Err.Description
    Debug.Print "エラー番号: " & Err.Number
End Sub

'*******************************************************************************
' 関数
' 機能：シート別のヘッダー情報を取得
' 引数：ws - ヘッダーを取得するシート名
' 戻値：ヘッダー配列、未定義の場合はEmpty
'*******************************************************************************
Private Function GetSheetHeaders(ByVal sheetName As String) As Variant
    
    Select Case Trim(sheetName)
        Case "LOG_Helmet"
            GetSheetHeaders = Array("ID", "試料ID", "品番", "試験内容", _
                "検査日", "温度", "最大値(kN)", "最大値の時間", "4.9(ms)", "7.3(ms)", _
                "前処理", "重量", "天頂すきま", "帽体色", "ロットNo.", _
                "帽体ロット", "内装ロット", "構造_検査結果", "耐貫通_検査結果", "試験区分")
        
        Case "LOG_FallArrest"
            GetSheetHeaders = Empty
            
        Case "LOG_Bicycle"
            GetSheetHeaders = Array("ID", "試料ID", "品番", "ロット番号", _
                "試験日", "温度", "湿度", "重量", "最大値(G)", "最大値の時間", "Gの継続時間", _
                "前処理", "試験箇所", "帽体色", "帽体の材質", "アンビル", "人頭模型", "試験後の状態", _
                "外観検査", "あごひも検査", "材料・付属品検査")
            
        Case "LOG_BaseBall"
            GetSheetHeaders = Empty
            
        Case Else
            GetSheetHeaders = Empty
    End Select
End Function
'*******************************************************************************
' サブプロシージャ
' 機能：シート別の列幅を設定
' 引数：ws - 列幅を設定するワークシート
'*******************************************************************************
Private Sub SetColumnWidths(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Debug.Print "列幅設定開始 - シート名: " & ws.Name
    
    With ws
        ' A列（順番）は共通
        .columns("A").ColumnWidth = 6
        
        Select Case ws.Name
            Case "LOG_Helmet"
                ' ID列とその他の列で幅を変える
                .columns("B").ColumnWidth = 20  ' ID
                .columns("C").ColumnWidth = 12  ' 試料ID
                .Range("D:U").ColumnWidth = 15  ' その他の列
                
            Case "LOG_Bicycle"
                .columns("B").ColumnWidth = 20  ' ID
                .columns("C").ColumnWidth = 12  ' 試料ID
                .Range("D:V").ColumnWidth = 10  ' その他の列
                
                ' 特定の列は幅を広げる
                .columns("S").ColumnWidth = 20  ' 試験後の状態
                .columns("T").ColumnWidth = 20  ' 概観検査
                .columns("U").ColumnWidth = 20  ' あごひも検査
                .columns("V").ColumnWidth = 20  ' 材料・付属品検査
        End Select
    End With
    
    Debug.Print "列幅設定完了 - シート名: " & ws.Name
    Exit Sub
    
ErrorHandler:
    Debug.Print "エラー発生 - 列幅設定中: " & ws.Name
    Debug.Print "  - エラー内容: " & Err.Description
End Sub
'*******************************************************************************
' サブプロシージャ
' 機能：新しいヘッダーを追加するための補助プロシージャ
' 引数：sheetName - ヘッダーを追加するシート名
'       headers - 追加するヘッダー配列
'*******************************************************************************
Public Sub AddNewHeaders(ByVal sheetName As String, ByRef headers As Variant)
    ' この関数は将来的にヘッダーを追加する際に使用
    ' 実装例：
    ' Dim newHeaders As Variant
    ' newHeaders = Array("新しいヘッダー1", "新しいヘッダー2", ...)
    ' AddNewHeaders "LOG_FallArrest", newHeaders
End Sub



' DeleteAllChartsAndSheets_シート中のグラフを削除する
Sub DeleteAllChartsAndSheets()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String

    ' シートのリスト
    Dim sheetList() As Variant
    sheetList = Array( _
        "Setting", _
        "LOG_Helmet", _
        "LOG_BaseBall", _
        "LOG_Bicycle", _
        "LOG_FallArrest", _
        "レポート本文", _
        "レポートグラフ", _
        "試験結果" _
    )
    Application.DisplayAlerts = False

    ' 各シートに対して処理を実行
    For Each sheet In ActiveWorkbook.Sheets
        sheetName = sheet.Name
        ' グラフの削除
        If IsInArray(sheetName, sheetList) Then
            For Each chart In sheet.ChartObjects
                chart.Delete
            Next chart
        End If
    Next sheet

    Application.DisplayAlerts = True
End Sub

' DeleteAllChartsAndSheets_配列内に特定の値が存在するかチェックする関数
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function





' 各列に書式設定をする
Public Sub CustomizeSheetFormats()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "ID") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試料ID") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "品番") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試験内容") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "検査日") > 0 Then ' Date
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToDate(rng)
            ElseIf InStr(1, cell.value, "温度") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "最大値(kN)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericFourDecimals(rng)
            ElseIf InStr(1, cell.value, "最大値の時間(ms)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "4.9kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "7.3kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "前処理") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "重量") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "天頂すきま") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "製品ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "帽体ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "内装ロット") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "構造検査") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "耐貫通検査") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "試験区分") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            End If
        Next cell
    Next sheet
End Sub

Sub ConvertToNumeric(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericTwoDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.00"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericFourDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0000"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToString(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "@"
    For Each cell In rng
        cell.value = CStr(cell.value)
    Next cell
End Sub

Sub ConvertToDate(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "yyyy/mm/dd"  ' 日付の表示形式を設定
    For Each cell In rng
        If IsDate(cell.value) Then
            cell.value = CDate(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub
' 空白セルに"-"を挿入
Public Sub FillBlanksWithHyphenInMultipleSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim sheetName As Variant

    ' 対象シートの名前を配列に設定
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' 各シートについて処理を行う
    For Each sheetName In sheetNames
        On Error Resume Next
        ' 対象シートを設定
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = Nothing ' ws変数をクリア
            GoTo NextSheet ' 次のシートに進む
        End If

        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        lastCol = ws.Cells(1, "Z").column ' Z列の列番号を設定

        ' 2行目から最終行までループ（1行目はヘッダーと仮定）
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    'Debug.Print "EmptyCell:" & "Cells&("; i; "," & j; ")"
                    ws.Cells(i, j).value = "-"
                End If
            Next j
        Next i

        ' シート処理の終了ラベル
NextSheet:
        ' 次のシートの処理に移る前に変数をクリア
        Set ws = Nothing
    Next sheetName
End Sub

