Attribute VB_Name = "OotputReport"
Option Explicit
Public Sub ShowOutputDialog()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("出力形式を選択してください" & vbNewLine & _
                     "はい：PDF出力" & vbNewLine & _
                     "いいえ：プリンタ出力", _
                     vbQuestion + vbYesNo, _
                     "出力形式の選択")
    
    If response = vbYes Then
        ExportReport "PDF"
    Else
        ExportReport "Print"
    End If
End Sub
' メイン処理のサブプロシージャ
Sub ExportReport(Optional outputType As String = "PDF")
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, pageNum As Long
    Dim currentHeight As Long
    Dim pageStartRow As Long
    Dim splitRow As Long
    Dim nextStartRow As Long
    Dim outputPath As String
    Dim workbookPath As String
    Dim hasNewColumn As Boolean
    
    ' "レポートグラフ"シートを設定
    If WorksheetExists("レポートグラフ") = False Then
        MsgBox "レポートグラフシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("レポートグラフ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    
    ' NewColumnの存在チェック
    hasNewColumn = CheckNewColumnExists(ws, lastRow)
    If Not hasNewColumn Then
        MsgBox "NewColumnが見つかりません。処理を中止します。", vbExclamation
        Exit Sub
    End If
    
    ' ワークブックのパスを取得と検証
    workbookPath = ValidateAndGetWorkbookPath()
    If workbookPath = "" Then Exit Sub
    
    ' 最初のNewColumn行を探す
    pageStartRow = FindFirstNewColumnRow(ws, lastRow)
    
    ' ページ番号と高さの初期化
    pageNum = 1
    currentHeight = 0
    
    ' 最初のページ設定
    SetupPageFormat ws
    
    ' 行を処理しながらページ分割
    For i = pageStartRow To lastRow
        currentHeight = currentHeight + ws.Rows(i).RowHeight
        
        ' 累積高さが728を超えた場合の処理
        If currentHeight >= 728 Then
            splitRow = FindSplitRow(ws, i, pageStartRow)
            nextStartRow = splitRow + 1
            
            ' 印刷範囲を設定
            ws.PageSetup.PrintArea = "$A$" & pageStartRow & ":$G$" & splitRow
            
            ' 出力処理
            Select Case outputType
                Case "PDF"
                    outputPath = workbookPath & "Report_" & Format(pageNum, "000") & ".pdf"
                    If Not ExportPageToPDF(ws, outputPath) Then Exit Sub
                Case "Print"
                    If Not ExportPageToPrinter(ws) Then Exit Sub
            End Select
            
            ' ページ設定をリセット
            ResetPageFormat ws
            
            ' 次のページの準備
            pageNum = pageNum + 1
            pageStartRow = nextStartRow
            i = nextStartRow - 1
            currentHeight = 0
            SetupPageFormat ws
            
        ElseIf i = lastRow Then
            ' 印刷範囲を設定
            ws.PageSetup.PrintArea = "$A$" & pageStartRow & ":$G$" & i
            
            ' 出力処理
            Select Case outputType
                Case "PDF"
                    outputPath = workbookPath & "Report_" & Format(pageNum, "000") & ".pdf"
                    If Not ExportPageToPDF(ws, outputPath) Then Exit Sub
                Case "Print"
                    If Not ExportPageToPrinter(ws) Then Exit Sub
            End Select
        End If
    Next i
    
    MsgBox "出力が完了しました。" & vbNewLine & _
           "保存先: " & workbookPath, vbInformation
    Exit Sub

ErrorHandler:
    HandleError Err.Number, Err.Description
End Sub

' PDF出力専用の関数
Private Function ExportPageToPDF(ws As Worksheet, pdfPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ファイルの使用状態をチェック
    If FileExists(pdfPath) Then
        If IsFileInUse(pdfPath) Then
            MsgBox "ファイル '" & pdfPath & "' が他のプロセスで使用中です。", vbExclamation
            ExportPageToPDF = False
            Exit Function
        End If
    End If
    
    ' PDFとして保存
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ExportPageToPDF = True
    Exit Function

ErrorHandler:
    HandleError Err.Number, Err.Description
    ExportPageToPDF = False
End Function

' プリンタ出力専用の関数
Private Function ExportPageToPrinter(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ' プリンタへ出力
    ws.PrintOut Copies:=1, Preview:=False, ActivePrinter:=Application.ActivePrinter
    
    ExportPageToPrinter = True
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 1004  ' プリンタ応答なし
            If InStr(1, Err.Description, "応答") > 0 Then
                MsgBox "プリンタが応答していません。" & vbNewLine & _
                       "プリンタの電源や接続を確認してください。", vbExclamation
            Else
                MsgBox "プリンタエラーが発生しました。" & vbNewLine & _
                       "プリンタの状態を確認してください。", vbExclamation
            End If
        Case Else
            MsgBox "プリンタエラーが発生しました。" & vbNewLine & _
                   "プリンタの状態を確認してください。", vbExclamation
    End Select
    ExportPageToPrinter = False
End Function

' ページ分割位置を探す関数
Private Function FindSplitRow(ws As Worksheet, currentRow As Long, startRow As Long) As Long
    Dim j As Long
    
    FindSplitRow = currentRow
    For j = currentRow To startRow Step -1
        If Left(ws.Cells(j, "I").value, 9) = "NewColumn" Then
            FindSplitRow = j - 1
            Exit Function
        End If
    Next j
End Function

' エラー処理関数
Private Sub HandleError(ErrorNum As Long, ErrorDesc As String)
    Select Case ErrorNum
        Case 1004  ' アプリケーションまたは権限エラー
            MsgBox "PDFファイルの作成権限がないか、または他のプロセスで使用中です。" & vbNewLine & _
                   "エラーの詳細: " & ErrorDesc, vbCritical
        Case 70, 75  ' ファイルアクセスエラー
            MsgBox "PDFファイルにアクセスできません。" & vbNewLine & _
                   "ファイルが他のプロセスで開かれているか、" & vbNewLine & _
                   "アクセス権限がない可能性があります。", vbCritical
        Case Else
            MsgBox "予期せぬエラーが発生しました。" & vbNewLine & _
                   "エラー番号: " & ErrorNum & vbNewLine & _
                   "エラーの説明: " & ErrorDesc, vbCritical
    End Select
End Sub

' NewColumnの存在チェック関数
Private Function CheckNewColumnExists(ws As Worksheet, lastRow As Long) As Boolean
    Dim i As Long
    
    CheckNewColumnExists = False
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                CheckNewColumnExists = True
                Exit Function
            End If
        End If
    Next i
End Function

' 最初のNewColumn行を探す関数
Private Function FindFirstNewColumnRow(ws As Worksheet, lastRow As Long) As Long
    Dim i As Long
    
    FindFirstNewColumnRow = 4  ' 開始行を4行目に変更
    For i = 4 To lastRow       ' 4行目から検索開始
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                FindFirstNewColumnRow = i
                Exit Function
            End If
        End If
    Next i
End Function

' ワークブックパスの検証関数
Private Function ValidateAndGetWorkbookPath() As String
    Dim workbookPath As String
    
    workbookPath = ThisWorkbook.Path
    If workbookPath = "" Then
        MsgBox "ワークブックが保存されていません。先に保存してください。", vbExclamation
        ValidateAndGetWorkbookPath = ""
        Exit Function
    End If
    
    If Right(workbookPath, 1) <> "\" Then
        workbookPath = workbookPath & "\"
    End If
    
    If Not FolderExists(workbookPath) Then
        MsgBox "保存先フォルダが見つかりません。", vbExclamation
        ValidateAndGetWorkbookPath = ""
        Exit Function
    End If
    
    ValidateAndGetWorkbookPath = workbookPath
End Function

' シートの存在確認関数
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    WorksheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function
' フォルダの存在確認関数
Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

' ファイルの存在確認関数
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(filePath) And vbDirectory) <> vbDirectory
    On Error GoTo 0
End Function
' ファイルの使用状態確認関数
Private Function IsFileInUse(ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    
    On Error Resume Next
    fileNum = FreeFile()
    Open filePath For Binary Access Read Write Lock Read Write As fileNum
    Close fileNum
    IsFileInUse = (Err.Number <> 0)
    On Error GoTo 0
End Function
' ページ設定を行うサブプロシージャ（ヘッダー行設定を追加）
Private Sub SetupPageFormat(ws As Worksheet)
    With ws.PageSetup
        ' 基本設定
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        
        ' ヘッダー行の設定（1-2行目を繰り返し表示）
        .PrintTitleRows = "$1:$2"
        
        ' 余白設定（必要に応じて調整可能）
        .HeaderMargin = Application.InchesToPoints(0.3)  ' ヘッダー余白
        .TopMargin = Application.InchesToPoints(0.75)    ' 上余白
        .BottomMargin = Application.InchesToPoints(0.75) ' 下余白
        .LeftMargin = Application.InchesToPoints(0.7)    ' 左余白
        .RightMargin = Application.InchesToPoints(0.7)   ' 右余白
    End With
End Sub

' ページ設定をリセットするサブプロシージャ
Private Sub ResetPageFormat(ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ""
        .PrintTitleRows = ""  ' ヘッダー行設定をクリア
        .Zoom = 100
        .FitToPagesWide = False
        .FitToPagesTall = False
    End With
End Sub
