Attribute VB_Name = "TranscribeInspectionReport"
Sub CopyFromExcelToWordBookmark()
    
    On Error GoTo ErrorHandler ' エラーハンドリング
    
    ' Excelのシートを設定
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' Wordアプリケーションとドキュメントを設定
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim filePath As String
    filePath = ThisWorkbook.Path & "PeriodicInspectionReport様品Ａ�Ｑ－０９－１４－０２　社内型式検定試験票_AutoTenki.docm"m"
    
    Set WordApp = New Word.Application
    
    ' Wordファイルが既に開いている場合、閉じる
    Dim docOpen As Boolean
    docOpen = False
    Dim doc As Word.Document
    For Each doc In WordApp.Documents
        If doc.fullName = filePath Then
            doc.Close SaveChanges:=wdSaveChanges
            docOpen = True
            Exit For
        End If
    Next doc
    
    ' Wordファイルを開く
    If docOpen Then
        Set WordDoc = WordApp.Documents.Open(filePath)
    Else
        Set WordDoc = WordApp.Documents.Open(filePath)
    End If
    
    ' ダイアログでIDを入力
    Dim id As String
    id = InputBox("Enter the ID to process", "ID Input")
    
    ' IDを基に行を検索
    Dim rng As Range
    Set rng = ws.Columns("B").Find(What:=id, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' IDが見つからない場合、処理を終了
    If rng Is Nothing Then
        MsgBox "ID not found"
        Exit Sub
    End If
    
    ' IDが見つかった行を取得
    Dim i As Long
    i = rng.row
    Dim productNumber As String
    productNumber = ws.Cells(i, "C").value
    
    With WordDoc
        ' ブックマークに値を転記
        If .Bookmarks.Exists("InspectionDate") Then .Bookmarks("InspectionDate").Range.text = ws.Cells(i, "F").value
        If .Bookmarks.Exists("ProductNumber") Then .Bookmarks("ProductNumber").Range.text = productNumber
        If .Bookmarks.Exists("Color") Then .Bookmarks("Color").Range.text = ws.Cells(i, "N").value
        If .Bookmarks.Exists("LotNumber") Then .Bookmarks("LotNumber").Range.text = ws.Cells(i, "O").value
        If .Bookmarks.Exists("TestContent") Then .Bookmarks("TestContent").Range.text = ws.Cells(i, "T").value
        If .Bookmarks.Exists("NaisouLot") Then .Bookmarks("NaisouLot").Range.text = ws.Cells(i, "Q").value
        If .Bookmarks.Exists("BoutaiLot") Then .Bookmarks("BoutaiLot").Range.text = ws.Cells(i, "P").value
        If .Bookmarks.Exists("Ondo") Then .Bookmarks("Ondo").Range.text = ws.Cells(i, "G").value
        If .Bookmarks.Exists("ResultA") Then .Bookmarks("ResultA").Range.text = ws.Cells(i, "R").value
        If .Bookmarks.Exists("ResultB") Then .Bookmarks("ResultB").Range.text = ws.Cells(i, "S").value
        If .Bookmarks.Exists("Pretreatment") Then .Bookmarks("Pretreatment").Range.text = ws.Cells(i, "K").value
        If .Bookmarks.Exists("Weight") Then .Bookmarks("Weight").Range.text = ws.Cells(i, "L").value
        If .Bookmarks.Exists("HeadClearance") Then .Bookmarks("HeadClearance").Range.text = ws.Cells(i, "M").value
        ' ドキュメントを保存して閉じる
        .SaveAs filePath & productNumber & .name
        .Close
    End With
    
    ' Wordアプリケーションを終了
    WordApp.Quit
    
    Exit Sub ' Clean-up とエラーハンドラの間に位置します。

ErrorHandler: ' エラーハンドラ
    MsgBox "An error has occurred: " & Err.Description
    ' オブジェクトを解放
    Set WordDoc = Nothing
    If Not WordApp Is Nothing Then WordApp.Quit
    Set WordApp = Nothing
    Set ws = Nothing
    Set rng = Nothing
End Sub



Sub ExportAllGraphsToWordAsPicture()

    Dim WordApp As Object
    Dim WordDoc As Object
    Dim ExcelWb As Workbook
    Dim ExcelWs As Worksheet
    Dim ExcelChart As ChartObject

    ' Wordアプリケーションを開始する
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    ' Wordアプリケーションを可視にする
    WordApp.Visible = True

    ' 新しいWordドキュメントを作成
    Set WordDoc = WordApp.Documents.Add

    ' Excelの指定されたワークブックとワークシートを開く
    Set ExcelWb = Workbooks.Open("グラフ作成用ファイル.xlsm")
    Set ExcelWs = ExcelWb.Sheets("LOG_Helmet")

    ' シート内のすべてのグラフをコピーしてWordにペースト
    For Each ExcelChart In ExcelWs.ChartObjects
        ' グラフの範囲を画像としてコピー
        ExcelChart.chart.CopyPicture Format:=xlPicture
    
        ' Wordのドキュメントの末尾にカーソルを移動
        Dim rng As Object
        Set rng = WordDoc.Content
        rng.Collapse Direction:=wdCollapseEnd  ' カーソルを末尾に移動
    
        ' グラフを画像としてペースト
        rng.Paste
        
        ' ペーストした画像の参照を取得
        Dim InlineShape As Object
        Set InlineShape = WordDoc.InlineShapes(WordDoc.InlineShapes.Count)
        
        ' 画像の大きさを調整
        InlineShape.LockAspectRatio = True   ' アスペクト比を保持
        InlineShape.Width = 200               ' ここでの値（200）は例としています。実際の値を指定してください。
        
        ' さらに、画像のレイアウトオプションを「前面」に設定
        InlineShape.ConvertToShape.WrapFormat.Type = wdWrapFront
    
        rng.InsertParagraphAfter
    Next ExcelChart

    ' すべてのオブジェクトをクリア
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set ExcelChart = Nothing
    Set ExcelWs = Nothing
    Set ExcelWb = Nothing

End Sub


Sub OpenWordTemplate()

    Dim WordApp As Object
    Dim WordDoc As Object
    Dim oneDrivePath As String
    Dim templatePath As String
    
    ' OneDriveのローカルパスを環境変数から取得
    oneDrivePath = Environ("OneDriveCommercial")

    ' OneDriveのパスと必要なサブフォルダ・ファイル名を組み合わせてテンプレートのパスを生成
    templatePath = oneDrivePath & "品質管理部の書類Ａ：保護帽依頼書２３－保護帽試験_テンプレート.docx"cx"

    ' Wordアプリケーションのオブジェクトを生成（Wordが既に開いている場合はそれを使用）
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    ' Wordを表示
    WordApp.Visible = True

    ' テンプレートファイルを開く
    Set WordDoc = WordApp.Documents.Open(templatePath)

End Sub
