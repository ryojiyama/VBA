Attribute VB_Name = "DataMigrate"
'*******************************************************************************
' メインプロシージャ
' 機能：Databaseフォルダの試験結果データベースファイルにデータを転記
' 引数：なし
'*******************************************************************************
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim relativePath As String
    Dim localPath As String
    Dim wb As Workbook
    Dim isOpen As Boolean

    '"OneDriveGraph:C:\Users\QC07\TSホールディングス株式会社\OfficeScriptの整理 - ドキュメント\QC_グラフ作成"
    localPath = Environ("OneDriveGraph") ' & "\Database\Database試験結果_データベース.xlsm"

    ' 現在のディレクトリを基準に相対パスを設定
    relativePath = localPath & "\Database\試験結果_データベース.xlsm"
    Set sourceWorkbook = ActiveWorkbook
    Set targetWorkbook = Workbooks.Open(relativePath)


    ' 試験結果_データベース.xlsmが既に開かれているかを確認
    isOpen = False
    For Each wb In Application.Workbooks
        If wb.FullName = relativePath Then
            Set targetWorkbook = wb
            isOpen = True
            Exit For
        End If
    Next wb

    ' 開かれていない場合はOpenWorkbook関数を使用して開く
    If Not isOpen Then
        Set targetWorkbook = OpenWorkbook(relativePath)
    End If

    On Error GoTo ErrorHandler

    ' データの転記処理を実行
    MigrateData sourceWorkbook, targetWorkbook
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True
End Sub

'*******************************************************************************
' サブプロシージャ
' 機能：指定されたパスのワークブックを開く
' 引数：fullPath - 開くワークブックの完全パス
' 戻値：開いたWorkbookオブジェクト
'*******************************************************************************
Function OpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    ' Debug.Print "fullPath: " & fullPath
    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

'*******************************************************************************
' サブプロシージャ
' 機能：ソースワークブックからターゲットワークブックへデータを転記
' 引数：sourceWorkbook - 転記元のワークブック
'       targetWorkbook - 転記先のワークブック
'*******************************************************************************
Sub MigrateData(ByRef sourceWorkbook As Workbook, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()のサブプロシージャ
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer
    Dim sheetExists As Boolean
    
    ' 元のシート名、ターゲットシート名、IDプレフィックスを配列として設定
    sourceSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    targetSheets = Array("HeLmetTestData", "FallArrestTestData", "biCycleHelmetTestData", "BaseBallTestData")
    IDPrefixes = Array("HBT-", "FAT-", "CHT-", "BBT-")
    
    Debug.Print "データ転記処理開始: " & Now
    
    ' 配列の各要素に対してデータの転記を行う
    For i = LBound(sourceSheets) To UBound(sourceSheets)
        ' ソースシートの存在チェック
        sheetExists = False
        On Error Resume Next
        sheetExists = Not sourceWorkbook.Sheets(sourceSheets(i)) Is Nothing
        On Error GoTo 0
        
        If Not sheetExists Then
            Debug.Print "警告: ソースシート '" & sourceSheets(i) & "' が見つかりません。スキップします。"
            GoTo NextIteration
        End If
        
        ' ターゲットシートの存在チェック
        sheetExists = False
        On Error Resume Next
        sheetExists = Not targetWorkbook.Sheets(targetSheets(i)) Is Nothing
        On Error GoTo 0
        
        If Not sheetExists Then
            Debug.Print "警告: ターゲットシート '" & targetSheets(i) & "' が見つかりません。スキップします。"
            GoTo NextIteration
        End If
        
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWorkbook.Sheets(sourceSheets(i))
        
        Dim targetSheet As Worksheet
        Set targetSheet = targetWorkbook.Sheets(targetSheets(i))
        
        Debug.Print "処理中: " & sourceSheets(i) & " → " & targetSheets(i)
        
        ' データのコピーを実行
        CopyData_CopyPaste sourceSheet, targetSheet, IDPrefixes(i), targetWorkbook
        
NextIteration:
    Next i
    
    Debug.Print "データ転記処理完了: " & Now
End Sub
'*******************************************************************************
' サブプロシージャ
' 機能：指定されたシート間でデータをコピー＆ペースト
' 引数：sourceSheet - 転記元のワークシート
'       targetSheet - 転記先のワークシート
'       IDPrefix - 新規ID生成時のプレフィックス
'       targetWorkbook - 転記先のワークブック
'*******************************************************************************
Sub CopyData_CopyPaste(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()のサブプロシージャ
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim currentID As String
    Dim newIDCollection As Collection
    Dim numRecords As Long
    Dim i As Long

    ' 転記元のシートの最終行と最終列を取得（ヘッダー行を除外）
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        Debug.Print "警告: シート '" & sourceSheet.Name & "' にデータがありません。"
        GoTo ExitHandler
    End If
    lastColumn = sourceSheet.Cells(1, sourceSheet.columns.Count).End(xlToLeft).column
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).row

    ' 転記元のレコード数を計算
    numRecords = lastRow - 1 ' ヘッダー行を除外
    currentID = targetSheet.Cells(targetLastRow, "B").value
    
    Debug.Print "処理開始: シート '" & sourceSheet.Name & "'"
    Debug.Print "  - 転記元レコード数: " & numRecords
    Debug.Print "  - 転記開始位置: " & (targetLastRow + 1)
    Debug.Print "  - 最終ID: " & currentID

    ' 新しいIDを生成
    Set newIDCollection = GetNewID(currentID, IDPrefix, numRecords)

    ' 転記元のデータ範囲をコピー（ヘッダー行を除外）
    sourceSheet.Range(sourceSheet.Cells(2, 1), sourceSheet.Cells(lastRow, lastColumn)).Copy
    ' 新しいデータをペーストする場所
    targetSheet.Cells(targetLastRow + 1, 1).PasteSpecial Paste:=xlPasteValues

    ' 新しいIDを追加
    For i = 1 To numRecords
        targetSheet.Cells(targetLastRow + i, "B").value = newIDCollection(i)
    Next i

    Debug.Print "処理完了: シート '" & sourceSheet.Name & "'"
    Debug.Print "  - 新規ID範囲: " & newIDCollection(1) & " 〜 " & newIDCollection(numRecords)

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "エラー発生 - シート '" & sourceSheet.Name & "'"
    Debug.Print "  - エラー内容: " & Err.Description
    Debug.Print "  - エラーコード: " & Err.Number
    Resume ExitHandler
End Sub
'*******************************************************************************
' サブプロシージャ
' 機能：指定されたプレフィックスで連番の新規IDを生成
' 引数：currentID - 現在の最新ID
'       IDPrefix - IDのプレフィックス
'       numRecords - 生成するID数
' 戻値：生成されたIDのCollection
'*******************************************************************************
Function GetNewID(ByVal currentID As String, ByVal IDPrefix As String, ByVal numRecords As Long) As Collection
    'DataMigration_GraphToTestDB_FromGraphbook()のサブプロシージャ
    Dim newIDCollection As Collection
    Set newIDCollection = New Collection

    Dim currentNumber As Long
    Dim i As Long
    Dim idNumberPart As String

    ' プレフィックスを取り除いて数値部分を抽出
    idNumberPart = Replace(currentID, IDPrefix, "")
    currentNumber = Val(idNumberPart)

    ' 複数の新しいIDを生成
    For i = 1 To numRecords
        currentNumber = currentNumber + 1
        newIDCollection.Add IDPrefix & Format(currentNumber, "00000")
    Next i

    Set GetNewID = newIDCollection
    ' Debug.Print "Generated " & numRecords & " new IDs starting from " & currentNumber - numRecords + 1
End Function
'*******************************************************************************
' メインプロシージャ
' 機能：LOGシートを指定された製品カテゴリのフォルダ内のワークブックにコピー
' 引数：selectedButton - コピー先を指定するボタン名
'*******************************************************************************
Sub CopySheetsToOtherWorkbooks(selectedButton As String)
    Dim sheetNames As Variant
    Dim folderNames As Variant
    Dim sheetName As Variant
    Dim folderName As Variant
    Dim ws As Worksheet
    Dim destWb As Workbook
    Dim destFile As String
    Dim destDir As String
    Dim file As String
    Dim fileCount As Integer
    Dim copySheetName As String
    Dim oneDrivePath As String
    
    Application.ScreenUpdating = False

    ' 環境変数からOneDriveのパスを取得
    oneDrivePath = Environ("OneDriveGraph")
    
    ' 対象シート名とフォルダ名のリスト
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    folderNames = Array("☆Helmet", "☆FallArrest", "☆Bicycle", "☆BaseBall")

    ' シートごとに処理
    For i = LBound(sheetNames) To UBound(sheetNames)
        sheetName = sheetNames(i)
        folderName = folderNames(i)
        
        ' 対象シートのオブジェクトを設定
        Set ws = ActiveWorkbook.Sheets(sheetName)

        ' B2セルが空かどうか確認
        If ws.Range("B2").value <> "" Then
            ' コピー先ディレクトリを設定
            destDir = oneDrivePath & "\" & folderName & "\"
            Debug.Print "DestDir:" & destDir
            
            ' コピー先ファイルをループで開く
            file = Dir(destDir & "*.xls*")
            Do While file <> ""
                ' selectedButtonの内容に基づいてフィルタリング
                If InStr(file, selectedButton) > 0 Then
                    destFile = destDir & file
                    Set destWb = Workbooks.Open(destFile)
                    
                    ' 連番をつけてコピー
                    fileCount = 1
                    copySheetName = sheetName & "-" & fileCount
                    Do While sheetExists(copySheetName, destWb)
                        fileCount = fileCount + 1
                        copySheetName = sheetName & "-" & fileCount
                    Loop
                    
                    ' シートをコピー
                    ws.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    destWb.Sheets(destWb.Sheets.Count).Name = copySheetName
                    destWb.Close SaveChanges:=True
                End If
                
                ' 次のファイルへ
                file = Dir
            Loop
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

'*******************************************************************************
' サブプロシージャ
' 機能：指定されたワークブック内にシートが存在するかを確認
' 引数：sheetName - 確認するシート名
'       wb - 確認対象のワークブック
' 戻値：シートが存在する場合はTrue
'*******************************************************************************
Function sheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    sheetExists = Not ws Is Nothing
End Function

