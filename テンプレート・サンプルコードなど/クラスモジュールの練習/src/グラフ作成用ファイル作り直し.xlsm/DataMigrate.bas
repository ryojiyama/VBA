Attribute VB_Name = "DataMigrate"
'Databaseフォルダの"試験結果_データベース.xlsm"に試験データを転記する。
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
    Set sourceWorkbook = ThisWorkbook
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

Function OpenWorkbook(ByVal fullPath As String) As Workbook
    'DataMigration_GraphToTestDB_FromGraphbook()のサブプロシージャ
    Dim wb As Workbook
    ' Debug.Print "fullPath: " & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

Sub MigrateData(ByRef sourceWorkbook As Workbook, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()のサブプロシージャ
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    ' 元のシート名、ターゲットシート名、IDプレフィックスを配列として設定
    sourceSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    targetSheets = Array("HeLmetTestData", "FallArrestTestData", "biCycleHelmetTestData", "BaseBallTestData")
    IDPrefixes = Array("HBT-", "FAT-", "CHT-", "BBT-")

    ' 配列の各要素に対してデータの転記を行う
    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWorkbook.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWorkbook.Sheets(targetSheets(i))

        ' データのコピーを実行
        CopyData_CopyPaste sourceSheet, targetSheet, IDPrefixes(i), targetWorkbook
    Next i
End Sub

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
        MsgBox "転記元のシート " & sourceSheet.Name & " にデータがありません。", vbExclamation
        Exit Sub
    End If
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).row

    ' 転記元のレコード数を計算
    numRecords = lastRow - 1 ' ヘッダー行を除外
    currentID = targetSheet.Cells(targetLastRow, "B").value

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

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True
End Sub

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
