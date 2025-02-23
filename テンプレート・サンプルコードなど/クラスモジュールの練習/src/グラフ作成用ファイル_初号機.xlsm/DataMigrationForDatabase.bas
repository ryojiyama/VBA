Attribute VB_Name = "DataMigrationForDatabase"
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDriveのパスを取得
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_試験グラフ作成" & "\" & "試験結果_データベース.xlsm"

    ' sourceWorkbookを開く
    Set sourceWorkbook = OpenWorkbook("C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\QC_試験グラフ作成\", "グラフ作成用ファイル.xlsm")

    ' myPathを使用してtargetWorkbookを開く
    Set targetWorkbook = OpenWorkbook(myPath, "")

    ' 転記処理
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Sub MigrateData(ByRef sourceWB As Workbook, ByRef targetWB As Workbook)
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    sourceSheets = Array("LOG_Helmet")
    targetSheets = Array("HeLmetTestData", "BaseBallTestData", "biCycleHelmetTestData", "FallArrestTestData")
    IDPrefixes = Array("HBT-", "BBT-", "CHT-", "FAT-")

    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWB.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWB.Sheets(targetSheets(i))

        CopyData sourceSheet, targetSheet, IDPrefixes(i)
    Next i
End Sub

Sub CopyData_CopyPaste(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim IDGenRow As Long
    Dim currentID As String

    ' 転記元のシートの最終行と最終列を取得
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).column

    ' 転記先のシートの最終行を取得
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row + 1

    For IDGenRow = 2 To lastRow
        ' 新しいIDを生成して転記先のシートのC列にセット
        currentID = GetNewID(targetSheet, IDPrefix)
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' 転記元から転記先へデータをコピー
        sourceSheet.Range(sourceSheet.Cells(IDGenRow, "C"), sourceSheet.Cells(IDGenRow, "U")).Copy _
            Destination:=targetSheet.Cells(targetLastRow, "D")

        ' D列から最終列までをD列から最終列へコピー
        If lastColumn > 4 Then ' 4列より多い場合のみ実行
            sourceSheet.Range(sourceSheet.Cells(IDGenRow, "D"), sourceSheet.Cells(IDGenRow, lastColumn)).Copy _
                Destination:=targetSheet.Cells(targetLastRow, "E")
        End If

        targetLastRow = targetLastRow + 1
    Next IDGenRow

    ' 転記した範囲を削除（データのみを削除）
    sourceSheet.Range(sourceSheet.Cells(2, "B"), sourceSheet.Cells(lastRow, lastColumn)).ClearContents

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description & " in " & sourceSheet.name, vbCritical
    Application.ScreenUpdating = True
End Sub
Sub CopyData(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim dataRange As Range
    Dim data As Variant
    Dim targetLastRow As Long
    Dim i As Long
    Dim currentID As String

    ' 転記元のシートの最終行を取得
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row

    ' 転記するデータ範囲を設定（ここで列範囲を適宜調整してください）
    Set dataRange = sourceSheet.Range("C2:U" & lastRow) ' 例: C列からU列まで

    ' データ範囲を配列に読み込む
    data = dataRange.Value

    ' 転記先のシートの最終行を取得
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row + 1

    ' 配列のデータを転記先に転記
    For i = LBound(data, 1) To UBound(data, 1)
        ' 新しいIDを生成
        currentID = GetNewID(targetSheet, IDPrefix)

        ' IDをセット
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' 配列からデータを転記
        Dim j As Long
        For j = LBound(data, 2) To UBound(data, 2)
            targetSheet.Cells(targetLastRow, j + 3).Value = data(i, j) ' D列から開始
        Next j

        targetLastRow = targetLastRow + 1
    Next i

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub


Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    If name = "" Then
        fullPath = path
    Else
        fullPath = path & "\" & name
    End If
    Debug.Print "fullPath" & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

Function GetNewID(ByVal targetSheet As Worksheet, ByVal IDPrefix As String) As String
    Dim lastRow As Long
    Dim currentID As String
    Dim currentNumber As Integer

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row
    If lastRow > 1 Then
        currentID = targetSheet.Cells(lastRow, "C").Value
        currentNumber = Val(Mid(currentID, Len(IDPrefix) + 1)) + 1
    Else
        currentNumber = 1
    End If
    GetNewID = IDPrefix & Format(currentNumber, "00000")
End Function

