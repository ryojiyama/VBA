Attribute VB_Name = "InspectionSheet"
Sub InspectionSheet_Make()
    Call SetupInspectionReport
'    Call TransferDataToAppropriateSheets
'    Call TransferDataToTopImpactTest
'    Call TransferDataToDynamicSheets
'    Call ImpactValueJudgement
'    Call FormatNonContinuousCells
'    Call DistributeChartsToSheets
End Sub

'*******************************************************************************
' メインプロシージャ
' 機能：検査報告書シートを作成し、シート名を管理
' 引数：なし
'*******************************************************************************
Sub SetupInspectionReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row

    Dim groupedData As Object
    Set groupedData = CreateObject("Scripting.Dictionary")
    Dim copiedSheets As Object
    Set copiedSheets = CreateObject("Scripting.Dictionary")
    Dim copiedSheetNames As Collection
    Set copiedSheetNames = New Collection

    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = ws.Cells(i, 2).value

        Dim HelmetData As New HelmetData
        Set HelmetData = ParseHelmetData(cellValue)

'        Dim productNameKey As String
'        productNameKey = HelmetData.GroupNumber & "-" & HelmetData.ProductName

        If Not groupedData.Exists(HelmetData.GroupNumber) Then
            groupedData.Add HelmetData.GroupNumber, New Collection
        End If
        groupedData(HelmetData.GroupNumber).Add HelmetData

        If Not copiedSheets.Exists(HelmetData.productName) Then
            ' 3種類のシートをコピーし、連番で名前を設定
            Dim sheetIndex As Long
            Dim sheetName As Variant
            For sheetIndex = 1 To 3
                sheetName = Array("Report01", "Report02", "Report03")(sheetIndex - 1)
                ThisWorkbook.Sheets(sheetName).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                ActiveSheet.Name = CreateUniqueName(HelmetData.productName & "_" & sheetIndex)
                copiedSheetNames.Add ActiveSheet.Name
            Next sheetIndex

            copiedSheets.Add HelmetData.productName, Nothing ' コピー済みフラグ設定
        End If

    Next i

    Debug.Print "Grouped Data:"
    PrintGroupedData groupedData
    SaveCopiedSheetNames copiedSheetNames
End Sub

'*******************************************************************************
' 機能：データ文字列を解析してHelmetDataオブジェクトを作成
' 引数：value - ハイフン区切りのデータ文字列
' 戻値：HelmetData - 解析されたデータオブジェクト
'*******************************************************************************
Function ParseHelmetData(value As String) As HelmetData
    Dim parts() As String
    parts = Split(value, "-")
    Dim result As New HelmetData
    
    If UBound(parts) >= 4 Then
        result.GroupNumber = parts(0)
        result.productName = parts(1)
        result.ImpactPosition = parts(2)
        result.ImpactTemp = parts(3)
        result.anvilForm = parts(4)
        result.headModel = parts(5)
    End If
    
    Set ParseHelmetData = result
End Function
'*******************************************************************************
' 機能：重複しないユニークなシート名を生成
' 引数：baseName - 基本となるシート名
' 戻値：String - ユニークなシート名
'*******************************************************************************
Function CreateUniqueName(baseName As String) As String
    Dim uniqueName As String
    uniqueName = baseName
    Dim count As Integer
    count = 1
    While SheetExists(uniqueName)
        uniqueName = baseName & count
        count = count + 1
    Wend
    CreateUniqueName = uniqueName ' 正しい戻り値の設定
End Function
'*******************************************************************************
' 機能：指定されたシート名が存在するか確認
' 引数：sheetName - 確認するシート名
' 戻値：Boolean - シートが存在する場合True
'*******************************************************************************
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing ' 正しい戻り値の設定
End Function

'*******************************************************************************
' 機能：グループ化されたデータをデバッグウィンドウに出力
' 引数：groupedData - 出力するデータオブジェクト
'*******************************************************************************
Private Sub PrintGroupedData(groupedData As Object)
    Dim key As Variant, item As HelmetData
    For Each key In groupedData.Keys
        Debug.Print "GroupNumber: " & key
        For Each item In groupedData(key)
            Debug.Print "  ProductName: " & item.productName
            Debug.Print "  ImpactPosition: " & item.ImpactPosition
            Debug.Print "  ImpactTemp: " & item.ImpactTemp
            Debug.Print "  Anvil: " & item.anvilForm
            Debug.Print "  Head: " & item.headModel
            Debug.Print "----------------------------"
        Next item
        Debug.Print "============================"
    Next key
End Sub
'*******************************************************************************
' 機能：コピーしたシート名をCopiedSheetNamesシートに保存
' 引数：sheetNames - 保存するシート名のコレクション
'*******************************************************************************
Sub SaveCopiedSheetNames(sheetNames As Collection)
' SetupInspectionReportのサブプロシージャ
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "CopiedSheetNames"
    End If

    ws.Cells.ClearContents

    Dim i As Long
    For i = 1 To sheetNames.count
        ws.Cells(i, 1).value = sheetNames(i)
    Next i
End Sub
'*******************************************************************************
' メインプロシージャ
' 機能：検査記録の転記と特定レコードの移動を実行
' 引数：なし
'*******************************************************************************
Sub ManageInspectionRecords()
    Call TransferDataToInspectionReports
    Call MoveSpecificRecords
End Sub

'*******************************************************************************
' 機能：LOG_Bicycleシートのデータを検査報告書に転記
' 引数：なし
'*******************************************************************************
Sub TransferDataToInspectionReports()
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    Dim wsTarget As Worksheet
    Dim i As Long
    Dim productNameKey As String
    Dim dataRange As Range
    Dim targetRow As Long

    ' LOG_Helmetシートの各行をループして処理します
    For i = 2 To lastRow
        ' GroupNumberとProductNameからproductNameKeyを構築します
        Dim parts() As String
        parts = Split(wsSource.Cells(i, "B").value, "-")
        productNameKey = parts(1) & "-" & parts(0)
        Dim productName As String
        productName = parts(1) ' "500" など

        Dim sheetIndex As Long
        Dim numericPart As Long
        numericPart = CLng(Split(productNameKey, "-")(1))

        If numericPart >= 1 And numericPart <= 6 Then ' 数値部分が1から6の場合
            ' シートインデックスを計算 (productName-4 も含む)
            Select Case numericPart
                Case 1: sheetIndex = 1
                Case 2, 3: sheetIndex = 2
                Case 4, 5, 6: sheetIndex = 3 ' productName-4 は productName_3 に転記
            End Select

            Dim targetSheetName As String
            targetSheetName = productName & "_" & sheetIndex

            TransferData productName, sheetIndex, i

        Else ' 数値部分が1から6以外の場合は転記しない
            Debug.Print "productNameKey: " & productNameKey & " は範囲外のため転記されません。"
        End If
    Next i
End Sub
'*******************************************************************************
' 機能：指定された製品名とインデックスに基づきデータを転記
' 引数：productName - 製品名
'       sheetIndex - シートインデックス
'       sourceRow - 転記元の行番号
'*******************************************************************************
Private Sub TransferData(productName As String, sheetIndex As Long, sourceRow As Long)
    Dim targetSheetName As String
    targetSheetName = productName & "_" & sheetIndex

    On Error Resume Next
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0

    If Not wsTarget Is Nothing Then
        ' ターゲットシートにヘッダーを転記する処理
        If wsTarget.Range("B30").value = "" Then ' ヘッダーが未転記であれば転記
            ThisWorkbook.Sheets("LOG_Bicycle").Range("B1:Z1").Copy Destination:=wsTarget.Range("B30")
        End If

        ' 最終行を見つけ、次の行からデータの転記を開始します
        targetRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row + 1
        If targetRow < 31 Then
            targetRow = 31 ' 最初のデータ転記開始位置をB31に設定
        End If

        ThisWorkbook.Sheets("LOG_Bicycle").Range("B" & sourceRow & ":Z" & sourceRow).Copy Destination:=wsTarget.Range("B" & targetRow)

        Set wsTarget = Nothing ' wsTarget をリセット
    End If
End Sub

'*******************************************************************************
' 機能：試料IDに'4がつくデータを_3シートから_2シートに移動
' 引数：なし
'*******************************************************************************
Sub MoveSpecificRecords()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim productName As String
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim headerRow As Long
    Dim sampleIDColumn As Long, testLocationColumn As Long
    Dim i As Long, j As Long ' ループカウンタ j を追加

    headerRow = 30 ' ヘッダー行

    ' 各シートをループ処理 (productName_3 シートを対象とする)
    For Each ws In ThisWorkbook.Worksheets
        If Right(ws.Name, 2) = "_3" Then ' シート名が "_3" で終わるシートのみ処理
            productName = Left(ws.Name, Len(ws.Name) - 2) ' productName を取得

            ' シートの有無を確認
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Worksheets(productName & "_2")
            Set wsSource = ThisWorkbook.Worksheets(productName & "_3") ' wsSource をここで設定
            On Error GoTo 0
            If wsTarget Is Nothing Then
                MsgBox productName & "_2 シートが存在しません。", vbCritical
                Exit Sub
            End If
            If wsSource Is Nothing Then
                MsgBox productName & "_3 シートが存在しません。", vbCritical
                Exit Sub
            End If

            ' "試料ID" と "試験箇所" の列番号を取得
            For j = 1 To wsSource.Cells(headerRow, Columns.count).End(xlToLeft).Column
                If wsSource.Cells(headerRow, j).value = "試料ID" Then
                    sampleIDColumn = j
                ElseIf wsSource.Cells(headerRow, j).value = "試験箇所" Then
                    testLocationColumn = j
                End If
            Next j

            If sampleIDColumn = 0 Or testLocationColumn = 0 Then
                MsgBox "「試料ID」または「試験箇所」のヘッダーが見つかりません。", vbCritical
                Exit Sub
            End If

            lastRowSource = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row
            lastRowTarget = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row

            ' 転記元シートのデータをループ処理 (下から上にループ)
            For i = lastRowSource To headerRow + 1 Step -1
                If wsSource.Cells(i, sampleIDColumn).value = 4 And _
                   (wsSource.Cells(i, testLocationColumn).value = "前頭部" Or wsSource.Cells(i, testLocationColumn).value = "後頭部") Then

                    ' データを転記先シートにコピー
                    wsSource.Rows(i).EntireRow.Copy Destination:=wsTarget.Rows(lastRowTarget + 1)
                    ' 転記先シートの最終行を更新
                    lastRowTarget = lastRowTarget + 1
                    wsSource.Rows(i).Delete
                End If
            Next i
        End If
    Next ws
End Sub



Sub CustomizeReportProcess()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long, checkRow As Long
    Dim sourceData As String
    Dim parts() As String
    Dim baseSheetName As String
    Dim ws As Worksheet
    Dim foundSheets As Collection
    Dim targetSheet As Worksheet
    Dim isValidData As Boolean
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "D").End(xlUp).Row
    
    ' Excelのパフォーマンス向上のための設定
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' B列のデータを取得して処理
    sourceData = wsSource.Cells(2, "B").value
    parts = Split(sourceData, "-")
    
    ' エラーチェック：データ形式の確認
    If UBound(parts) < 4 Then
        MsgBox "データ形式が不正です: " & sourceData, vbExclamation
        GoTo CleanExit
    End If
    
    ' D列のデータ検証
    isValidData = True
    For checkRow = 2 To lastRow
        If wsSource.Cells(checkRow, "D").value <> parts(1) Then
            isValidData = False
            Exit For
        End If
    Next checkRow
    
    ' データ検証結果の確認
    If Not isValidData Then
        MsgBox "エラー: D列に異なる値が存在します。処理を中止します。" & vbCrLf & _
               "期待値: " & parts(1) & vbCrLf & _
               "確認行: " & checkRow, vbCritical
        GoTo CleanExit
    End If
    
    ' シート名のベース部分を生成
    baseSheetName = parts(1) & "_1"
    
    ' 該当するシートを探索
    Set foundSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, baseSheetName) > 0 Then
            foundSheets.Add ws
        End If
    Next ws
    
    ' 見つかったシートがない場合の処理
    If foundSheets.count = 0 Then
        MsgBox "警告: " & baseSheetName & " に該当するシートが見つかりません。", vbExclamation
        GoTo CleanExit
    End If
    
    ' 見つかった各シートに対して処理を実行
    For Each targetSheet In foundSheets
        ' データの転記処理
        With targetSheet
            .Range("D3").value = wsSource.Cells(2, "D").value
            .Range("D4").value = wsSource.Cells(2, "O").value
            .Range("D5").value = wsSource.Cells(2, "E").value
            .Range("D6").value = wsSource.Cells(2, "Q").value
            .Range("I3").value = wsSource.Cells(2, "F").value
            .Range("I4").value = wsSource.Cells(2, "G").value
        End With
    Next targetSheet
    
CleanExit:
    ' Excelの設定を元に戻す
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    ' エラー発生時の処理
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

