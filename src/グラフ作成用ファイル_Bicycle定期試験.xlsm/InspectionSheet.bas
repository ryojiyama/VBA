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
' 既存のシートをコピーし、productName_1 などの名前をつける。
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
                sheetName = Array("InspectionSheet01", "InspectionSheet02", "InspectionSheet03")(sheetIndex - 1)
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
Function ParseHelmetData(value As String) As HelmetData
' SetupInspectionReportのサブプロシージャ
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

Function CreateUniqueName(baseName As String) As String
' SetupInspectionReportのサブプロシージャ
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
Function SheetExists(sheetName As String) As Boolean
' SetupInspectionReportのサブプロシージャ
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing ' 正しい戻り値の設定
End Function
Private Sub PrintGroupedData(groupedData As Object)
' SetupInspectionReportのサブプロシージャ
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

' コピーしたシートにヘッダーと試験結果を転記する。
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

Private Sub TransferData(productName As String, sheetIndex As Long, sourceRow As Long)
' TransferDataToInspectionReportsのサブプロシージャ。データ転記処理を関数化
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

' _4のデータのみを転記する
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






    
    
    '天頂試験のみのシートを作成する。
Sub TransferDataToTopImpactTest()
    '"Log_Helmet"からコピーした検査票に値を転記する。
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim firstDashPos As Integer
    Dim secondDashPos As Integer
    Dim matchName As String
    Dim TemperatureCondition As String

    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("Log_Bicycle")

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    ' 2行目から最終行までループ
    For i = 2 To lastRow
        ' C列の値から製品コードを取得
        firstDashPos = InStr(wsSource.Cells(i, "B").value, "-")
        If firstDashPos > 0 Then
            secondDashPos = InStr(firstDashPos + 1, wsSource.Cells(i, "B").value, "-")
            If secondDashPos > 0 Then
                matchName = Left(wsSource.Cells(i, "B").value, secondDashPos - 1)
            End If
        End If

        ' 各シートをループして条件に一致するシートを検索
        For Each wsDestination In ThisWorkbook.Sheets
            If wsDestination.Name = matchName Then ' シート名が製品コードに一致するか確認
                ' 条件に一致した場合、転記を実行
                ' 以下のコードは変更なし
                wsDestination.Range("C2").value = wsSource.Cells(i, 21).value
                wsDestination.Range("F2").value = wsSource.Cells(i, 6).value
                wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                wsDestination.Range("A10").value = "※前処理：" & wsSource.Cells(i, 12).value
                wsDestination.Range("A14").value = "検査対象外"
                wsDestination.Range("A19").value = "検査対象外"
                Exit For ' 転記後は次の行へ
            End If
        Next wsDestination
    Next i
End Sub

' productName_1のシートに転記する。
Sub TransferDataToDynamicSheets()

    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim destinationSheetName As String

    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row
    
    ' Excelのパフォーマンス向上のための設定
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSourceのC列をループしてデータを処理
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, "B").value
        checkData = wsSource.Cells(i, 5).value
        parts = Split(sourceData, "-")

        ' シート名の生成
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(0) & "-" & parts(1)

            ' 転記先シートの存在確認
            On Error Resume Next
            Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
            On Error GoTo 0

            ' シートが存在し、かつ条件が一致する場合にデータを転記
            If Not wsDestination Is Nothing Then
                Select Case parts(2)
                    Case "天"
                        If checkData = "天頂" Then
                            ' 天に関するデータ転記
                            wsDestination.Range("C2").value = wsSource.Cells(i, 21).value
                            wsDestination.Range("F2").value = wsSource.Cells(i, 6).value
                            wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                            wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                            wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                            wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                            wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                            wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                            wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                            wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                            wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                            wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("A10").value = "※前処理：" & wsSource.Cells(i, 12).value
                        End If
                    Case "前"
                        If checkData = "前頭部" Then
                            ' 前頭部に関するデータ転記
                            wsDestination.Range("E13").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E14").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E15").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A13").value = "前頭部"
                        End If
                    Case "後"
                        If checkData = "後頭部" Then
                            ' 後頭部に関するデータ転記
                            wsDestination.Range("E17").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E18").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E19").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A17").value = "後頭部"
                        End If
                End Select
            End If
        End If
    Next i
    
    ' Excelの設定を元に戻す
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub ImpactValueJudgement()
    'CopiedSheetNamesシートのA列に基づいて各検査票シートの衝撃値を判定する
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim resultE11 As Boolean, resultE14 As Boolean, resultE19 As Boolean
    Dim targetSheets As Collection
    
    ' 処理するシート名を取得
    Set targetSheets = GetTargetSheetNames()
    
    ' 対象のシート名に基づいて処理を行う
    For i = 1 To targetSheets.count
        sheetName = targetSheets(i)
        ' 対象のシートを設定
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        
        ' D11, D14, D19の値を基に判定
        resultE11 = wsTarget.Range("E11").value <= 4.9
        resultE14 = IsEmpty(wsTarget.Range("E13")) Or wsTarget.Range("E13").value <= 9.81
        resultE19 = IsEmpty(wsTarget.Range("E17")) Or wsTarget.Range("E17").value <= 9.81
        
        ' 全ての条件がTrueの場合は"合格"、それ以外は"不合格"をG9に記入
        If resultE11 And resultE14 And resultE19 Then
            wsTarget.Range("H9").value = "合格"
        Else
            wsTarget.Range("H9").value = "不合格"
        End If
    Next i
End Sub

Function GetTargetSheetNames() As Collection
    ' CopiedSheetNamesシートのA列からシート名を取得
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetNames As New Collection
    
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        sheetNames.Add ws.Cells(i, 1).value
    Next i
    
    Set GetTargetSheetNames = sheetNames
End Function
    ' CopiedSheetNamesシートのA列に基づいて検査票に書式を設定する
Sub FormatNonContinuousCells()
    Dim wsTarget As Worksheet
    Dim i As Long
    Dim sheetName As String
    Dim targetSheets As Collection
    Dim rng As Range
    Dim cell As Range
    
    ' 処理するシート名を取得
    Set targetSheets = GetTargetSheetNames()
    
    ' 対象のシート名に基づいて処理を行う
    For i = 1 To targetSheets.count
        sheetName = targetSheets(i)
        
        ' ワークシートが存在するかチェック
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        ' ワークシートが存在すれば、指定したセル範囲に書式を設定
        If Not wsTarget Is Nothing Then
            ' 範囲と書式設定を関連付け
            FormatRange wsTarget.Range("E7"), "游明朝", 12, True
            FormatRange wsTarget.Range("E8"), "游明朝", 12, True
            FormatRange wsTarget.Range("E9"), "游明朝", 12, True

            ' E13に値がない場合、A14:E14とB15:D16をグレーアウト
            If IsEmpty(wsTarget.Range("E13").value) Then
                wsTarget.Range("A13").value = "検査対象外"
                FormatRange wsTarget.Range("A13"), "游ゴシック", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B13:F13, B14:E15"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A13"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("E13:E15"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If

            ' E17に値がない場合、A19:E19とB20:D21をグレーアウト
            If IsEmpty(wsTarget.Range("E17").value) Then
                wsTarget.Range("A17").value = "検査対象外"
                FormatRange wsTarget.Range("A17"), "游ゴシック", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B17:F17, B18:E19"), "游ゴシック", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A17"), "游ゴシック", 12, True
                FormatRange wsTarget.Range("E17:E19"), "游ゴシック", 10, False, RGB(255, 255, 255)
            End If
            
            ' 特定の文字に書式を適用
            FormatSpecificEndStrings wsTarget.Range("A10"), "游ゴシック", 12, True
            
            ' セルの書式設定
            With wsTarget.Range("C2:C4, F2:F4, H2:H4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("F3").NumberFormat = "0.0"" g"""
            wsTarget.Range("H2").NumberFormat = "0"" ℃"""
            wsTarget.Range("H3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("E11, E14, E19").NumberFormat = "0.00"" kN"""
            
            ' E14:E15, E18:E19の値に応じて書式を設定
            Set rng = wsTarget.Range("E14:E15, E18:E19")
            For Each cell In rng
                If cell.value <= 0.01 Then
                    cell.value = "―"
                Else
                    cell.NumberFormat = "0.00"" ms"""
                End If
            Next cell
            
            ' 他の範囲も同様に設定可能
            ' FormatRange wsTarget.Range("その他の範囲"), "フォント名", フォントサイズ, 太字かどうか, 背景色

            Set wsTarget = Nothing
        End If
    Next i
End Sub


Sub FormatSpecificEndStrings(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean)
    ' セルの特定の文字(前処理)に書式を適用するサブプロシージャ
    Dim cell As Range

    For Each cell In rng
        Dim text As String
        text = cell.value
        Dim textLength As Integer
        textLength = Len(text)

        If textLength >= 2 Then
            If Right(text, 2) = "高温" Or Right(text, 2) = "低温" Then
                With cell.Characters(Start:=textLength - 1, Length:=2).Font
                    .Name = fontName
                    .size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "浸せき" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .Name = fontName
                    .size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub

Sub FormatRange(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean, Optional bgColor As Variant)
    ' 範囲に書式を適用するためのサブプロシージャ
    With rng
        .Font.Name = fontName
        .Font.size = fontSize
        .Font.Bold = isBold
        If Not IsMissing(bgColor) Then
            .Interior.Color = bgColor
        Else
            .Interior.colorIndex = xlColorIndexAutomatic ' 背景色を自動に設定
        End If
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub
' チャートを各シートに分配する。
Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"シートを対象にする
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"シートのチャートオブジェクトをグループ分け
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartNameを"-"で分割し、sheetNameを取得
        parts = Split(chartObj.Name, "-")
        If UBound(parts) >= 1 Then
            sheetName = parts(0) & "-" & parts(1)
        Else
            sheetName = parts(0)
        End If
        
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' グループごとにチャートを対応するシートに移動
    Dim key As Variant
    For Each key In groups.Keys
        ' シートの存在を確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' シートが存在しない場合、チャートを移動しない
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' チャートの移動
            Dim chart As ChartObject
            For Each chart In groups(key)
                chart.chart.Location Where:=xlLocationAsObject, Name:=targetSheet.Name
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not moved."
        End If
    Next key
End Sub
