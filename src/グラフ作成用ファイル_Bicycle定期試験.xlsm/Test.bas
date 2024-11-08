Attribute VB_Name = "Test"
Sub TransferDataToDynamicSheets()

    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim destinationSheetName As String
    Dim productNum As String

    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    ' Excelのパフォーマンス向上のための設定
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSourceのC列をループしてデータを処理
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, "B").value
        parts = Split(sourceData, "-")

        ' シート名を生成し、wsDesitinationに代入。
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(1) & "_" & 1

            ' 転記先シートの存在確認
            On Error Resume Next
            Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
            On Error GoTo 0

            ' シートが存在する場合にデータを転記
            If Not wsDestination Is Nothing Then
                productNum = wsSource.Cells(i, "D").value
                wsDestination.Range("D3").value = "No." & Left(productNum, Len(productNum) - 1) & "-" & Right(productNum, 1)
                wsDestination.Range("D4").value = wsSource.Cells(i, "O").value
                wsDestination.Range("D5").value = wsSource.Cells(i, "E").value
                wsDestination.Range("D6").value = wsSource.Cells(i, "Q").value
                wsDestination.Range("I3").value = wsSource.Cells(i, "F").value
                wsDestination.Range("I4").value = wsSource.Cells(i, "G").value
                ' 試験データの転記
                wsDestination.Range("D22").value = wsSource.Cells(i, "J").value
                wsDestination.Range("D23").value = wsSource.Cells(i, "L").value
            End If
        End If
    Next i

    ' Excelの設定を元に戻す
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
' 文字列変換用の関数
Function ConvertCompareString(ByVal strValue As String) As String
    ' 頭部関連の変換
    strValue = Replace(strValue, "前頭部", "前")
    strValue = Replace(strValue, "後頭部", "後")
    strValue = Replace(strValue, "右側頭部", "右")
    strValue = Replace(strValue, "左側頭部", "左")
    
    ' 形状関連の変換
    strValue = Replace(strValue, "平面", "平")
    strValue = Replace(strValue, "半球", "球")
    
    ConvertCompareString = strValue
End Function

Sub 転記処理()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long, sheetNum As Long, j As Long, k As Long
    Dim productCode As String, productName As String, productSheetName As String
    Dim parts() As String, inspectionSheetPartsB() As String, inspectionSheetPartsG() As String
    Dim impactCellB As String, impactCellG As String
    Dim impactRowsB As Variant, impactRowsG As Variant
    Dim foundB As Boolean, foundG As Boolean
    Dim mergeArea As Range

    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = logSheet.Cells(Rows.count, "B").End(xlUp).Row

    For i = 2 To lastRow
        productCode = logSheet.Cells(i, "B").value
        parts = Split(productCode, "-")
        productName = parts(1)
        foundB = False
        foundG = False

        For sheetNum = 1 To 3
            productSheetName = productName & "_" & sheetNum

            On Error Resume Next
            Set productSheet = ThisWorkbook.Sheets(productSheetName)
            On Error GoTo 0

            If Not productSheet Is Nothing Then
                impactRowsB = FindAllRows(productSheet, "B", "衝撃点&アンビル")
                impactRowsG = FindAllRows(productSheet, "G", "衝撃点&アンビル")

                ' B列の処理
                If IsArray(impactRowsB) Then
                    For j = LBound(impactRowsB) To UBound(impactRowsB)
                        If productSheet.Cells(impactRowsB(j), "B").MergeCells Then
                            Set mergeArea = productSheet.Cells(impactRowsB(j), "B").mergeArea
                            Dim nextColB As Long
                            nextColB = mergeArea.Column + mergeArea.Columns.count
                            impactCellB = productSheet.Cells(impactRowsB(j), nextColB).value

                            If Len(Trim(impactCellB)) > 0 Then
                                inspectionSheetPartsB = Split(impactCellB, "・")
                                If UBound(inspectionSheetPartsB) >= 1 Then
                                    Dim convertedFirst As String, convertedSecond As String
                                    convertedFirst = ConvertCompareString(inspectionSheetPartsB(0))
                                    convertedSecond = ConvertCompareString(inspectionSheetPartsB(1))

                                    If parts(2) = convertedFirst And parts(4) = convertedSecond Then
                                        productSheet.Cells(impactRowsB(j), nextColB).value = logSheet.Cells(i, "J").value
                                        foundB = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If

                ' G列の処理
                If IsArray(impactRowsG) Then
                    For k = LBound(impactRowsG) To UBound(impactRowsG)
                        If productSheet.Cells(impactRowsG(k), "G").MergeCells Then
                            Set mergeArea = productSheet.Cells(impactRowsG(k), "G").mergeArea
                            Dim nextColG As Long
                            nextColG = mergeArea.Column + mergeArea.Columns.count
                            impactCellG = productSheet.Cells(impactRowsG(k), nextColG).value

                            If Len(Trim(impactCellG)) > 0 Then
                                inspectionSheetPartsG = Split(impactCellG, "・")
                                If UBound(inspectionSheetPartsG) >= 1 Then
                                    convertedFirst = ConvertCompareString(inspectionSheetPartsG(0))
                                    convertedSecond = ConvertCompareString(inspectionSheetPartsG(1))

                                    If parts(2) = convertedFirst And parts(4) = convertedSecond Then
                                        productSheet.Cells(impactRowsG(k), nextColG).value = logSheet.Cells(i, "J").value
                                        foundG = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If

                If foundB Or foundG Then Exit For
            End If
            Set productSheet = Nothing
        Next sheetNum
    Next i
End Sub

Function FindAllRows(sheet As Worksheet, Col As String, searchStr As String) As Variant
    Dim result(0 To 1000) As Long
    Dim resultCount As Long
    Dim lastRow As Long
    Dim i As Long
    Dim mergeArea As Range

    resultCount = -1
    lastRow = sheet.Cells(sheet.Rows.count, Col).End(xlUp).Row

    For i = 1 To lastRow
        If sheet.Cells(i, Col).MergeCells Then
            Set mergeArea = sheet.Cells(i, Col).mergeArea
            If sheet.Cells(i, Col).Address = mergeArea.Cells(1, 1).Address Then
                If InStr(1, mergeArea.Cells(1, 1).value, searchStr) > 0 Then
                    resultCount = resultCount + 1
                    result(resultCount) = i
                End If
            End If
        Else
            If InStr(1, sheet.Cells(i, Col).value, searchStr) > 0 Then
                resultCount = resultCount + 1
                result(resultCount) = i
            End If
        End If
    Next i

    If resultCount >= 0 Then
        Dim finalResult() As Long
        ReDim finalResult(0 To resultCount)
        For i = 0 To resultCount
            finalResult(i) = result(i)
        Next i
        FindAllRows = finalResult
    Else
        FindAllRows = Array()
    End If
End Function

Sub セル値確認()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim targetCells As Variant
    Dim i As Long, j As Long
    
    ' 確認するシート名を配列に格納
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' 確認するセルの位置を配列に格納 (列, 行)
    targetCells = Array(Array("B", 21), Array("B", 25), Array("G", 21), Array("G", 25))
    
    Debug.Print "セル値確認開始"
    Debug.Print "-------------------"
    
    ' 各シートをループ
    For i = 0 To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            Debug.Print sheetNames(i) & " のセル値:"
            
            ' 各対象セルをループ
            For j = 0 To UBound(targetCells)
                Dim Col As String
                Dim Row As Long
                Col = targetCells(j)(0)
                Row = targetCells(j)(1)
                
                ' セルが結合されているか確認
                If ws.Cells(Row, Col).MergeCells Then
                    Dim mergeArea As Range
                    Set mergeArea = ws.Cells(Row, Col).mergeArea
                    Debug.Print "  " & Col & Row & ": " & mergeArea.Cells(1, 1).value & _
                              " (結合セル: " & mergeArea.Address & ")"
                Else
                    Debug.Print "  " & Col & Row & ": " & ws.Cells(Row, Col).value
                End If
            Next j
            Debug.Print "-------------------"
        Else
            Debug.Print "シート " & sheetNames(i) & " が見つかりません"
            Debug.Print "-------------------"
        End If
    Next i
    
    Debug.Print "確認完了"
End Sub



' -------------------------------------------------------------------------------------------------------------
Sub GroupAndListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim part0 As String
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    ' アクティブシートのチャートオブジェクトをループ処理
    For Each chartObj In ActiveSheet.ChartObjects
        ' グラフにタイトルがあるかどうかをチェック
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"  ' タイトルがない場合
        End If

        ' chartNameを"-"で分割し、part(0)を取得
        part0 = Split(chartObj.Name, "-")(0)

        ' グループがまだ存在しない場合、新規作成
        If Not groups.Exists(part0) Then
            groups.Add part0, New Collection
        End If

        ' グループにチャート名とタイトルを追加
        groups(part0).Add "Chart Name: " & chartObj.Name & "; Title: " & chartTitle
    Next chartObj

    ' 各グループの内容をイミディエイトウィンドウに出力
    Dim key As Variant
    For Each key In groups.Keys
        Debug.Print "Group: " & key
        Dim item As Variant
        For Each item In groups(key)
            Debug.Print item
        Next item
    Next key
End Sub

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





