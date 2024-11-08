Attribute VB_Name = "CompareInspection_1600"
Option Explicit

' 試料情報を格納する型定義
Private Type TestSample
    Number As String
    Condition As String
    Row As Long
End Type

' 測定点情報を格納する型定義
Private Type MeasurementPoint
    Position As String
    Shape As String
    Row As Long
    Column As Long
End Type

' 状態の変換用定数
Private Const HIGH_TEMP As String = "Hot"
Private Const LOW_TEMP As String = "Cold"
Private Const WET_CONDITION As String = "Wet"

' 文字列変換用の辞書を初期化
Private Function InitializeConversionDicts() As Object
    Dim locationDict As Object
    Set locationDict = CreateObject("Scripting.Dictionary")
    
    ' 位置と形状の変換マッピング
    locationDict.Add "前頭部", "前"
    locationDict.Add "後頭部", "後"
    locationDict.Add "右側頭部", "右"
    locationDict.Add "左側頭部", "左"
    locationDict.Add "平面", "平"
    locationDict.Add "半球", "球"
    
    Set InitializeConversionDicts = locationDict
End Function

' メイン処理
Sub 転記処理()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim convDict As Object
    Dim sheetNames As Variant
    Dim sheetName As Variant
    
    ' 初期化
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    Set convDict = InitializeConversionDicts()
    
    ' 処理対象シート名の配列
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' 各製品シートを処理
    For Each sheetName In sheetNames
        Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
        ProcessProductSheet productSheet, logSheet, convDict
    Next sheetName
End Sub

' 製品シート処理
Private Sub ProcessProductSheet(ByRef productSheet As Worksheet, _
                              ByRef logSheet As Worksheet, _
                              ByRef convDict As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSample As TestSample
    Dim hasSample As Boolean
    Dim cellValue As String
    
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    ' シートの各行を処理
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' 試料行の検出
        If InStr(1, cellValue, "試料") > 0 Then
            currentSample = GetSampleInfo(cellValue)
            currentSample.Row = i
            hasSample = True
        End If
        
        ' 衝撃点の検出と測定点処理
        If hasSample And InStr(1, cellValue, "衝撃点&アンビル") > 0 Then
            Debug.Print "衝撃点検出 - シート:" & productSheet.Name & ", 行:" & i & ", 値:" & cellValue
            ProcessMeasurementPoints productSheet, logSheet, i, currentSample, convDict
        End If
    Next i
End Sub

' 試料情報の取得
Private Function GetSampleInfo(ByVal cellValue As String) As TestSample
    Dim sample As TestSample
    Dim parts As Variant
    
    ' "試料1 高温" のような文字列を分解
    parts = Split(cellValue)
    
    ' 試料番号（2桁に整形）
    sample.Number = Format(Val(Mid(parts(0), 3)), "00")
    
    ' 状態の判定
    Select Case parts(1)
        Case "高温"
            sample.Condition = HIGH_TEMP
        Case "低温"
            sample.Condition = LOW_TEMP
        Case "浸せき"
            sample.Condition = WET_CONDITION
    End Select
    
    ' デバッグ出力
    Debug.Print cellValue & " → " & sample.Number & ", " & sample.Condition
    
    GetSampleInfo = sample
End Function

' 測定点の処理
Private Sub ProcessMeasurementPoints(ByRef productSheet As Worksheet, _
                                   ByRef logSheet As Worksheet, _
                                   ByVal currentRow As Long, _
                                   ByRef sample As TestSample, _
                                   ByRef convDict As Object)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim point As MeasurementPoint
    Dim searchCode As String
    Dim logLastRow As Long
    Dim i As Long
    Dim valueCell As Range
    Dim valueCellBelow As Range
    Dim foundMatch As Boolean
    Dim skippedLogs As String

    targetColumns = Array(2, 7)  ' B=2, G=7
    logLastRow = logSheet.Cells(Rows.count, "B").End(xlUp).Row
    skippedLogs = ""

    For Each colIndex In targetColumns
        point = GetMeasurementPoint(productSheet, currentRow, CLng(colIndex), convDict)
        If Len(point.Position) > 0 Then
            searchCode = sample.Number & "-500S-" & point.Position & "-" & _
                        sample.Condition & "-" & point.Shape

            Set valueCell = productSheet.Cells(currentRow + 1, CLng(colIndex) + 2)
            Set valueCellBelow = productSheet.Cells(currentRow + 2, CLng(colIndex) + 2)
            foundMatch = False

            For i = 2 To logLastRow
                Dim logValue As String
                logValue = logSheet.Cells(i, "B").value

                If Replace(logValue, "-E", "") = searchCode Then
                    foundMatch = True

                    If Len(Trim(logSheet.Cells(i, "V").value)) = 0 Then
                        ' 最初の値（J列）の転記
                        If valueCell.MergeCells Then
                            valueCell.mergeArea.item(1).value = logSheet.Cells(i, "J").value
                        Else
                            valueCell.value = logSheet.Cells(i, "J").value
                        End If

                        ' 二つ目の値（L列）の転記
                        If valueCellBelow.MergeCells Then
                            valueCellBelow.mergeArea.item(1).value = logSheet.Cells(i, "L").value
                        Else
                            valueCellBelow.value = logSheet.Cells(i, "L").value
                        End If

                        logSheet.Cells(i, "V").value = "済"
                    Else
                        ' スキップしたログを記録
                        skippedLogs = skippedLogs & "シート: " & productSheet.Name & _
                                    ", コード: " & logValue & _
                                    ", LOG行: " & i & _
                                    ", 値1: " & logSheet.Cells(i, "J").value & _
                                    ", 値2: " & logSheet.Cells(i, "L").value & vbCrLf
                    End If
                    Exit For
                End If
            Next i
        End If
    Next colIndex

    ' スキップしたログがある場合、最後にまとめて表示
    If Len(skippedLogs) > 0 Then
        MsgBox "以下のデータは既に転記済みのためスキップされました：" & vbCrLf & vbCrLf & _
               skippedLogs, vbInformation, "転記スキップログ"
    End If
End Sub

' 測定点情報の取得
Private Function GetMeasurementPoint(ByRef sheet As Worksheet, _
                                   ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   ByRef convDict As Object) As MeasurementPoint
    Dim point As MeasurementPoint
    Dim targetCell As Range
    Dim cellValue As String
    Dim parts() As String
    
    Set targetCell = sheet.Cells(Row, Col + 1)
    
    If targetCell.MergeCells Then
        cellValue = Trim(targetCell.mergeArea.item(1).Offset(0, 1).value)
    Else
        cellValue = Trim(targetCell.Offset(0, 1).value)
    End If
    
    If Len(cellValue) > 0 Then
        parts = Split(cellValue, "・")
        If UBound(parts) >= 1 Then
            point.Position = convDict(parts(0))
            point.Shape = convDict(parts(1))
        End If
    End If
    
    GetMeasurementPoint = point
End Function


' チャート分配のメイン処理
' チャート分配のメイン処理
Sub チャート分配処理()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim chartErrorLogs As String
    
    ' 初期化
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    chartErrorLogs = ""
    
    Debug.Print "========== チャート分配処理開始 ==========" & vbCrLf
    
    ' 処理対象シート名の配列
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' 各製品シートを処理
    For Each sheetName In sheetNames
        Debug.Print "シート処理開始: " & CStr(sheetName)
        Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
        ProcessChartDistribution productSheet, logSheet, chartErrorLogs
    Next sheetName
    
    ' エラーログの表示とデバッグ出力
    If Len(chartErrorLogs) > 0 Then
        Dim logMessage As String
        logMessage = "チャート処理で以下の問題が発生しました：" & vbCrLf & vbCrLf & chartErrorLogs
        
        ' イミディエイトウィンドウに出力
        Debug.Print "---------- エラーログ ----------"
        Debug.Print logMessage
        Debug.Print "--------------------------------"
        
        ' メッセージボックスで表示
        MsgBox logMessage, vbInformation, "チャート処理結果"
    End If
    
    Debug.Print "========== チャート分配処理終了 ==========" & vbCrLf
End Sub

' 各シートのチャート分配処理
Private Sub ProcessChartDistribution(ByRef productSheet As Worksheet, _
                                   ByRef logSheet As Worksheet, _
                                   ByRef errorLogs As String)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSample As TestSample
    Dim hasSample As Boolean
    Dim cellValue As String
    Dim convDict As Object
    
    ' 初期化
    Set convDict = InitializeConversionDicts()
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    Debug.Print "シート[" & productSheet.Name & "] 処理開始 - 最終行: " & lastRow
    
    ' シートの各行を処理
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' 試料行の検出
        If InStr(1, cellValue, "試料") > 0 Then
            currentSample = GetSampleInfo(cellValue)
            hasSample = True
            Debug.Print "  試料検出: " & cellValue & " -> サンプル番号: " & currentSample.Number & ", 状態: " & currentSample.Condition
        End If
        
        ' 衝撃点の検出と処理
        If hasSample And InStr(1, cellValue, "衝撃点&アンビル") > 0 Then
            Debug.Print "  衝撃点検出 - 行: " & i & ", 値: " & cellValue
            ProcessChartPoints productSheet, logSheet, i, currentSample, convDict, errorLogs
        End If
    Next i
    
    Debug.Print "シート[" & productSheet.Name & "] 処理終了" & vbCrLf
End Sub

' 測定点のチャート処理
Private Sub ProcessChartPoints(ByRef productSheet As Worksheet, _
                             ByRef logSheet As Worksheet, _
                             ByVal currentRow As Long, _
                             ByRef sample As TestSample, _
                             ByRef convDict As Object, _
                             ByRef errorLogs As String)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim point As MeasurementPoint
    Dim searchCode As String
    Dim valueCell As Range
    
    targetColumns = Array(2, 7)  ' B=2, G=7
    
    For Each colIndex In targetColumns
        point = GetMeasurementPoint(productSheet, currentRow, CLng(colIndex), convDict)
        If Len(point.Position) > 0 Then
            ' 検索用コードの生成
            searchCode = sample.Number & "-500S-" & point.Position & "-" & _
                        sample.Condition & "-" & point.Shape & "-E"
            
            Debug.Print "    検索コード生成: " & searchCode
            
            ' 値が記入されるセルの位置を取得
            Set valueCell = productSheet.Cells(currentRow + 1, CLng(colIndex) + 2)
            Debug.Print "    対象セル: " & valueCell.Address
            
            ' チャートのコピー処理
            CopyMatchingChart logSheet, productSheet, searchCode, valueCell, errorLogs
        End If
    Next colIndex
End Sub

' チャートのコピー処理
Private Sub CopyMatchingChart(ByRef sourceSheet As Worksheet, _
                            ByRef targetSheet As Worksheet, _
                            ByVal searchID As String, _
                            ByRef targetCell As Range, _
                            ByRef errorLogs As String)
    Dim cht As ChartObject
    Dim foundCharts As Long
    Dim retryCount As Integer
    Const MAX_RETRIES As Integer = 3
    
    On Error Resume Next
    
    foundCharts = 0
    Debug.Print "      チャート検索開始 - ID: " & searchID
    
    ' ソースシートの全チャートをループ
    For Each cht In sourceSheet.ChartObjects
        Debug.Print "        確認中のチャート - Name: " & cht.Name
        
        ' チャートIDと検索IDが一致する場合
        If cht.Name = searchID Then
            foundCharts = foundCharts + 1
            
            ' 最初に見つかったチャートの場合
            If foundCharts = 1 Then
                retryCount = 0
                Do
                    ' クリップボードをクリア
                    Application.CutCopyMode = False
                    Err.Clear
                    
                    ' チャートをコピー
                    cht.Copy
                    
                    If Err.Number = 0 Then
                        ' 少し待機してからペースト
                        Application.Wait Now + TimeValue("00:00:00.2")
                        targetSheet.Paste targetCell.Offset(0, 3)
                        
                        If Err.Number = 0 Then
                            Debug.Print "        チャートを複製: " & targetCell.Offset(0, 3).Address & " に配置成功"
                            Exit Do
                        End If
                    End If
                    
                    ' エラーが発生した場合
                    If Err.Number <> 0 Then
                        retryCount = retryCount + 1
                        If retryCount >= MAX_RETRIES Then
                            errorLogs = errorLogs & "配置エラー - シート: " & targetSheet.Name & _
                                      ", ID: " & searchID & _
                                      ", エラー: " & Err.Description & vbCrLf
                            Exit Do
                        End If
                        ' 再試行前に少し長めに待機
                        Application.Wait Now + TimeValue("00:00:00.5")
                    End If
                Loop While retryCount < MAX_RETRIES
                
            Else
                errorLogs = errorLogs & "重複チャート - シート: " & targetSheet.Name & _
                           ", ID: " & searchID & _
                           " (" & foundCharts & "個目)" & vbCrLf
            End If
        End If
    Next cht
    
    If foundCharts = 0 Then
        errorLogs = errorLogs & "未発見チャート - シート: " & targetSheet.Name & _
                    ", ID: " & searchID & vbCrLf
    End If
    
    ' 処理完了後にクリップボードをクリア
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub

