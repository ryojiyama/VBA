Attribute VB_Name = "ChartBunpai_Seikou"
' 2024-11-08作成 うまく行ったが、課題も多い。


' ===== 型定義を最初に配置 =====
Private Type sampleInfo
    SampleNumber As String
    Condition As String
End Type

Private Type pointInfo
    Position As String
    Shape As String
End Type

' ===== 定数定義 =====
Private Const CHART_SUFFIX As String = "-E"
Private Const SERIES_PREFIX As String = "500S"

' ===== 辞書オブジェクト格納用変数 =====
Private locationDict As Object
Private conditionDict As Object
Private shapeDict As Object
Private searchPatternDict As Object


' ===== 辞書の初期化 =====
Private Sub InitializeDictionaries()
    On Error GoTo ErrorHandler
    
    ' 既存の辞書をクリア
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    ' 位置変換用辞書
    Set locationDict = CreateObject("Scripting.Dictionary")
    With locationDict
        .Add "前頭部", "前"
        .Add "後頭部", "後"
        .Add "右側頭部", "右"
        .Add "左側頭部", "左"
    End With
    
    ' 状態変換用辞書
    Set conditionDict = CreateObject("Scripting.Dictionary")
    With conditionDict
        .Add "高温", "Hot"
        .Add "低温", "Cold"
        .Add "浸せき", "Wet"
    End With
    
    ' 形状変換用辞書
    Set shapeDict = CreateObject("Scripting.Dictionary")
    With shapeDict
        .Add "平面", "平"
        .Add "半球", "球"
    End With
    
    ' 検索パターン用辞書
    Set searchPatternDict = CreateObject("Scripting.Dictionary")
    With searchPatternDict
        .Add "format", "{0}-" & SERIES_PREFIX & "-{1}-{2}-{3}" & CHART_SUFFIX
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "辞書の初期化エラー: " & Err.Description
    
    ' 辞書オブジェクトのクリーンアップ
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    Err.Raise Err.Number, "InitializeDictionaries", "辞書の初期化に失敗しました"
End Sub

' ===== メインのチャート分配処理 =====
Sub チャート分配処理()
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.screenUpdating = False
    
    ' 辞書の初期化
    InitializeDictionaries
    
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim chartErrorLogs As String
    
    ' シートの存在確認
    If Not SheetExists("LOG_Bicycle") Then
        MsgBox "LOG_Bicycleシートが見つかりません。", vbCritical
        GoTo CleanUp
    End If
    
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    chartErrorLogs = ""
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' 各シートの処理
    For Each sheetName In sheetNames
        If SheetExists(CStr(sheetName)) Then
            Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
            ProcessSheet productSheet, logSheet, chartErrorLogs
        Else
            chartErrorLogs = chartErrorLogs & "シートが見つかりません: " & CStr(sheetName) & vbCrLf
        End If
    Next sheetName

CleanUp:
    ' 辞書オブジェクトの解放
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.screenUpdating = True
    
    If Len(chartErrorLogs) > 0 Then
        MsgBox "チャート処理で以下の問題が発生しました：" & vbCrLf & vbCrLf & _
               chartErrorLogs, vbInformation, "チャート処理結果"
    End If
    Exit Sub

ErrorHandler:
    chartErrorLogs = chartErrorLogs & "予期せぬエラー: " & Err.Description & vbCrLf
    Resume CleanUp
End Sub

' ===== シート処理 =====
Private Sub ProcessSheet(ByRef productSheet As Worksheet, _
                        ByRef logSheet As Worksheet, _
                        ByRef errorLogs As String)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSampleInfo As sampleInfo
    Dim hasSample As Boolean
    Dim cellValue As String
    
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' 試料情報の取得
        If InStr(1, cellValue, "試料") > 0 Then
            currentSampleInfo = ExtractSampleInfo(cellValue)
            hasSample = True
        End If
        
        ' 衝撃点の処理
        If hasSample And InStr(1, cellValue, "衝撃点&アンビル") > 0 Then
            ProcessMeasurementPoint productSheet, logSheet, i, currentSampleInfo, errorLogs
        End If
    Next i
End Sub

' ===== 試料情報抽出 =====
Private Function ExtractSampleInfo(ByVal cellValue As String) As sampleInfo
    Dim info As sampleInfo
    Dim parts As Variant
    
    On Error GoTo ErrorHandler
    
    parts = Split(cellValue)
    
    ' 配列の境界チェック
    If UBound(parts) >= 1 Then
        ' 試料番号の抽出（数値以外の文字を除去）
        Dim numStr As String
        numStr = Replace(Replace(parts(0), "試料", ""), " ", "")
        info.SampleNumber = Format(Val(numStr), "00")
        
        ' 状態の判定（辞書に存在するかチェック）
        If conditionDict.Exists(parts(1)) Then
            info.Condition = conditionDict(parts(1))
        Else
            info.Condition = "Unknown"
            Debug.Print "未定義の状態: " & parts(1)
        End If
    End If
    
    ExtractSampleInfo = info
    Exit Function
    
ErrorHandler:
    info.SampleNumber = "00"
    info.Condition = "Error"
    ExtractSampleInfo = info
End Function

' ===== 測定点処理 =====
Private Sub ProcessMeasurementPoint(ByRef productSheet As Worksheet, _
                                  ByRef logSheet As Worksheet, _
                                  ByVal currentRow As Long, _
                                  ByRef sampleInfo As sampleInfo, _
                                  ByRef errorLogs As String)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim pointInfo As pointInfo
    Dim chartId As String
    
    targetColumns = Array(2, 7)  ' B=2, G=7
    
    For Each colIndex In targetColumns
        pointInfo = ExtractPointInfo(productSheet, currentRow, CLng(colIndex))
        If Len(pointInfo.Position) > 0 Then
            ' チャートIDの生成
            chartId = GenerateChartId(logSheet, sampleInfo, pointInfo)
            
            ' チャートのコピー
            CopyChart logSheet, productSheet, chartId, _
                     productSheet.Cells(currentRow + 1, CLng(colIndex) + 2), errorLogs
        End If
    Next colIndex
End Sub

' ===== 測定点情報抽出 =====
Private Function ExtractPointInfo(ByRef sheet As Worksheet, _
                                ByVal Row As Long, _
                                ByVal Col As Long) As pointInfo
    Dim info As pointInfo
    Dim targetCell As Range
    Dim cellValue As String
    Dim parts() As String
    
    On Error GoTo ErrorHandler
    
    Set targetCell = sheet.Cells(Row, Col + 1)
    
    If targetCell.MergeCells Then
        cellValue = Trim(targetCell.mergeArea.item(1).Offset(0, 1).value)
    Else
        cellValue = Trim(targetCell.Offset(0, 1).value)
    End If
    
    If Len(cellValue) > 0 Then
        parts = Split(cellValue, "・")
        If UBound(parts) >= 1 Then
            ' 辞書の存在チェック
            If locationDict.Exists(parts(0)) And shapeDict.Exists(parts(1)) Then
                info.Position = locationDict(parts(0))
                info.Shape = shapeDict(parts(1))
            Else
                Debug.Print "未定義の位置または形状: " & cellValue
            End If
        End If
    End If
    
    ExtractPointInfo = info
    Exit Function
    
ErrorHandler:
    info.Position = ""
    info.Shape = ""
    ExtractPointInfo = info
End Function

' ===== チャートID生成 =====
Private Function GetChartPattern(ByRef logSheet As Worksheet) As String
    ' 初期値（エラー時のフォールバック用）
    Dim suffixPattern As String: suffixPattern = "-XX"
    Dim seriesPrefix As String: seriesPrefix = "Sample"
    
    On Error Resume Next
    
    ' シートから値を取得
    If Not logSheet Is Nothing Then
        ' サフィックスパターンの取得
        If Len(Trim(logSheet.Cells(2, "Q").value)) > 0 Then
            suffixPattern = Trim(logSheet.Cells(2, "Q").value)
            If Left(suffixPattern, 1) <> "-" Then
                suffixPattern = "-" & suffixPattern
            End If
        End If
        
        ' シリーズプレフィックスの取得
        If Len(Trim(logSheet.Cells(2, "D").value)) > 0 Then
            seriesPrefix = Trim(logSheet.Cells(2, "D").value)
        End If
    End If
    
    ' パターン文字列を生成
    GetChartPattern = "{0}-" & seriesPrefix & "-{1}-{2}-{3}" & suffixPattern
End Function

' ===== チャートID生成（修正版） =====
Private Function GenerateChartId(ByRef logSheet As Worksheet, _
                               ByRef sampleInfo As sampleInfo, _
                               ByRef pointInfo As pointInfo) As String
    On Error GoTo ErrorHandler
    
    If Len(sampleInfo.SampleNumber) = 0 Or Len(sampleInfo.Condition) = 0 _
       Or Len(pointInfo.Position) = 0 Or Len(pointInfo.Shape) = 0 Then
        GenerateChartId = ""
        Exit Function
    End If
    
    ' パターンを動的に取得
    Dim pattern As String
    pattern = GetChartPattern(logSheet)
    
    ' IDを生成
    GenerateChartId = Replace(pattern, "{0}", sampleInfo.SampleNumber)
    GenerateChartId = Replace(GenerateChartId, "{1}", pointInfo.Position)
    GenerateChartId = Replace(GenerateChartId, "{2}", sampleInfo.Condition)
    GenerateChartId = Replace(GenerateChartId, "{3}", pointInfo.Shape)
    Exit Function
    
ErrorHandler:
    GenerateChartId = ""
End Function

' ===== チャートコピー処理 =====
Private Sub CopyChart(ByRef sourceSheet As Worksheet, _
                     ByRef targetSheet As Worksheet, _
                     ByVal chartId As String, _
                     ByRef targetCell As Range, _
                     ByRef errorLogs As String)
    On Error GoTo ErrorHandler
    
    Dim cht As ChartObject
    Dim foundCharts As Long
    Const WAIT_TIME As String = "0:00:02.20"  ' 待機時間を2.20秒に延長
    Dim retryCount As Integer
    Const MAX_RETRY As Integer = 2  ' リトライ回数
    
    foundCharts = 0
    Debug.Print "検索ID: " & chartId
    
    For Each cht In sourceSheet.ChartObjects
        Debug.Print "チェック中のチャート: " & cht.Name
        
        If cht.Name = chartId Then
            foundCharts = foundCharts + 1
            If foundCharts = 1 Then
                ' オリジナルのサイズを保存
                Dim originalWidth As Double
                Dim originalHeight As Double
                originalWidth = cht.Width
                originalHeight = cht.Height
                
                ' コピー処理（リトライ付き）
                For retryCount = 0 To MAX_RETRY
                    ' クリップボードをクリア
                    Application.CutCopyMode = False
                    DoEvents
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' チャートをコピー
                    cht.Copy
                    DoEvents
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' ペースト実行
                    targetSheet.Paste Destination:=targetCell.Offset(0, 3)
                    DoEvents
                    
                    ' 新しいチャートのサイズを調整
                    Dim newChart As ChartObject
                    Set newChart = targetSheet.ChartObjects(targetSheet.ChartObjects.count)
                    
                    ' サイズを元のチャートに合わせる
                    newChart.Width = originalWidth
                    newChart.Height = originalHeight
                    
                    ' 完了後に待機
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' クリップボードをクリア
                    Application.CutCopyMode = False
                    
                    ' 成功確認
                    If Not newChart Is Nothing Then Exit For
                    
                    ' リトライ時のログ
                    If retryCount < MAX_RETRY Then
                        Debug.Print "コピー再試行: " & (retryCount + 1) & " 回目"
                    End If
                Next retryCount
            Else
                errorLogs = errorLogs & "重複チャート - シート: " & targetSheet.Name & _
                           ", ID: " & chartId & " (" & foundCharts & "個目)" & vbCrLf
            End If
        End If
    Next cht
    
    If foundCharts = 0 Then
        errorLogs = errorLogs & "未発見チャート - シート: " & targetSheet.Name & _
                   ", ID: " & chartId & vbCrLf
    End If
    
    Exit Sub

ErrorHandler:
    errorLogs = errorLogs & "エラー発生 - シート: " & targetSheet.Name & _
               ", ID: " & chartId & ", エラー: " & Err.Description & vbCrLf
    Resume Next
End Sub



' ===== ユーティリティ関数 =====
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function

