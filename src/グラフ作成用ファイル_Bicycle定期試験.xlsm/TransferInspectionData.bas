Attribute VB_Name = "TransferInspectionData"
Option Explicit

' 試料情報を格納する型定義
Private Type TestSample
    Number As String
    condition As String
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

'*******************************************************************************
' 型定義と定数
' 試料および測定点の情報を格納する構造体と状態を示す定数の定義
'*******************************************************************************

'*******************************************************************************
' 文字列変換用の辞書を初期化
' 機能：位置と形状の日本語表記を省略形に変換するための辞書を作成
' 戻値：Dictionaryオブジェクト（前頭部→前、後頭部→後、など）
'*******************************************************************************
Private Function InitializeConversionDicts()
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

'*******************************************************************************
' メインプロシージャ
' 概要：LOG_Bicycleシートから各製品シートへ試験データを転記
' 対象：500S_1, 500S_2, 500S_3シートのデータ転記
' 依存：InitializeConversionDicts, ProcessProductSheet
'*******************************************************************************
Sub TransferBicycleTestData()
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

'*******************************************************************************
' TransferBicycleTestDataのサブプロシージャ
' 機能：個別シートの試験データを検出し、転記処理を実行
' 引数：productSheet - 転記先の製品シート
'       logSheet     - データ元のLOGシート
'       convDict     - 変換用辞書オブジェクト
' 処理：試料情報の検出と測定点データの転記を実行
'*******************************************************************************
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

'*******************************************************************************
' ProcessProductSheetのサブプロシージャ
' 機能：セルの文字列から試料番号と試験条件を抽出
' 引数：cellValue - "試料1 高温" 形式の文字列
' 戻値：TestSample型（Number = "01", Condition = "Hot" など）
'*******************************************************************************
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
            sample.condition = HIGH_TEMP
        Case "低温"
            sample.condition = LOW_TEMP
        Case "浸せき"
            sample.condition = WET_CONDITION
    End Select
    
    ' デバッグ出力
    Debug.Print cellValue & " → " & sample.Number & ", " & sample.condition
    
    GetSampleInfo = sample
End Function

'*******************************************************************************
' ProcessProductSheetのサブプロシージャ
' 機能：検出された測定点のデータをLOGシートから製品シートへ転記
' 引数：productSheet - 転記先シート
'       logSheet     - LOG元シート
'       currentRow   - 処理中の行番号
'       sample       - 試料情報
'       convDict     - 変換辞書
' 特記：既転記データのスキップ処理とログ出力を含む
'*******************************************************************************
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
                        sample.condition & "-" & point.Shape

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

'*******************************************************************************
' ProcessMeasurementPointsのサブプロシージャ
' 機能：シートのセルから測定点の位置と形状を抽出
' 引数：sheet    - 対象シート
'       Row      - 対象行
'       Col      - 対象列
'       convDict - 変換辞書
' 戻値：MeasurementPoint型（Position = "前", Shape = "平" など）
' 特記：マージセルに対応
'*******************************************************************************
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


