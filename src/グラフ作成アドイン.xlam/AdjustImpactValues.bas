Attribute VB_Name = "AdjustImpactValues"
'*******************************************************************************
' メインプロシージャ
' 機能：LOG_シートの衝撃値を製品種別に応じて調整し、フォーマットを設定
' 引数：なし
' 補足：各製品の衝撃値に行番号に応じた微小値を加算して重複を防ぐ
'*******************************************************************************
Sub AdjustImpactValuesWithCustomFormatForAllLOGSheets()
    Dim ws As Worksheet
    Dim impactCol As Long
    Dim impactValue As Double
    Dim rowNum As Long
    Dim lastRow As Long
    Dim i As Long
    Dim backupCol As Long
    Dim adjustmentFactor As Double
    Dim displayFormat As String

    ' バックアップを保存する列（X列＝24列目）
    backupCol = 24

    ' すべてのシートを順に探索
    For Each ws In ActiveWorkbook.Sheets
        ' シート名に"LOG_"を含むシートに対してのみ処理を行う
        If InStr(ws.Name, "LOG_") > 0 Then

            ' ヘッダー行から"最大値("を含む列を見つける
            For i = 1 To ws.Cells(1, ws.columns.Count).End(xlToLeft).column
                If InStr(ws.Cells(1, i).value, "最大値(") > 0 Then
                    impactCol = i
                    Exit For
                End If
            Next i

            ' "最大値"列が見つからなかった場合は次のシートへ
            If impactCol = 0 Then
                MsgBox ws.Name & " シートに最大値列が見つかりません。"
                GoTo NextSheet
            End If

            ' 最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, impactCol).End(xlUp).row

            ' シート名に応じて調整係数を設定
            Select Case ws.Name
                Case "LOG_Helmet", "LOG_FallArrest"
                    adjustmentFactor = 0.000001
                    displayFormat = "0.000000" ' 小数点以下6桁まで表示
                Case "LOG_BaseBall", "LOG_Bicycle"
                    adjustmentFactor = 0.01
                    displayFormat = "0.00" ' 小数点以下2桁まで表示
                Case Else
                    MsgBox ws.Name & " シート名が適切ではありません。処理をスキップします。"
                    GoTo NextSheet
            End Select

            ' 2行目以降のセルに対して処理を行う
            For rowNum = 2 To lastRow
                impactValue = ws.Cells(rowNum, impactCol).value

                ' 元の impactValue を X列にバックアップ
                ws.Cells(rowNum, backupCol).value = impactValue

                ' 計算式を適用
                impactValue = impactValue + (rowNum * adjustmentFactor)

                ' 計算結果を元の列に代入
                ws.Cells(rowNum, impactCol).value = impactValue

                ' セルの表示形式を設定
                ws.Cells(rowNum, impactCol).NumberFormat = displayFormat
            Next rowNum

NextSheet:
            impactCol = 0 ' 次のシートのためにリセット

        End If
    Next ws
    Call HighlightDuplicateValues
End Sub
'*******************************************************************************
' サブプロシージャ
' 機能：衝撃値の重複をチェックし、重複値に色付けを行う
' 引数：なし
' 補足：製品種別ごとに対象列を変えて重複チェックを実施
'*******************************************************************************
Sub HighlightDuplicateValues()
    ' 対象シート名のリスト
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    
    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    Dim sheetName As Variant
    
    ' シートごとに処理
    For Each sheetName In sheetNames
        ' シートオブジェクトを設定（エラーハンドリングを追加）
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' シートが存在する場合のみ処理を実行
        If Not ws Is Nothing Then
            ' シートに応じて対象列を設定
            Dim targetColumn As String
            Select Case CStr(sheetName)
                Case "LOG_Helmet"
                    targetColumn = "H"  ' ヘルメットのログは H 列
                Case "LOG_FallArrest"
                    targetColumn = "H"  ' 墜落制止用器具のログは I 列
                Case "LOG_Bicycle"
                    targetColumn = "J"  ' 自転車のログは J 列
                Case "LOG_BaseBall"
                    targetColumn = "H"  ' 野球用具のログは K 列
            End Select
            
            ' 最終行を取得
            lastRow = ws.Cells(ws.Rows.Count, targetColumn).End(xlUp).row
            
            ' 対象範囲の色をクリア
            ws.Range(targetColumn & "2:" & targetColumn & lastRow).Interior.colorIndex = xlNone
            
            ' 色のインデックスを初期化
            colorIndex = 3 ' Excelの色インデックスは3から始まる
            
            ' シートごとの重複チェック
            For i = 2 To lastRow
                ' 現在のセルの値を取得
                valueToFind = ws.Cells(i, targetColumn).value
                
                ' 値が空でないことを確認
                If Not IsEmpty(valueToFind) Then
                    ' 同じ値を持つセルが既に色付けされていないかチェック
                    If ws.Cells(i, targetColumn).Interior.colorIndex = xlNone Then
                        Dim duplicateFound As Boolean
                        duplicateFound = False
                        
                        For j = i + 1 To lastRow
                            ' シートごとの重複チェックロジック
                            Select Case CStr(sheetName)
                                Case "LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall"
                                    ' すべて完全一致でチェック
                                    If ws.Cells(j, targetColumn).value = valueToFind Then
                                        duplicateFound = True
                                    End If
                            End Select
                            
                            ' 重複が見つかった場合、色を付ける
                            If duplicateFound Then
                                ws.Cells(i, targetColumn).Interior.colorIndex = colorIndex
                                ws.Cells(j, targetColumn).Interior.colorIndex = colorIndex
                                duplicateFound = False  ' 次のチェックのためにリセット
                            End If
                        Next j
                        
                        ' 色インデックスを更新
                        colorIndex = colorIndex + 1
                        If colorIndex > 56 Then colorIndex = 3
                    End If
                End If
            Next i
            
            ' オブジェクトのクリア
            Set ws = Nothing
        Else
            ' シートが見つからない場合のデバッグ出力
            Debug.Print "シートが見つかりませんでした: " & sheetName
        End If
    Next sheetName
End Sub


