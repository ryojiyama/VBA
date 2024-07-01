Attribute VB_Name = "Test"
Sub TransferDataToTopImpactTest()
    '天頂試験のみのシートを作成する。
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
    Set wsSource = ThisWorkbook.Sheets("Log_Helmet")

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 2行目から最終行までループ
    For i = 2 To lastRow
        ' C列の値から製品コードを取得
        firstDashPos = InStr(wsSource.Cells(i, 3).value, "-")
        If firstDashPos > 0 Then
            secondDashPos = InStr(firstDashPos + 1, wsSource.Cells(i, 3).value, "-")
            If secondDashPos > 0 Then
                matchName = Left(wsSource.Cells(i, 3).value, secondDashPos - 1)
            End If
        End If

        ' 各シートをループして条件に一致するシートを検索
        For Each wsDestination In ThisWorkbook.Sheets
            If wsDestination.name = matchName Then ' シート名が製品コードに一致するか確認
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


Sub TransferDataToDynamicSheets()
    ' 帽体の試験結果を対応するシートに転記する。
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim destinationSheetName As String

    ' ソースシートの設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row
    
    ' Excelのパフォーマンス向上のための設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSourceのC列をループしてデータを処理
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).value
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
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

