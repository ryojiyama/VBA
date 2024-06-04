Attribute VB_Name = "InspectionSheet"

Sub TransferDataWithDistinctSheets()
    On Error GoTo ErrorHandler

    Dim wsLog As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long, destCol As Long, destRow As Long
    Dim sheetName As String, impact As String

    ' Initialize sheet
    Set wsLog = ThisWorkbook.Sheets("Log_Helmet")
    lastRow = wsLog.Cells(wsLog.Rows.Count, "B").End(xlUp).row

    ' Data processing for each row
    For i = 2 To lastRow
        sheetName = GetSheetName(wsLog.Cells(i, "E").Value)
        If sheetName <> "" Then
            Set wsDestination = ThisWorkbook.Sheets(sheetName)
            destCol = GetDestinationColumn(sheetName, wsLog.Cells(i, "L").Value)
            destRow = GetDestinationRow(sheetName, wsLog.Cells(i, "B").Value)

            ' Copy data if valid column and row
            If destCol <> 0 And destRow <> 0 Then
                ' Copy H column to the determined column
                CopyData wsLog, wsDestination, i, destRow, destCol, "H"

                ' Check and copy J only if it's numeric and >= 0.1
                If IsNumeric(wsLog.Cells(i, "J").Value) And wsLog.Cells(i, "J").Value >= 0.01 Then
                    Dim destColForJ As Long
                    ' Adjust destination column for J as needed, here using destCol + 1 as an example
                    destColForJ = destCol + 1
                    CopyData wsLog, wsDestination, i, destRow, destColForJ, "J"
                End If

                ' Check and copy K only if it's numeric and >= 0.1
                If IsNumeric(wsLog.Cells(i, "K").Value) And wsLog.Cells(i, "K").Value >= 0.01 Then
                    Dim destColForK As Long
                    ' Adjust destination column for K as needed, here using destCol + 2 as an example
                    destColForK = destCol + 2
                    CopyData wsLog, wsDestination, i, destRow, destColForK, "K"
                End If
            End If
        End If
    Next i

CleanUp:
    Set wsLog = Nothing
    Set wsDestination = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Function GetSheetName(impact As String) As String
    Select Case impact
        Case "天頂部": GetSheetName = "Impact_Top"
        Case "前頭部": GetSheetName = "Impact_Front"
        Case "後頭部": GetSheetName = "Impact_Back"
        Case "側頭部": GetSheetName = "Impact_Side"
        Case Else: GetSheetName = ""
    End Select
End Function

Function GetDestinationColumn(sheetName As String, condition As String) As Long
    Dim cols As Object
    Set cols = CreateObject("Scripting.Dictionary")

    ' Define column mapping for each condition and sheet
    cols("Impact_Top高温") = 3
    cols("Impact_Top低温") = 5
    cols("Impact_Top浸せき") = 7
    cols("Impact_Side高温") = 5
    cols("Impact_Side低温") = 6
    cols("Impact_Side浸せき") = 7
    ' Add other mappings as necessary

    GetDestinationColumn = cols(sheetName & condition)
End Function

Function GetDestinationRow(sheetName As String, refVal As String) As Long
    Dim lastDigit As String
    lastDigit = Right(refVal, 1)  ' 最後の文字を取得
    
    Select Case sheetName
        Case "Impact_Top"
            ' Impact_Topの場合の行の割り当て
            Select Case lastDigit
                Case "1", "4", "7"
                    GetDestinationRow = 6
                Case "2", "5", "8"
                    GetDestinationRow = 8
                Case "3", "6", "9"
                    GetDestinationRow = 10
                Case Else
                    GetDestinationRow = 0
            End Select
            
        Case "Impact_Side"
            ' Impact_Sideの場合の行の割り当て
            Select Case lastDigit
                Case "1", "4", "7"
                    GetDestinationRow = 7
                Case "2", "5", "8"
                    GetDestinationRow = 9
                Case "3", "6", "9"
                    GetDestinationRow = 11
                Case Else
                    GetDestinationRow = 0
            End Select
            
        Case Else
            ' 他のシートの場合のデフォルトの行の割り当て
            Select Case lastDigit
                Case "1", "4", "7"
                    GetDestinationRow = 6
                Case "2", "5", "8"
                    GetDestinationRow = 9
                Case "3", "6", "9"
                    GetDestinationRow = 12
                Case Else
                    GetDestinationRow = 0
            End Select
    End Select
End Function


Sub CopyData(wsSource As Worksheet, wsDest As Worksheet, sourceRow As Long, destRow As Long, destCol As Long, sourceCol As String)
    wsDest.Cells(destRow, destCol).Value = wsSource.Cells(sourceRow, sourceCol).Value
End Sub

Sub TransferDataWithDistinctSheets_Old()
    On Error GoTo ErrorHandler

    Dim wsLog As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String, destCol As Long, destRow As Long

    ' シートの初期設定
    Set wsLog = ThisWorkbook.Sheets("Log_Helmet")

    ' Log_Helmet シートの最終行を取得
    lastRow = wsLog.Cells(wsLog.Rows.Count, "B").End(xlUp).row

    ' データを行ごとに処理
    For i = 2 To lastRow
        ' E列に基づいたシート名の決定
        Select Case wsLog.Cells(i, "E").Value
            Case "天頂部"
                sheetName = "Impact_Top"
            Case "前頭部"
                sheetName = "Impact_Front"
            Case "後頭部"
                sheetName = "Impact_Back"
            Case "側頭部"
                sheetName = "Impact_Side"
            Case Else
                sheetName = ""
        End Select

        If sheetName <> "" Then
            Set wsDestination = ThisWorkbook.Sheets(sheetName)

            ' L列とシート構造に基づいた列の決定
            Select Case sheetName
                Case "Impact_Top", "Impact_Front", "Impact_Back"
                    Select Case wsLog.Cells(i, "L").Value
                        Case "高温"
                            destCol = 3 ' C列
                        Case "低温"
                            destCol = 5 ' E列
                        Case "浸せき"
                            destCol = 7 ' G列
                        Case Else
                            destCol = 0
                    End Select
                Case "Impact_Side"
                    Select Case wsLog.Cells(i, "L").Value
                        Case "高温"
                            destCol = 5 ' E列
                        Case "低温"
                            destCol = 6 ' F列
                        Case "浸せき"
                            destCol = 7 ' G列
                        Case Else
                            destCol = 0
                    End Select
            End Select

            ' B列の最後の文字に基づいた行の決定、シート構造に応じて
            If sheetName = "Impact_Top" Then
                Select Case Right(wsLog.Cells(i, "B").Value, 1)
                    Case "1"
                        destRow = 6
                    Case "2"
                        destRow = 8
                    Case "3"
                        destRow = 10
                    Case "4"
                        destRow = 6
                    Case "5"
                        destRow = 8
                    Case "6"
                        destRow = 10
                    Case "7"
                        destRow = 6
                    Case "8"
                        destRow = 8
                    Case "9"
                        destRow = 10
                    Case Else
                        destRow = 0
                End Select
            Else
                Select Case Right(wsLog.Cells(i, "B").Value, 1)
                    Case "1"
                        destRow = 6
                    Case "2"
                        destRow = 9
                    Case "3"
                        destRow = 12
                    Case "4"
                        destRow = 6
                    Case "5"
                        destRow = 9
                    Case "6"
                        destRow = 12
                    Case "7"
                        destRow = 6
                    Case "8"
                        destRow = 9
                    Case "9"
                        destRow = 12
                    Case Else
                        destRow = 0
                End Select
            End If
            If destCol <> 0 And destRow <> 0 Then
                ' 値を適切な位置に転記
                wsDestination.Cells(destRow, destCol).Value = wsLog.Cells(i, "H").Value
            End If
        End If
    Next i

CleanUp:
    ' リソースの解放
    Set wsLog = Nothing
    Set wsDestination = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub
