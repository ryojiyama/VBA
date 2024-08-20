Attribute VB_Name = "SheetHuriwake"
Sub TransferDataBasedOnID()
    Call Utlities.DeleteRowsBelowHeader

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim maxValue As Double, duration49kN As Double, duration73kN As Double
    Dim nextRow As Long
    Dim tempArray As Variant
    Dim data As Collection
    Dim dataItem As Variant
    
    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割
        idParts = Split(wsSource.Cells(i, 3).value, "-")
        If UBound(idParts) >= 2 Then
            ' グループ名（部位）を取得
            group = idParts(2)
            
            ' グループ名に基づいてシート名を設定
            Select Case group
                Case "天"
                    targetSheetName = "Impact_Top"
                Case "前"
                    targetSheetName = "Impact_Front"
                Case "後"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' 対応するグループがない場合はスキップ
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).value
                    GoTo NextIteration
            End Select
            
            groupName = "Group:" & idParts(0) & group
            maxValue = wsSource.Range("H" & i).value
            duration49kN = wsSource.Range("J" & i).value
            duration73kN = wsSource.Range("K" & i).value

            ' グループ名とシート名の対応を確認
'            Debug.Print "Group: " & groupName & "; Sheet: " & targetSheetName
'            Debug.Print "Max Value: " & Format(maxValue, "0.00") & " 49kN Duration: " & Format(duration49kN, "0.00") & " 73kN Duration: " & Format(duration73kN, "0.00")

            ' データをコレクションに追加
            tempArray = Array( _
            groupName, _
            targetSheetName, _
            Format(maxValue, "0.00"), _
            Format(duration49kN, "0.00"), _
            Format(duration73kN, "0.00") _
            )
            data.Add tempArray
        End If
NextIteration:
    Next i
    
    ' コレクションから各シートにデータを転記
    For Each dataItem In data
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)
        ' 目的のシートを作成
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsDest.name = targetSheetName
        End If
        On Error GoTo 0
        
        ' ヘッダー行を設定（14行目）
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "Max"
            wsDest.Range("C14").value = "4.9kN"
            wsDest.Range("D14").value = "7.3kN"
        End If
        nextRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        
        'データを転記
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = maxValue
        wsDest.Range("C" & nextRow).value = duration49kN
        wsDest.Range("D" & nextRow).value = duration73kN
    Next dataItem

    ' リソースを解放
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub



