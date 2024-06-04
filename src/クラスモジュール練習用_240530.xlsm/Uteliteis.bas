Attribute VB_Name = "Uteliteis"
Public Sub ListSheetCustomProperties()
    Dim sheet As Worksheet
    Dim prop As CustomProperty
    Dim output As String
    
    For Each sheet In ThisWorkbook.Sheets
        output = "Sheet Name: " & sheet.Name & vbCrLf
        For Each prop In sheet.CustomProperties
            output = output & "    Property Name: " & prop.Name & ", Value: " & prop.Value & vbCrLf
        Next prop
        Debug.Print output
    Next sheet
End Sub

' 複製したシートを削除するプロシージャ
Sub DeleteSheetsWithTempProperties()
    Dim ws As Worksheet
    Dim i As Integer

    Application.DisplayAlerts = False ' シート削除の確認ダイアログを非表示にする

    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        On Error Resume Next
        ' カスタムプロパティに 'Temp_' が含まれているか確認
        If ws.CustomProperties.Count > 0 Then
            Dim cp As CustomProperty
            For Each cp In ws.CustomProperties
                If InStr(cp.Name, "Temp_") > 0 Then
                    ws.Delete
                    Exit For
                End If
            Next cp
        End If
        On Error GoTo 0
    Next i

    Application.DisplayAlerts = True ' 確認ダイアログを再表示する
End Sub

