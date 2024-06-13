Attribute VB_Name = "Uteliteis"
Public Sub ListSheetCustomProperties()
    Dim sheet As Worksheet
    Dim prop As CustomProperty
    Dim output As String
    
    For Each sheet In ThisWorkbook.Sheets
        output = "Sheet Name: " & sheet.name & vbCrLf
        For Each prop In sheet.CustomProperties
            output = output & "    Property Name: " & prop.name & ", Value: " & prop.Value & vbCrLf
        Next prop
        Debug.Print output
    Next sheet
End Sub

Sub DeleteSheetsWithTempProperties()
    Dim ws As Worksheet
    Dim i As Integer

    Application.DisplayAlerts = False ' シート削除の確認ダイアログを非表示にする

    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        ' オブジェクト名に 'Temp_' が含まれているか確認
        If InStr(1, ws.CodeName, "Temp", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True ' 確認ダイアログを再表示する
End Sub







