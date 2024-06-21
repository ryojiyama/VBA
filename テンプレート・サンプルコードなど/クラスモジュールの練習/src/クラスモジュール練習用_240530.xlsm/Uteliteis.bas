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

    Application.DisplayAlerts = False ' �V�[�g�폜�̊m�F�_�C�A���O���\���ɂ���

    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        ' �I�u�W�F�N�g���� 'Temp_' ���܂܂�Ă��邩�m�F
        If InStr(1, ws.CodeName, "Temp", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True ' �m�F�_�C�A���O���ĕ\������
End Sub







