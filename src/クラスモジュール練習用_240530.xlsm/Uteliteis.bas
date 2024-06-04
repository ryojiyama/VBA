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

' ���������V�[�g���폜����v���V�[�W��
Sub DeleteSheetsWithTempProperties()
    Dim ws As Worksheet
    Dim i As Integer

    Application.DisplayAlerts = False ' �V�[�g�폜�̊m�F�_�C�A���O���\���ɂ���

    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        On Error Resume Next
        ' �J�X�^���v���p�e�B�� 'Temp_' ���܂܂�Ă��邩�m�F
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

    Application.DisplayAlerts = True ' �m�F�_�C�A���O���ĕ\������
End Sub

