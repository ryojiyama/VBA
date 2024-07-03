Attribute VB_Name = "Utliteis"


' �]�L����15�s�ڈȉ����폜����B
Sub ClearTransferData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim sheetNames As Variant
    Dim i As Long
    
    ' �N���A����V�[�g�������X�g��
    sheetNames = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    ' �e�V�[�g�����[�v����
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        ' �V�[�g�����݂���ꍇ�A�f�[�^���N���A
        If Not ws Is Nothing Then
            startRow = 16 ' �f�[�^�̊J�n�s�i�w�b�_�[�̎��̍s�j
            lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
            
            ' �f�[�^�͈͂����݂���ꍇ�A�N���A
            If lastRow >= startRow Then
                ws.Range("B" & startRow & ":Z" & lastRow).ClearContents
            End If
            
            ' ws�����Z�b�g
            Set ws = Nothing
        End If
    Next i
End Sub

