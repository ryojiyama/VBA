VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SpecSheetStylerHelmet 
   Caption         =   "Template Form"
   ClientHeight    =   5136
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   5724
   OleObjectBlob   =   "Form_SpecSheetStylerHelmet.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "Form_SpecSheetStylerHelmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub RunButton_Click()
    Dim sheetName As String
    sheetName = ComboBox_Type.value ' ComboBox_Type�ɑI�����ꂽ�V�[�g�����i�[

    ' �V�[�g�����I������Ă���ꍇ��CreateID�����s
    If sheetName <> "" Then
        Call CreateID(sheetName)
        Call TransferValuesBetweenSheets
        Call SyncSpecSheetToLogHel
    Else
        MsgBox "�V�[�g�����I������Ă��܂���B", vbExclamation
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet

    ' �R���{�{�b�N�X�� "Hel_SpecSheet" ���܂ރV�[�g����ǉ�
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(1, ws.Name, "SpecSheet", vbTextCompare) > 0 Then
            ComboBox_Type.AddItem ws.Name
        End If
    Next ws

End Sub




