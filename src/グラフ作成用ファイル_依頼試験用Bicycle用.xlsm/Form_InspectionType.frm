VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_InspectionType 
   Caption         =   "Template Form"
   ClientHeight    =   5136
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   5724
   OleObjectBlob   =   "Form_InspectionType.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form_InspectionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    With ComboBox_Type
        .AddItem "定期試験用"
        .AddItem "型式申請試験用"
        .AddItem "その他"
    End With
End Sub


Private Sub RunButton_Click()
    Dim selectedType As String
    selectedType = ComboBox_Type.value
    
    If selectedType = "" Then
        MsgBox "グラフの種類を選択してください。", vbExclamation
        Exit Sub
    End If
    
    CreateGraphHelmet selectedType
    Call InspectHelmetDurationTime
    Call Utlities.AdjustingDuplicateValues
End Sub






'定期試験用、型式申請試験用、その他

