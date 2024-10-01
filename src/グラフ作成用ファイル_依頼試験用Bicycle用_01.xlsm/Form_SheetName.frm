VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SheetName 
   Caption         =   "Template Form"
   ClientHeight    =   5136
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   5724
   OleObjectBlob   =   "Form_SheetName.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form_SheetName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox_Type_Change()

End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet

    ' コンボボックスに "Hel_SpecSheet" を含むシート名を追加
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Hel_SpecSheet", vbTextCompare) > 0 Then
            ComboBox_Type.AddItem ws.Name
        End If
    Next ws

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



