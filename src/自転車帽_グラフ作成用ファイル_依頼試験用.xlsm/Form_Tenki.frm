VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Tenki 
   Caption         =   "Template Form"
   ClientHeight    =   7860
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   5736
   OleObjectBlob   =   "Form_Tenki.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form_Tenki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private selectedButton As String

Private Sub UserForm_Initialize()
    selectedButton = ""
End Sub

Private Sub Button_Teiki_Click()
    selectedButton = "定期"
End Sub

Private Sub Button_Katashiki_Click()
    selectedButton = "型式"
End Sub

Private Sub Button_Irai_Click()
    selectedButton = "依頼"
End Sub

Private Sub RunButton_Click()
    If selectedButton = "" Then
        MsgBox "ボタンを選択してください。", vbExclamation
        Exit Sub
    End If

    CopySheetsToOtherWorkbooks selectedButton
End Sub









