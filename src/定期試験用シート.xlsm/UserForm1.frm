VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4212
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   3012
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Unload Me
End Sub
'
'Private Sub CommandButton2_Click()
'    Select Case ListBox1.value
'        Case "保護帽"
'            Call HelmetCreate.CreateGraphHelmet
'            Call HelmetCreate.InspectHelmetDurationTime
'            MsgBox "保護帽のグラフが完成しました。", vbInformation, "操作完了"
'        Case "自転車帽"
'            Call BicycleCreate.CreateGraphBicycle
'            Call BicycleCreate.Bicycle_150G_DurationTime
'            MsgBox "自転車帽のグラフが完了しました。", vbInformation, "操作完了"
'        Case "野球帽"
'            Call BaseBallCreate.CreateGraphBaseBall
'            Call BaseBallCreate.BaseBall_5kN7kN_DurationTime
'            MsgBox "野球帽のグラフが完了しました。", vbInformation, "操作完了"
'        Case "墜落制止用器具"
'            Call FallArrestCreate.CreateGraphFallArrest
'            Call FallArrest_2kN_DurationTime
'            MsgBox "墜落制止用器具のグラフが完了しました。", vbInformation, "操作完了"
'    End Select
'    Unload Me
'End Sub




Private Sub UserForm_Initialize()
With UserForm1.Controls("ListBox1")
.AddItem "保護帽"
.AddItem "自転車帽"
.AddItem "野球帽"
.AddItem "墜落制止用器具"
End With
End Sub

