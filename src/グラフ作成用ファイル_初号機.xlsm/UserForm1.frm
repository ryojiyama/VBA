VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4212
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   3012
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Select Case ListBox1.Value
        Case "�ی�X"
            Call HelmetCreate.CreateGraphHelmet
            Call HelmetCreate.InspectHelmetDurationTime
            MsgBox "�ی�X�̃O���t���������܂����B", vbInformation, "���슮��"
        Case "���]�ԖX"
            Call BicycleCreate.CreateGraphBicycle
            Call BicycleCreate.Bicycle_150G_DurationTime
            MsgBox "���]�ԖX�̃O���t���������܂����B", vbInformation, "���슮��"
        Case "�싅�X"
            Call BaseBallCreate.CreateGraphBaseBall
            Call BaseBallCreate.BaseBall_5kN7kN_DurationTime
            MsgBox "�싅�X�̃O���t���������܂����B", vbInformation, "���슮��"
        Case "�ė����~�p���"
            Call FallArrestCreate.CreateGraphFallArrest
            Call FallArrest_2kN_DurationTime
            MsgBox "�ė����~�p���̃O���t���������܂����B", vbInformation, "���슮��"
    End Select
    Unload Me
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case ListBox1.Value
        Case "�ی�X"
            Call Procedure1
        Case "���]�ԖX"
            Call Procedure2
        Case "�싅�X"
            Call Procedure3
        Case "�ė����~�p���"
            Call Procedure4
    End Select
    Unload Me
End Sub


Private Sub UserForm_Initialize()
With UserForm1.Controls("ListBox1")
.AddItem "�ی�X"
.AddItem "���]�ԖX"
.AddItem "�싅�X"
.AddItem "�ė����~�p���"
End With
End Sub

