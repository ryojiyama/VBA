VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4212
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   3012
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'ƒI[ƒi[ ƒtƒH[ƒ€‚Ì’†‰›
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
        Case "•ÛŒì–X"
            Call HelmetCreate.CreateGraphHelmet
            Call HelmetCreate.InspectHelmetDurationTime
            MsgBox "•ÛŒì–X‚ÌƒOƒ‰ƒt‚ªŠ®¬‚µ‚Ü‚µ‚½B", vbInformation, "‘€ìŠ®—¹"
        Case "©“]Ô–X"
            Call BicycleCreate.CreateGraphBicycle
            Call BicycleCreate.Bicycle_150G_DurationTime
            MsgBox "©“]Ô–X‚ÌƒOƒ‰ƒt‚ªŠ®—¹‚µ‚Ü‚µ‚½B", vbInformation, "‘€ìŠ®—¹"
        Case "–ì‹…–X"
            Call BaseBallCreate.CreateGraphBaseBall
            Call BaseBallCreate.BaseBall_5kN7kN_DurationTime
            MsgBox "–ì‹…–X‚ÌƒOƒ‰ƒt‚ªŠ®—¹‚µ‚Ü‚µ‚½B", vbInformation, "‘€ìŠ®—¹"
        Case "’Ä—§~—pŠí‹ï"
            Call FallArrestCreate.CreateGraphFallArrest
            Call FallArrest_2kN_DurationTime
            MsgBox "’Ä—§~—pŠí‹ï‚ÌƒOƒ‰ƒt‚ªŠ®—¹‚µ‚Ü‚µ‚½B", vbInformation, "‘€ìŠ®—¹"
    End Select
    Unload Me
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Select Case ListBox1.Value
        Case "•ÛŒì–X"
            Call Procedure1
        Case "©“]Ô–X"
            Call Procedure2
        Case "–ì‹…–X"
            Call Procedure3
        Case "’Ä—§~—pŠí‹ï"
            Call Procedure4
    End Select
    Unload Me
End Sub


Private Sub UserForm_Initialize()
With UserForm1.Controls("ListBox1")
.AddItem "•ÛŒì–X"
.AddItem "©“]Ô–X"
.AddItem "–ì‹…–X"
.AddItem "’Ä—§~—pŠí‹ï"
End With
End Sub

