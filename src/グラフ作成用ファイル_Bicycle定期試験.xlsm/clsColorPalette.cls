VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private clsApi As New clsWinApi
Private WithEvents colorLabel As MSForms.label
Attribute colorLabel.VB_VarHelpID = -1
Private uf As Object
Private number As Integer

Private Sub Class_Terminate()
    Set clsApi = Nothing
End Sub

Property Let LabelObject(l As Object)
Attribute LabelObject.VB_UserMemId = 0
    Set colorLabel = l
    Set uf = l.Parent
    number = CInt(Replace(l.name, "color", ""))
End Property

Private Sub colorLabel_Click()
    Call uf.Makeup(number)
End Sub

Private Sub colorLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsApi.ChangeCursor(Hand)
End Sub










