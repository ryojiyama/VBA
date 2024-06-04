VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Helmet 
   Caption         =   "Template Form"
   ClientHeight    =   11892
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9084.001
   OleObjectBlob   =   "Form_Helmet.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form_Helmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 最終更新日：2021/05/05


'**************************
'※　依存関係(以下のモジュールが無いと最低限の動作出来ないです)
' クラスモジュール
'  - clsFormAssistant.cls
'  - clsWinApi.cls
'
' clsColorPalette.cls はユーザーコンテンツですので、不要であれば消していただいても構いません。
'**************************

Private clsAssist As New clsFormAssistant
Private palette As New Collection


Private Sub CalenderButton1_Click()
    DateLabel_BoutaiLot.Caption = CalendarForm.ShowCalender(Date, , clsAssist.ThemeColor)
End Sub

Private Sub CalenderButton2_Click()
    DateLabel_NaisouLot.Caption = CalendarForm.ShowCalender(Date, , clsAssist.ThemeColor)
End Sub

Private Sub CloseButton_Click()

End Sub

Private Sub DateLabel_NaisouLot_Click()

End Sub

Private Sub Header_Click()

End Sub

Private Sub Label53_Click()

End Sub

Private Sub Label58_Click()

End Sub

Private Sub Label61_Click()

End Sub

Private Sub Label62_Click()

End Sub

Private Sub Label65_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox_Ondo_Change()

End Sub

'**********************************
'user form
'**********************************
Private Sub UserForm_Initialize()

    clsAssist = Me
    clsAssist.ThemeColor = ToyoBlue   ' お好きな初期色を設定
    clsAssist.Version = "2.0"

    Call SetCalender
    
End Sub

Private Sub UserForm_Terminate()
    Set clsAssist = Nothing
End Sub

'**********************************
'出力ボタンを押した時に走る処理
'**********************************
Public Sub ClickRunButton()
    Debug.Print "run"
End Sub

'**********************************
'ユーザーコンテンツ
'**********************************
Private Sub SetPalette()     ' カラーパレットの色選択でフォームカラーを変更します。

    
End Sub

Public Sub Makeup(n As Integer)
    clsAssist.ThemeColor = n
End Sub

Private Sub SetCalender()
'    CalenderButton.Caption = clsAssist.GetCharactor(ICalender)
End Sub

Private Sub CalenderButton_Click()
    DateLabel.Caption = CalendarForm.ShowCalender(Date, , clsAssist.ThemeColor)
End Sub

Private Sub CalenderButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsAssist.ChangeCursor(Hand)
End Sub







