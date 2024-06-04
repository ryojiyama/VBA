VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "date & picker"
   ClientHeight    =   6024
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6468
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 最終更新日：2021/05/05

'**************************
'モジュール側の処理はこんな感じ
'res = Format(CalendarForm.ShowCalender(Date, "title"), "yyyy/mm/dd")
'**************************

Private the_date As Date
Private clndr_date As Date
Private color As Long

Public Function ShowCalender(color_date As Date, Optional title As String = "", Optional theme_color As Long = 0)

    If title <> "" Then Me.Caption = title
    the_date = color_date
    color = theme_color
    Call UserForm_Initialize
    Me.Show
    ShowCalender = clndr_date

End Function

Private Sub UserForm_Initialize() 'Formが開くとき
    Dim i As Integer

    If the_date = 0 Then Exit Sub
    For i = -3 To 3 '前後3年分の年を登録
      Me.ComboBox1.AddItem CStr((Year(the_date)) + i)
    Next i
    For i = 1 To 12 '月を登録
      Me.ComboBox2.AddItem CStr(i)
    Next i
     
    Me.ComboBox1 = Year(the_date) '年を指定
    Me.ComboBox2 = Month(the_date) '月を指定
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    For i = 1 To 42 'グレーのとこだけ初期化
        If Me("Label" & i).BackColor = RGB(221, 222, 211) Then Me("Label" & i).BackColor = Me.BackColor
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'Formが閉じるとき
  If CloseMode = 0 Then '×ボタンを押された場合
    clndr_date = the_date 'テキストボックスと同じ値を入れておく
  End If
End Sub
 
Private Sub ComboBox1_Change() '年が変更されたとき
  Call clndr_set
End Sub
 
Private Sub ComboBox2_Change() '月が変更されたとき
  Call clndr_set
End Sub
 
Private Sub clndr_set() 'カレンダーの作成と表示
  Dim yy As Integer, mm As Integer, i As Integer, n As Integer, endDay As Integer
   
  If Me.ComboBox1 = "" Or Me.ComboBox2 = "" Then Exit Sub '年か月どちらか入ってなければ中止
  yy = Me.ComboBox1 '年
  mm = Me.ComboBox2 '月
    
  For i = 1 To 42 'ラベルの初期化
    Me("Label" & i).Caption = ""
    Me("Label" & i).ForeColor = RGB(0, 0, 0)
    Me("Label" & i).BackColor = Me.BackColor
  Next
   
  n = Weekday(yy & "/" & mm & "/" & 1) - 1 'その月の1日の曜日番号に、マイナス1したもの
  endDay = Day(DateAdd("d", -1, DateAdd("m", 1, yy & "/" & mm & "/" & "1"))) '月末日の算出
  For i = 1 To endDay
    Me("Label" & i + n).Caption = i '日を入れる
    If CDate(yy & "/" & mm & "/" & i) = the_date Then       'TextBoxの日と同じなら色をつける
        Me("Label" & i + n).ForeColor = RGB(255, 255, 255)
        Me("Label" & i + n).BackColor = color
    End If
  Next i
End Sub

Private Sub LblClk(ByVal i As Integer)
  If Me("Label" & i).Caption = "" Then Exit Sub 'ラベルが空だったら中止
  clndr_date = Me.ComboBox1 & "/" & Me.ComboBox2 & "/" & Me("Label" & i).Caption '日付を生成して変数に格納
  Unload Me 'カレンダーを閉じる
End Sub

Private Sub SpinButton1_SpinUp()
  If Me.ComboBox2 + 1 > 12 Then
    Me.ComboBox1 = Me.ComboBox1 + 1
    Me.ComboBox2 = 1
  Else
    Me.ComboBox2 = Me.ComboBox2 + 1
  End If
End Sub
 
Private Sub SpinButton1_SpinDown()
  If Me.ComboBox2 - 1 < 1 Then
    Me.ComboBox1 = Me.ComboBox1 - 1
    Me.ComboBox2 = 12
  Else
    Me.ComboBox2 = Me.ComboBox2 - 1
  End If
End Sub
 
Private Sub Label1_Click()
  Call LblClk(1)
End Sub
 
Private Sub Label2_Click()
  Call LblClk(2)
End Sub
 
Private Sub Label3_Click()
  Call LblClk(3)
End Sub

Private Sub Label4_Click()
  Call LblClk(4)
End Sub

Private Sub Label5_Click()
  Call LblClk(5)
End Sub

Private Sub Label6_Click()
  Call LblClk(6)
End Sub

Private Sub Label7_Click()
  Call LblClk(7)
End Sub

Private Sub Label8_Click()
  Call LblClk(8)
End Sub

Private Sub Label9_Click()
  Call LblClk(9)
End Sub

Private Sub Label10_Click()
  Call LblClk(10)
End Sub

Private Sub Label11_Click()
  Call LblClk(11)
End Sub

Private Sub Label12_Click()
  Call LblClk(12)
End Sub

Private Sub Label13_Click()
  Call LblClk(13)
End Sub

Private Sub Label14_Click()
  Call LblClk(14)
End Sub

Private Sub Label15_Click()
  Call LblClk(15)
End Sub

Private Sub Label16_Click()
  Call LblClk(16)
End Sub

Private Sub Label17_Click()
  Call LblClk(17)
End Sub

Private Sub Label18_Click()
  Call LblClk(18)
End Sub

Private Sub Label19_Click()
  Call LblClk(19)
End Sub

Private Sub Label20_Click()
  Call LblClk(20)
End Sub

Private Sub Label21_Click()
  Call LblClk(21)
End Sub

Private Sub Label22_Click()
  Call LblClk(22)
End Sub

Private Sub Label23_Click()
  Call LblClk(23)
End Sub

Private Sub Label24_Click()
  Call LblClk(24)
End Sub

Private Sub Label25_Click()
  Call LblClk(25)
End Sub

Private Sub Label26_Click()
  Call LblClk(26)
End Sub

Private Sub Label27_Click()
  Call LblClk(27)
End Sub

Private Sub Label28_Click()
  Call LblClk(28)
End Sub

Private Sub Label29_Click()
  Call LblClk(29)
End Sub

Private Sub Label30_Click()
  Call LblClk(30)
End Sub

Private Sub Label31_Click()
  Call LblClk(31)
End Sub

Private Sub Label32_Click()
  Call LblClk(32)
End Sub

Private Sub Label33_Click()
  Call LblClk(33)
End Sub

Private Sub Label34_Click()
  Call LblClk(34)
End Sub

Private Sub Label35_Click()
  Call LblClk(35)
End Sub

Private Sub Label36_Click()
  Call LblClk(36)
End Sub

Private Sub Label37_Click()
  Call LblClk(37)
End Sub

Private Sub Label38_Click()
  Call LblClk(38)
End Sub

Private Sub Label39_Click()
  Call LblClk(39)
End Sub

Private Sub Label40_Click()
  Call LblClk(40)
End Sub

Private Sub Label41_Click()
  Call LblClk(41)
End Sub

Private Sub Label42_Click()
  Call LblClk(42)
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label1").Caption <> "" And Me("Label1").BackColor <> color Then Me("Label1").BackColor = RGB(221, 222, 211)     '選択色
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label2").Caption <> "" And Me("Label2").BackColor <> color Then Me("Label2").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label3").Caption <> "" And Me("Label3").BackColor <> color Then Me("Label3").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label4").Caption <> "" And Me("Label4").BackColor <> color Then Me("Label4").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label5").Caption <> "" And Me("Label5").BackColor <> color Then Me("Label5").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label6").Caption <> "" And Me("Label6").BackColor <> color Then Me("Label6").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label7").Caption <> "" And Me("Label7").BackColor <> color Then Me("Label7").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label8").Caption <> "" And Me("Label8").BackColor <> color Then Me("Label8").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label9").Caption <> "" And Me("Label9").BackColor <> color Then Me("Label9").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label10").Caption <> "" And Me("Label10").BackColor <> color Then Me("Label10").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label11").Caption <> "" And Me("Label11").BackColor <> color Then Me("Label11").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label12").Caption <> "" And Me("Label12").BackColor <> color Then Me("Label12").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label13").Caption <> "" And Me("Label13").BackColor <> color Then Me("Label13").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label14").Caption <> "" And Me("Label14").BackColor <> color Then Me("Label14").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label15").Caption <> "" And Me("Label15").BackColor <> color Then Me("Label15").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label16").Caption <> "" And Me("Label16").BackColor <> color Then Me("Label16").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label17").Caption <> "" And Me("Label17").BackColor <> color Then Me("Label17").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label18").Caption <> "" And Me("Label18").BackColor <> color Then Me("Label18").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label19").Caption <> "" And Me("Label19").BackColor <> color Then Me("Label19").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label20").Caption <> "" And Me("Label20").BackColor <> color Then Me("Label20").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label21").Caption <> "" And Me("Label21").BackColor <> color Then Me("Label21").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label22").Caption <> "" And Me("Label22").BackColor <> color Then Me("Label22").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label23").Caption <> "" And Me("Label23").BackColor <> color Then Me("Label23").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label24").Caption <> "" And Me("Label24").BackColor <> color Then Me("Label24").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label25").Caption <> "" And Me("Label25").BackColor <> color Then Me("Label25").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label26").Caption <> "" And Me("Label26").BackColor <> color Then Me("Label26").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label27").Caption <> "" And Me("Label27").BackColor <> color Then Me("Label27").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label28").Caption <> "" And Me("Label28").BackColor <> color Then Me("Label28").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label29").Caption <> "" And Me("Label29").BackColor <> color Then Me("Label29").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label30").Caption <> "" And Me("Label30").BackColor <> color Then Me("Label30").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label31").Caption <> "" And Me("Label31").BackColor <> color Then Me("Label31").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label32").Caption <> "" And Me("Label32").BackColor <> color Then Me("Label32").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label33").Caption <> "" And Me("Label33").BackColor <> color Then Me("Label33").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label34").Caption <> "" And Me("Label34").BackColor <> color Then Me("Label34").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label35").Caption <> "" And Me("Label35").BackColor <> color Then Me("Label35").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label36").Caption <> "" And Me("Label36").BackColor <> color Then Me("Label36").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label37").Caption <> "" And Me("Label37").BackColor <> color Then Me("Label37").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label38").Caption <> "" And Me("Label38").BackColor <> color Then Me("Label38").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label39").Caption <> "" And Me("Label39").BackColor <> color Then Me("Label39").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label40").Caption <> "" And Me("Label40").BackColor <> color Then Me("Label40").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label41").Caption <> "" And Me("Label41").BackColor <> color Then Me("Label41").BackColor = RGB(221, 222, 211)
End Sub

Private Sub Label42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me("Label42").Caption <> "" And Me("Label42").BackColor <> color Then Me("Label42").BackColor = RGB(221, 222, 211)
End Sub

