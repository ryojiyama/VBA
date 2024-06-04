VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' 最終更新日：2021/04/30

'**************************
'※　依存関係(以下のモジュールが無いと動作出来ないです)
' clsWinApi.cls
'**************************

'**********************************
'列挙型定義
'**********************************

Public Enum FormColors          ' GetColorで使用。
    WHITE = 0
    BLACK
    GRAY01
    GRAY02
    GRAY03
    GRAY04
    RED

    LightRED            ' = 7
    LightGREEN
    LightBLUE
    LightBROUN
    LightPURPLE

    DarkRED             ' = 12
    DarkGREEN
    DarkBLUE
    ToyoBlue

    ORANGE              ' = 16
    PEAR
    MINT

End Enum

Public Enum SpCharType          ' GetCharactorで使用。
    CheckBoxOff = 0
    CheckBoxOn = 1
    CheckBoxX = 2
    OptionOff = 3
    OptionOn = 4
    
    ISeting = 10                       'Win8,10用
    ICamera = 11
    ISave = 12
    IDelete = 13
    ICalender = 14
    IShere = 15
    IStart = 16
    IPause = 17
    IStop = 18

    ICherryBlossom = 19
    ICircle = 20
    ICopyRight = 21
End Enum

'**********************************
'ローカル変数
'**********************************
Private clsApi As New clsWinApi
Private color As Long
Private madeup As Boolean

Private uf As Object
Private WithEvents Header As MSForms.label
Attribute Header.VB_VarHelpID = -1
Private WithEvents Bottom As MSForms.label
Attribute Bottom.VB_VarHelpID = -1
Private WithEvents CloseButton As MSForms.label
Attribute CloseButton.VB_VarHelpID = -1
Private WithEvents RunButton As MSForms.CommandButton
Attribute RunButton.VB_VarHelpID = -1
Private WithEvents Ver As MSForms.label
Attribute Ver.VB_VarHelpID = -1

'**********************************
'クラスプロパティ
'**********************************

' [Let] デフォルトプロパティ
Public Property Let UserFormObject(o As Object)
Attribute UserFormObject.VB_UserMemId = 0

    Set uf = o
    With uf
        Set Header = .Controls("Header")
        Set Bottom = .Controls("Bottom")
        Set CloseButton = .Controls("CloseButton")
        Set RunButton = .Controls("RunButton")
        Set Ver = .Controls("Version")
    End With

    Call clsApi.NonTitleBar(uf.name)                            ' フラットスタイル

    color = GetColor(DarkBLUE)
    Call InitColor
    Call InitPosition

End Property

' [Let] テーマカラー
Public Property Let ThemeColor(ByVal c As FormColors)
    color = GetColor(c)
    madeup = False
    Call InitColor
End Property

' [Let] バージョン情報
Public Property Let Version(v As String)
    Ver.Visible = True
    Ver.Caption = "version " & v
End Property

' [Let] テーマカラー
Public Property Get ThemeColor() As Long
    ThemeColor = color
End Property

'**********************************
'クラスメソッド
'**********************************
' ChangeCursor:        カーソルの形状変更
' GetColor:                 色データの取得
' GetCharactor:         特殊文字の取得
' BeepEx:                   効果音を鳴らす
' ViewDesktop:           デスクトップを表示する

'<summary>
'  カーソルの形状変更
'
'  cType: カーソルタイプ　列挙型から選んで
'</summary>
Public Sub ChangeCursor(cType As CursorType)
    Call clsApi.ChangeCursor(cType)
End Sub

'<summary>
'  色データの取得
'
'  colors: 色　列挙型から選んで
'</summary>
Public Function GetColor(ByVal colors As FormColors) As Long
    Select Case colors
        Case FormColors.WHITE:                  GetColor = RGB(255, 255, 255)
        Case FormColors.GRAY01:                GetColor = RGB(50, 50, 50) 'RGB(64, 64, 64)
        Case FormColors.GRAY02:                GetColor = RGB(192, 192, 192)
        Case FormColors.GRAY03:                GetColor = RGB(224, 224, 224)
        Case FormColors.GRAY04:                GetColor = RGB(238, 238, 238)
        Case FormColors.RED:                      GetColor = RGB(230, 0, 51)

        Case FormColors.LightRED:              GetColor = RGB(208, 87, 107)
        Case FormColors.LightGREEN:          GetColor = RGB(108, 187, 90)
        Case FormColors.LightBLUE:             GetColor = RGB(39, 146, 195)
        Case FormColors.LightBROUN:         GetColor = RGB(141, 100, 73)
        Case FormColors.LightPURPLE:         GetColor = RGB(145, 92, 139)

        Case FormColors.DarkRED:              GetColor = RGB(164, 55, 58)
        Case FormColors.DarkGREEN:          GetColor = RGB(0, 110, 84)
        Case FormColors.DarkBLUE:             GetColor = RGB(50, 66, 107)
        Case FormColors.ToyoBlue:          GetColor = RGB(48, 84, 150)

        Case FormColors.ORANGE:               GetColor = RGB(238, 121, 72)
        Case FormColors.PEAR:                GetColor = RGB(199, 179, 112)
        Case FormColors.MINT:                    GetColor = RGB(0, 172, 154)

        Case FormColors.BLACK:                  GetColor = RGB(60, 60, 60)
    End Select
End Function

'<summary>
'  特殊文字の取得
'
'  chars: 文字　列挙型から選んで
'</summary>
Function GetCharactor(c As SpCharType) As String
    Select Case c
        Case SpCharType.CheckBoxOff:     GetCharactor = ChrW(&H2610)
        Case SpCharType.CheckBoxOn:     GetCharactor = ChrW(&H2611)
        Case SpCharType.CheckBoxX:       GetCharactor = ChrW(&H2612)
        Case SpCharType.OptionOff:          GetCharactor = ChrW(&H25CB)
        Case SpCharType.OptionOn:          GetCharactor = ChrW(&H25C9)

        Case SpCharType.ISeting:              GetCharactor = ChrW(&HE115)
        Case SpCharType.ICamera:            GetCharactor = ChrW(&HE114)
        Case SpCharType.ISave:                GetCharactor = ChrW(&HE105)
        Case SpCharType.IDelete:              GetCharactor = ChrW(&HE107)
        Case SpCharType.ICalender:          GetCharactor = ChrW(&HE163)
        Case SpCharType.IShere:               GetCharactor = ChrW(&HE2FD)
        Case SpCharType.IStart:                GetCharactor = ChrW(&HE102)
        Case SpCharType.IPause:               GetCharactor = ChrW(&HE103)
        Case SpCharType.IStop:                 GetCharactor = ChrW(&HE004)

        Case SpCharType.ICherryBlossom: GetCharactor = ChrW(&H273F)
        Case SpCharType.ICircle:                GetCharactor = ChrW(&H26AB)
        Case SpCharType.ICopyRight:         GetCharactor = ChrW(&HA9)
        
    End Select
End Function

'<summary>
'  C:WindowsMedia フォルダの中の.wav音源を鳴らします。。
'
'  sounds: 音　列挙型から選んで
'</summary>
Public Sub BeepEx(sounds As SoundType)
    Call clsApi.BeepEx(sounds)
End Sub

'<summary>
'  デスクトップを表示する
'</summary>
Public Sub ViewDesktop()
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.MinimizeAll
    Set shell = Nothing
End Sub

'**********************************
'クラスイベント
'**********************************
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Application.WindowState = xlMaximized
    Set clsApi = Nothing
End Sub

'**********************************
'ユーザーフォーム
'**********************************
Private Sub uf_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If madeup = False Then Call InitColor
End Sub

'**********************************
'ヘッダーラベル
'**********************************
Private Sub Header_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If madeup = False Then Call InitColor
    Call clsApi.ChangeCursor(Cross)
End Sub

Private Sub Header_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsApi.FormDrag(uf.name, Button)
End Sub

'**********************************
'ボトムラベル
'**********************************
Private Sub Bottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If madeup = False Then Call InitColor
End Sub

'**********************************
'右上×ラベル
'**********************************
Private Sub CloseButton_Click()
    Unload uf
End Sub

Private Sub CloseButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CloseButton.ForeColor = GetColor(WHITE)
    CloseButton.BackColor = GetColor(RED)
    madeup = False
End Sub

'**********************************
'実行ボタン
'**********************************
Private Sub RunButton_Click()
    Call uf.ClickRunButton
End Sub

Private Sub RunButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call clsApi.ChangeCursor(Hand)
    RunButton.ForeColor = color
    RunButton.BackColor = GetColor(WHITE)
    madeup = False
End Sub

'**********************************
'サブルーチン
'**********************************
Private Sub InitPosition()
    
    uf.Top = 0
    uf.Left = 0

    Header.Top = uf.Top + 1
    Header.Left = uf.Left + 1
    Header.Width = uf.Width - 3

    Bottom.Top = uf.Height - Bottom.Height
    Bottom.Left = uf.Left
    Bottom.Width = uf.Width

    CloseButton.Top = uf.Top + 1
    CloseButton.Left = Header.Width - CloseButton.Width

    Ver.Visible = False

End Sub

Private Sub InitColor()

    uf.BorderColor = color

    Header.ForeColor = GetColor(WHITE)
    Header.BackColor = color

    Bottom.BackColor = GetColor(GRAY04)

    CloseButton.ForeColor = GetColor(WHITE)
    CloseButton.BackColor = color

    RunButton.ForeColor = GetColor(WHITE)
    RunButton.BackColor = color

    Ver.ForeColor = GetColor(GRAY01)
    Ver.BackColor = GetColor(GRAY04)

    madeup = True
End Sub













