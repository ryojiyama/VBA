VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' 最終更新日：2021/04/30


' この辺はForm Windowの枠消し&ウィンドウドラックのため
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "user32.dll" ()
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAcessible As Object, ByRef hWnd As Long) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Const GWL_STYLE = (-16&)
Const GWL_EXSTYLE = (-20&)
Const WS_CAPTION = &HC00000
Const WS_EX_DLGMODALFRAME = &H1&

' マウスカーソルを変更するためのもの
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_ARROW = 32512&        '矢印カーソル
Private Const IDC_HAND = 32649&           'ハンドカーソル
Private Const IDC_IBEAM = 32513&          'Iカーソル
Private Const IDC_WAIT = 32514&            '砂時計カーソル
Private Const IDC_SIZEALL = 32646&        '4方向の矢印カーソル
Private mCursorDefault As Long, mCursorHand As Long, mCursorI As Long, mCursorWait As Long, mCursorCross As Long

' 待ち時間処理のため
'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' フォームフェードインのため
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As AnimateWindowFlags) As Boolean

' 効果音用
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ASYNC = &H1


'///////// enumerated type //////////

Public Enum CursorType          ' ChangeCursorで使用。
    Default = 0: Hand = 1: Ibeam = 2: Wait = 3: Cross = 4
End Enum

Public Enum SoundType          ' BeepExで使用。
    StartUp = 1
    ShutDown = 2
    Infomation = 3
    Error = 4
    Warning = 5
    
    Completed = 10
End Enum

Public Enum AnimateWindowFlags
    AW_HOR_POSITIVE = &H1
    AW_HOR_NEGATIVE = &H2
    AW_VER_POSITIVE = &H4
    AW_VER_NEGATIVE = &H8
    AW_CENTER = &H10
    AW_HIDE = &H10000
    AW_ACTIVATE = &H20000
    AW_SLIDE = &H40000
    AW_BLEND = &H80000
End Enum


'///////// module list //////////

' NonTitleBar:            ユーザーフォームのタイトルバー非表示
' FormDrag:               ユーザーフォームのタイトルバー以外でドラッグ可能にする
' ChangeCursor:        カーソルの形状変更
' BeepEx:                   効果音を鳴らす
' ThisScreenWidth:     画面の幅
' ThisScreenHeight:    画面の高さ



'///////// module //////////

Private Sub Class_Initialize()
    mCursorDefault = LoadCursor(0&, IDC_ARROW)   'マウスポインターを標準の矢印の形にするためのマウスカーソルを取得
    mCursorHand = LoadCursor(0&, IDC_HAND)         '手の形
    mCursorI = LoadCursor(0&, IDC_IBEAM)                'I の形
    mCursorWait = LoadCursor(0&, IDC_WAIT)            '砂時計の形
    mCursorCross = LoadCursor(0&, IDC_SIZEALL)      '上下左右矢印の形
End Sub

'<summary>
'  ユーザーフォームのタイトルバー非表示　返り値：0=失敗 0<>成功 変更前のウィンドウスタイルの値
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'</summary>
Function NonTitleBar(objName As String) As Long

    Dim wnd As Long, formHeight As Double
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.name = objName Then Set uForm = uf: Exit For
    Next
    
    formHeight = uForm.InsideHeight
    WindowFromAccessibleObject uForm, wnd
    
    SetWindowLong wnd, GWL_EXSTYLE, GetWindowLong(wnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    NonTitleBar = SetWindowLong(wnd, GWL_STYLE, GetWindowLong(wnd, GWL_STYLE) And Not WS_CAPTION)
    
    DrawMenuBar wnd
    uForm.Height = uForm.Height - uForm.InsideHeight + formHeight
End Function

'<summary>
'  ユーザーフォームのタイトルバー以外でドラッグ可能にする
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'  Button: MouseMoveイベントのButtonをそのまま渡す
'</summary>
Sub FormDrag(objName As String, ByVal Button As Integer)

    Dim hWnd As Long
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.name = objName Then Set uForm = uf: Exit For
    Next
    
    If Button = 1 Then
        WindowFromAccessibleObject uForm, hWnd
        ReleaseCapture
        Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub

'<summary>
'  カーソルの形状変更
'
'  cType: カーソルタイプ　列挙型から選んで
'</summary>
Sub ChangeCursor(cType As CursorType)
    Select Case cType
        Case CursorType.Default: Call SetCursor(mCursorDefault)
        Case CursorType.Hand:    Call SetCursor(mCursorHand)
        Case CursorType.Ibeam:  Call SetCursor(mCursorI)
        Case CursorType.Wait:     Call SetCursor(mCursorWait)
        Case CursorType.Cross:    Call SetCursor(mCursorCross)
    End Select
End Sub



'<summary>
'  画面の幅
'</summary>
Function ThisScreenWidth() As Long
    Windows(ThisWorkbook.name).WindowState = xlMaximized           ' win10/Excel2013 環境だと上２行
    ThisScreenWidth = Windows(ThisWorkbook.name).Width
'    Application.WindowState = xlMaximized                                            ' win7/Excel2010 環境だと下２行
'    ThisScreenWidth = Application.Width
End Function

'<summary>
'  画面の高さ
'</summary>
Function ThisScreenHeight() As Long
    Windows(ThisWorkbook.name).WindowState = xlMaximized           ' win10/Excel2013 環境だと上２行
    ThisScreenHeight = Windows(ThisWorkbook.name).Height
'    Application.WindowState = xlMaximized                                            ' win7/Excel2010 環境だと下２行
'    ThisScreenHeight = Application.Height
End Function
















