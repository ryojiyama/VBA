VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Helmet 
   Caption         =   "Template Form"
   ClientHeight    =   12348
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   7260
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

Private Sub Label15_Click()

End Sub

Private Sub ComboBox_Hinban_Enter()
    If Me.ComboBox_Hinban.text = "品番を半角で入力してください。No.はいらないです" Then
        Me.ComboBox_Hinban.text = ""
    End If
End Sub




Private Sub RunButton_Click()

    Dim ws As Worksheet
    Dim iRow As Long
    Dim id As String
    Dim rng As Range
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' TextBox_IDが空でなければ、そのIDがある行を探し、なければ最終行を選択
    If TextBox_ID.Value <> "" Then
        Set rng = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row).Find(TextBox_ID.Value, LookIn:=xlValues)
        If Not rng Is Nothing Then
            iRow = rng.row
        Else
            iRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
            ws.Cells(iRow, "B").Value = TextBox_ID.Value  'TextBox_IDが見つからなかった場合はB列にTextBox_IDの値を記入
        End If
    Else
        ' 最終行を取得
        iRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
    End If

    ' Captionを格納する変数を定義
    Dim boutaiLot As String
    Dim naiouLot As String

    ' Captionの値を取得
    boutaiLot = Form_Helmet.DateLabel_BoutaiLot.Caption
    ws.Cells(iRow, ws.Rows(1).Find("帽体ロット").Column).Value = boutaiLot
    naiouLot = Form_Helmet.DateLabel_NaisouLot.Caption
    ws.Cells(iRow, ws.Rows(1).Find("内装ロット").Column).Value = naiouLot

    ' 各列を見つけてデータを入力
    ws.Cells(iRow, ws.Rows(1).Find("温度").Column).Value = TextBox_Ondo.Value
    ws.Cells(iRow, ws.Rows(1).Find("品番").Column).Value = ComboBox_Hinban.Value
    ws.Cells(iRow, ws.Rows(1).Find("帽体色").Column).Value = ComboBox_Iro.Value
    ws.Cells(iRow, ws.Rows(1).Find("前処理").Column).Value = ComboBox_Syori.Value
    ws.Cells(iRow, ws.Rows(1).Find("天頂すきま").Column).Value = TextBox_Sukima.Value
    ws.Cells(iRow, ws.Rows(1).Find("重量").Column).Value = TextBox_Jyuryo.Value
    ws.Cells(iRow, ws.Rows(1).Find("構造_検査結果").Column).Value = "合格"
    ws.Cells(iRow, ws.Rows(1).Find("耐貫通_検査結果").Column).Value = "合格"
End Sub




Private Sub TextBox_Ondo_Enter()
    If Me.TextBox_Ondo.text = "数値を半角で入力してください" Then
        Me.TextBox_Ondo.text = ""
    End If
End Sub


Private Sub TextBox_Syori_Change()

End Sub

'**********************************
'user form
'**********************************
Private Sub UserForm_Initialize()

    ' clsAssistの設定
    clsAssist = Me
    clsAssist.ThemeColor = ToyoBlue   ' お好きな初期色を設定
    clsAssist.Version = "2.0"

    ' 試験IDに説明を表示する
    Me.Label_ID.ControlTipText = "IDを入力しない場合は最新の結果に項目を追加します"

    ' Worksheetの定義
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Setting")
    
    ' Hinbanのリストデータを定義
    Dim rngHinban As Range
    Set rngHinban = ws.Range("F2:F43")
    
    ' リストデータを配列に読み込む
    Dim dataArrHinban As Variant
    dataArrHinban = rngHinban.Value
    
    ' ComboBox_Hinbanにデータを追加
    Dim i As Long
    For i = 1 To UBound(dataArrHinban, 1)
        Me.ComboBox_Hinban.AddItem dataArrHinban(i, 1)
    Next i
    Call SetCalender
    
    ' Hinbanのリストデータを定義
    Dim rngSyori As Range
    Set rngSyori = ws.Range("I2:I4")
    
    ' リストデータを配列に読み込む
    Dim dataArrSyori As Variant
    dataArrSyori = rngSyori.Value
    
    ' ComboBox_Hinbanにデータを追加
    For i = 1 To UBound(dataArrSyori, 1)
        Me.ComboBox_Syori.AddItem dataArrSyori(i, 1)
    Next i
    
    Call SetCalender
End Sub
Private Sub TextBox_ID_Enter()
    If TextBox_ID.text = "デフォルトテキスト" Then
        TextBox_ID.text = "HTC"
    End If
End Sub

Private Sub TextBox_ID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox_ID.text = "HTC" Then
        TextBox_ID.text = "デフォルトテキスト"
    End If
End Sub


' 帽体の種類に合わせて'ComboBox_Hinban'の値を変化させる。
Private Sub ComboBox_Hinban_Change()
    Dim hinbanValue As String
    Dim listName As String
    Dim listRange As Range
    Dim cell As Range
    
    ' 'ComboBox_Hinban'の値を取得
    hinbanValue = Me.ComboBox_Hinban.Value
    
    ' リストの名前を設定
    Select Case True
        Case InStr(hinbanValue, "100") > 0, InStr(hinbanValue, "101") > 0
            listName = "ColourList_100"
        Case InStr(hinbanValue, "105") > 0
            listName = "ColourList_105"
        Case InStr(hinbanValue, "110") > 0, InStr(hinbanValue, "S110") > 0
            listName = "ColourList_White"
        Case InStr(hinbanValue, "140") > 0
            listName = "ColourList_140"
        Case InStr(hinbanValue, "170") > 0
            listName = "ColourList_170"
        Case InStr(hinbanValue, "300") > 0, InStr(hinbanValue, "310") > 0
            listName = "ColourList_300"
        Case InStr(hinbanValue, "360") > 0
            listName = "ColourList_360"
        Case InStr(hinbanValue, "370") > 0
            listName = "ColourList_370"
        Case InStr(hinbanValue, "380") > 0
            listName = "ColourList_380"
        Case InStr(hinbanValue, "390") > 0
            listName = "ColourList_390"
        Case InStr(hinbanValue, "391") > 0
            listName = "ColourList_391"
        Case InStr(hinbanValue, "393") > 0
            listName = "ColourList_393"
        Case InStr(hinbanValue, "396") > 0
            listName = "ColourList_396"
        Case InStr(hinbanValue, "170") > 0, InStr(hinbanValue, "LF170") > 0, InStr(hinbanValue, "170S") > 0, InStr(hinbanValue, "170F") > 0
    listName = "ColourList_LF170"
        Case InStr(hinbanValue, "215") > 0, InStr(hinbanValue, "220") > 0, InStr(hinbanValue, "260") > 0, InStr(hinbanValue, "280") > 0
    listName = "ColourList_White"

        ' 追加の条件に対応する場合は、Case文を追加します
        ' Case InStr(hinbanValue, "XXX") > 0
        '     listName = "ColourList_XXX"
        ' Case InStr(hinbanValue, "YYY") > 0
        '     listName = "ColourList_YYY"
        ' Case Else
        '     listName = ""
    End Select
    
    ' 'ComboBox_Iro'のリストを変更
    If listName <> "" Then
        Set listRange = Worksheets("Setting").Range("G2:G100") ' G列の範囲を設定 (適切な範囲に変更してください)
        
        ' 'ComboBox_Iro'のリストをクリア
        Me.ComboBox_Iro.Clear
        
        ' 選択されたリスト名に対応する値をリストに追加
        For Each cell In listRange
            If cell.Value = listName Then
                Me.ComboBox_Iro.AddItem cell.Offset(0, 1).Value ' H列の値を追加 (適切なオフセット値に変更してください)
            End If
        Next cell
    End If
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







