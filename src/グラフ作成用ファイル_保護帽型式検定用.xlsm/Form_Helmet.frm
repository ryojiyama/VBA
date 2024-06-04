VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Helmet 
   Caption         =   "Template Form"
   ClientHeight    =   12348
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   7260
   OleObjectBlob   =   "Form_Helmet.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "Form_Helmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' �ŏI�X�V���F2021/05/05


'**************************
'���@�ˑ��֌W(�ȉ��̃��W���[���������ƍŒ���̓���o���Ȃ��ł�)
' �N���X���W���[��
'  - clsFormAssistant.cls
'  - clsWinApi.cls
'
' clsColorPalette.cls �̓��[�U�[�R���e���c�ł��̂ŁA�s�v�ł���Ώ����Ă��������Ă��\���܂���B
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
    If Me.ComboBox_Hinban.text = "�i�Ԃ𔼊p�œ��͂��Ă��������BNo.�͂���Ȃ��ł�" Then
        Me.ComboBox_Hinban.text = ""
    End If
End Sub




Private Sub RunButton_Click()

    Dim ws As Worksheet
    Dim iRow As Long
    Dim id As String
    Dim rng As Range
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' TextBox_ID����łȂ���΁A����ID������s��T���A�Ȃ���΍ŏI�s��I��
    If TextBox_ID.Value <> "" Then
        Set rng = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row).Find(TextBox_ID.Value, LookIn:=xlValues)
        If Not rng Is Nothing Then
            iRow = rng.row
        Else
            iRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
            ws.Cells(iRow, "B").Value = TextBox_ID.Value  'TextBox_ID��������Ȃ������ꍇ��B���TextBox_ID�̒l���L��
        End If
    Else
        ' �ŏI�s���擾
        iRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
    End If

    ' Caption���i�[����ϐ����`
    Dim boutaiLot As String
    Dim naiouLot As String

    ' Caption�̒l���擾
    boutaiLot = Form_Helmet.DateLabel_BoutaiLot.Caption
    ws.Cells(iRow, ws.Rows(1).Find("�X�̃��b�g").Column).Value = boutaiLot
    naiouLot = Form_Helmet.DateLabel_NaisouLot.Caption
    ws.Cells(iRow, ws.Rows(1).Find("�������b�g").Column).Value = naiouLot

    ' �e��������ăf�[�^�����
    ws.Cells(iRow, ws.Rows(1).Find("���x").Column).Value = TextBox_Ondo.Value
    ws.Cells(iRow, ws.Rows(1).Find("�i��").Column).Value = ComboBox_Hinban.Value
    ws.Cells(iRow, ws.Rows(1).Find("�X�̐F").Column).Value = ComboBox_Iro.Value
    ws.Cells(iRow, ws.Rows(1).Find("�O����").Column).Value = ComboBox_Syori.Value
    ws.Cells(iRow, ws.Rows(1).Find("�V��������").Column).Value = TextBox_Sukima.Value
    ws.Cells(iRow, ws.Rows(1).Find("�d��").Column).Value = TextBox_Jyuryo.Value
    ws.Cells(iRow, ws.Rows(1).Find("�\��_��������").Column).Value = "���i"
    ws.Cells(iRow, ws.Rows(1).Find("�ϊђ�_��������").Column).Value = "���i"
End Sub




Private Sub TextBox_Ondo_Enter()
    If Me.TextBox_Ondo.text = "���l�𔼊p�œ��͂��Ă�������" Then
        Me.TextBox_Ondo.text = ""
    End If
End Sub


Private Sub TextBox_Syori_Change()

End Sub

'**********************************
'user form
'**********************************
Private Sub UserForm_Initialize()

    ' clsAssist�̐ݒ�
    clsAssist = Me
    clsAssist.ThemeColor = ToyoBlue   ' ���D���ȏ����F��ݒ�
    clsAssist.Version = "2.0"

    ' ����ID�ɐ�����\������
    Me.Label_ID.ControlTipText = "ID����͂��Ȃ��ꍇ�͍ŐV�̌��ʂɍ��ڂ�ǉ����܂�"

    ' Worksheet�̒�`
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Setting")
    
    ' Hinban�̃��X�g�f�[�^���`
    Dim rngHinban As Range
    Set rngHinban = ws.Range("F2:F43")
    
    ' ���X�g�f�[�^��z��ɓǂݍ���
    Dim dataArrHinban As Variant
    dataArrHinban = rngHinban.Value
    
    ' ComboBox_Hinban�Ƀf�[�^��ǉ�
    Dim i As Long
    For i = 1 To UBound(dataArrHinban, 1)
        Me.ComboBox_Hinban.AddItem dataArrHinban(i, 1)
    Next i
    Call SetCalender
    
    ' Hinban�̃��X�g�f�[�^���`
    Dim rngSyori As Range
    Set rngSyori = ws.Range("I2:I4")
    
    ' ���X�g�f�[�^��z��ɓǂݍ���
    Dim dataArrSyori As Variant
    dataArrSyori = rngSyori.Value
    
    ' ComboBox_Hinban�Ƀf�[�^��ǉ�
    For i = 1 To UBound(dataArrSyori, 1)
        Me.ComboBox_Syori.AddItem dataArrSyori(i, 1)
    Next i
    
    Call SetCalender
End Sub
Private Sub TextBox_ID_Enter()
    If TextBox_ID.text = "�f�t�H���g�e�L�X�g" Then
        TextBox_ID.text = "HTC"
    End If
End Sub

Private Sub TextBox_ID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox_ID.text = "HTC" Then
        TextBox_ID.text = "�f�t�H���g�e�L�X�g"
    End If
End Sub


' �X�̂̎�ނɍ��킹��'ComboBox_Hinban'�̒l��ω�������B
Private Sub ComboBox_Hinban_Change()
    Dim hinbanValue As String
    Dim listName As String
    Dim listRange As Range
    Dim cell As Range
    
    ' 'ComboBox_Hinban'�̒l���擾
    hinbanValue = Me.ComboBox_Hinban.Value
    
    ' ���X�g�̖��O��ݒ�
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

        ' �ǉ��̏����ɑΉ�����ꍇ�́ACase����ǉ����܂�
        ' Case InStr(hinbanValue, "XXX") > 0
        '     listName = "ColourList_XXX"
        ' Case InStr(hinbanValue, "YYY") > 0
        '     listName = "ColourList_YYY"
        ' Case Else
        '     listName = ""
    End Select
    
    ' 'ComboBox_Iro'�̃��X�g��ύX
    If listName <> "" Then
        Set listRange = Worksheets("Setting").Range("G2:G100") ' G��͈̔͂�ݒ� (�K�؂Ȕ͈͂ɕύX���Ă�������)
        
        ' 'ComboBox_Iro'�̃��X�g���N���A
        Me.ComboBox_Iro.Clear
        
        ' �I�����ꂽ���X�g���ɑΉ�����l�����X�g�ɒǉ�
        For Each cell In listRange
            If cell.Value = listName Then
                Me.ComboBox_Iro.AddItem cell.Offset(0, 1).Value ' H��̒l��ǉ� (�K�؂ȃI�t�Z�b�g�l�ɕύX���Ă�������)
            End If
        Next cell
    End If
End Sub





Private Sub UserForm_Terminate()
    Set clsAssist = Nothing
End Sub

'**********************************
'�o�̓{�^�������������ɑ��鏈��
'**********************************
Public Sub ClickRunButton()
    Debug.Print "run"
End Sub

'**********************************
'���[�U�[�R���e���c
'**********************************
Private Sub SetPalette()     ' �J���[�p���b�g�̐F�I���Ńt�H�[���J���[��ύX���܂��B

    
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







