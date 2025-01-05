'*******************************************************************************
' ���W���[�����FSheetSelector
' �ړI�FTestData�V�[�g�̑I���@�\��񋟂���
' �쐬�F2024/12/27
'*******************************************************************************
Option Explicit

'�萔��`
Private Const SHEET_KEYWORD As String = "TestData"
Private mSelectedSheet As String

'*******************************************************************************
' GetTestDataSheets
' �T�v�FTestData���܂ރV�[�g���̔z����擾����
' �ߒl�FVariant - �V�[�g���̔z��A�V�[�g���Ȃ��ꍇ��Empty
'*******************************************************************************
Private Function GetTestDataSheets() As Variant
    Dim ws As Worksheet
    Dim sheetNames As New Collection

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, SHEET_KEYWORD, vbTextCompare) > 0 Then
            sheetNames.Add ws.Name
        End If
    Next ws

    If sheetNames.Count = 0 Then
        GetTestDataSheets = Empty
        Exit Function
    End If

    ' �R���N�V������z��ɕϊ�
    Dim result() As String
    ReDim result(1 To sheetNames.Count)

    Dim i As Long
    For i = 1 To sheetNames.Count
        result(i) = sheetNames(i)
    Next i

    GetTestDataSheets = result
End Function

'*******************************************************************************
' ShowSheetSelector
' �T�v�F�V�[�g�I���t�H�[����\������
' �ߒl�FString - �I�����ꂽ�V�[�g���A�L�����Z�����͋󕶎�
'*******************************************************************************
Public Function ShowSheetSelector() As String
    Dim sheets As Variant
    sheets = GetTestDataSheets()

    If IsEmpty(sheets) Then
        MsgBox "TestData���܂ރV�[�g��������܂���B", vbExclamation
        ShowSheetSelector = ""
        Exit Function
    End If

    ' UserForm�̍쐬�ƕ\��
    With SheetSelectorForm
        .Initialize sheets
        .Show
        ShowSheetSelector = .SelectedSheet
    End With
End Function

' UserForm�̃R�[�h�iSheetSelectorForm�j
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetSelectorForm
   Caption         =   "�f�[�^�V�[�g�̑I��"
   ClientHeight    =   2280
   ClientWidth     =   4560
   OleObjectBlob   =   "SheetSelectorForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End

Option Explicit

Private mSelectedSheet As String

Private Sub UserForm_Initialize()
    Me.Caption = "�f�[�^�V�[�g�̑I��"
    lstSheets.MultiSelect = False
End Sub

Public Sub Initialize(sheets As Variant)
    Dim i As Long
    lstSheets.Clear

    For i = LBound(sheets) To UBound(sheets)
        lstSheets.AddItem sheets(i)
    Next i

    If lstSheets.ListCount > 0 Then
        lstSheets.Selected(0) = True
    End If
End Sub

Private Sub cmdOK_Click()
    If lstSheets.ListIndex = -1 Then
        MsgBox "�V�[�g��I�����Ă��������B", vbExclamation
        Exit Sub
    End If

    mSelectedSheet = lstSheets.Value
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mSelectedSheet = ""
    Me.Hide
End Sub

Public Property Get SelectedSheet() As String
    SelectedSheet = mSelectedSheet
End Property

'*******************************************************************************
' TransferSortedData
' �T�v�F�\�[�g���ꂽ�f�[�^��ChartSheet�Ɉڍs����
' ��ʁF���C���v���V�[�W��
' �ΏہF����=LOG_Helmet, �o��=ChartSheet
' �ߒl�FBoolean - ������������True�A���s����False
'*******************************************************************************
Public Function TransferSortedData() As Boolean
    ' �V�[�g�I��
    Dim sourceSheetName As String
    sourceSheetName = ShowSheetSelector()

    If sourceSheetName = "" Then
        MsgBox "�V�[�g���I������Ă��܂���B", vbExclamation
        TransferSortedData = False
        Exit Function
    End If

    ' �ȉ��A�����̃R�[�h���C��
    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long

    ' ������
    TransferSortedData = False

    ' �V�[�g�̑��݊m�F�Ǝ擾
    If Not CheckAndGetSheets(wsSource, wsTarget, sourceSheetName) Then
        MsgBox "�K�v�ȃV�[�g�̊m�F�Ɏ��s���܂����B", vbExclamation
        Exit Function
    End If

    ' �f�[�^�����̊m�F
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "�f�[�^�����݂��܂���B", vbExclamation
        Exit Function
    End If

    ' ���R�[�h���`�F�b�N
    If (lastRow - 1) > MAX_RECORDS Then
        MsgBox "�f�[�^��" & MAX_RECORDS & "���𒴂��Ă��܂��B" & vbNewLine & _
               "����Ƀ\�[�g���Č������i���Ă��������B" & vbNewLine & _
               "���݂̃��R�[�h��: " & (lastRow - 1) & "��", vbExclamation
        Exit Function
    End If

    ' �^�[�Q�b�g�V�[�g�̃N���A
    ClearTargetSheet wsTarget

    ' �f�[�^�̈ڍs
    If Not CopyDataToTarget(wsSource, wsTarget, lastRow) Then
        Exit Function
    End If

    TransferSortedData = True
    Exit Function

ErrorHandler:
    MsgBox "�G���[���������܂����B" & vbNewLine & _
           "�G���[�ԍ�: " & Err.Number & vbNewLine & _
           "�G���[�̐���: " & Err.Description, vbCritical
    TransferSortedData = False
End Function

'*******************************************************************************
' CheckAndGetSheets
' �T�v�F�K�v�ȃV�[�g�̑��݊m�F�Ǝ擾
' ��ʁF�T�u�v���V�[�W��
' �����FwsSource - ���V�[�g, wsTarget - �ڍs��V�[�g
' �ߒl�FBoolean - ������True
'*******************************************************************************
Private Function CheckAndGetSheets(ByRef wsSource As Worksheet, _
                                 ByRef wsTarget As Worksheet, _
                                 ByVal sourceSheetName As String) As Boolean
    On Error GoTo ErrorHandler

    CheckAndGetSheets = False

    ' �\�[�X�V�[�g�̊m�F
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    If wsSource Is Nothing Then
        MsgBox sourceSheetName & "�V�[�g��������܂���B", vbExclamation
        Exit Function
    End If

    ' �^�[�Q�b�g�V�[�g�̊m�F
    Set wsTarget = ThisWorkbook.Sheets("ChartSheet")
    If wsTarget Is Nothing Then
        MsgBox "ChartSheet��������܂���B", vbExclamation
        Exit Function
    End If

    CheckAndGetSheets = True
    Exit Function

ErrorHandler:
    MsgBox "�V�[�g�̊m�F���ɃG���[���������܂����B" & vbNewLine & _
           "�G���[�ԍ�: " & Err.Number & vbNewLine & _
           "�G���[�̐���: " & Err.Description, vbCritical
    CheckAndGetSheets = False
End Function

'*******************************************************************************
' ClearTargetSheet
' �T�v�F�ڍs��V�[�g�̃N���A
' ��ʁF�T�u�v���V�[�W��
' �����FwsTarget - �ڍs��V�[�g
'*******************************************************************************
Private Sub ClearTargetSheet(ByRef wsTarget As Worksheet)
    On Error Resume Next
    With wsTarget
        .Cells.Clear
        .Cells.Interior.ColorIndex = xlNone
    End With
End Sub

'*******************************************************************************
' CopyDataToTarget
' �T�v�F�f�[�^���ڍs��V�[�g�ɃR�s�[
' ��ʁF�T�u�v���V�[�W��
' �����FwsSource - ���V�[�g, wsTarget - �ڍs��V�[�g, lastRow - �ŏI�s
' �ߒl�FBoolean - ������True
'*******************************************************************************
Private Function CopyDataToTarget(ByRef wsSource As Worksheet, _
                                ByRef wsTarget As Worksheet, _
                                ByVal lastRow As Long) As Boolean
    On Error GoTo ErrorHandler

    CopyDataToTarget = False

    ' �w�b�_�[�s�̃R�s�[
    wsSource.Rows(1).Copy Destination:=wsTarget.Rows(1)

    ' �f�[�^�s�̃R�s�[
    wsSource.Range("A2:ZZ" & lastRow).Copy _
        Destination:=wsTarget.Range("A2")

    ' �����̒���
    With wsTarget
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
    End With

    CopyDataToTarget = True
    Exit Function

ErrorHandler:
    MsgBox "�f�[�^�̃R�s�[���ɃG���[���������܂����B" & vbNewLine & _
           "�G���[�ԍ�: " & Err.Number & vbNewLine & _
           "�G���[�̐���: " & Err.Description, vbCritical
    CopyDataToTarget = False
End Function
