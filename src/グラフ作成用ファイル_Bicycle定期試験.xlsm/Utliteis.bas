Attribute VB_Name = "Utliteis"
'CopiedSheetNames�ŋL����Ă���V�[�g���폜����B
Sub DeleteCopiedSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "CopiedSheetNames�V�[�g��������܂���B"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    Application.DisplayAlerts = False
    For i = 1 To lastRow
        Dim sheetName As String
        sheetName = ws.Cells(i, 1).value
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    ClearCopiedSheetNames
End Sub
'CopiedSheetNames���N���A����B
Sub ClearCopiedSheetNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Cells.ClearContents
    End If
End Sub
    ' CopiedSheetNames�ɋL�ڂ���Ă���V�[�g�̃`���[�g���폜����v���V�[�W��
Public Sub DeleteChartsOnListedSheets()

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim processedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    Dim chartObj As ChartObject
    
    ' CopiedSheetNames �V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("CopiedSheetNames")
    Set processedSheets = New Collection ' �����ς݃V�[�g����ǐՂ���R���N�V����
    
    ' A��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' A��̒l�����[�v
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value
        
        On Error Resume Next
        ' �R���N�V�����ɓ������O�����ɑ��݂��邩�`�F�b�N
        processedSheets.Add sheetName, sheetName
        
        If Err.Number = 0 Then ' �ǉ������������ꍇ�A�V�[�g�͂܂���������Ă��Ȃ�
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                ' �V�[�g��̂��ׂẴO���t�I�u�W�F�N�g���폜
                For Each chartObj In wsTarget.ChartObjects
                    chartObj.Delete
                Next chartObj
            End If
        End If
        
        On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g
        Set wsTarget = Nothing
    Next i
End Sub

Sub PrintFirstPageOfUniqueListedSheets()
    ' �w�肳�ꂽ�����[��1�y�[�W�ڂ��A�d���Ȃ�1�񂸂������v���V�[�W��
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim printedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' CopiedSheetNames �V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("CopiedSheetNames")
    Set printedSheets = New Collection ' ������ꂽ�V�[�g����ǐՂ���R���N�V����

    ' A��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    ' A��̒l�����[�v
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value

        On Error Resume Next
        ' �R���N�V�����ɓ������O�����ɑ��݂��邩�`�F�b�N
        printedSheets.Add sheetName, sheetName
        If Err.Number = 0 Then ' �ǉ������������ꍇ�A�V�[�g�͂܂��������Ă��Ȃ�
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' �V�[�g��1�y�[�W�ڂ݂̂����
            End If
        End If
        On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g

        Set wsTarget = Nothing
    Next i
End Sub



' Setting��"B2"�Z���Ƀt�H�[�J�X
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    If SheetExists("Setting") Then
        Application.GoTo ActiveWorkbook.Sheets("Setting").Range("B2")
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
