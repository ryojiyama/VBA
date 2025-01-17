Attribute VB_Name = "OotputReport"
Option Explicit
Public Sub ShowOutputDialog()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("�o�͌`����I�����Ă�������" & vbNewLine & _
                     "�͂��FPDF�o��" & vbNewLine & _
                     "�������F�v�����^�o��", _
                     vbQuestion + vbYesNo, _
                     "�o�͌`���̑I��")
    
    If response = vbYes Then
        ExportReport "PDF"
    Else
        ExportReport "Print"
    End If
End Sub
' ���C�������̃T�u�v���V�[�W��
Sub ExportReport(Optional outputType As String = "PDF")
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, pageNum As Long
    Dim currentHeight As Long
    Dim pageStartRow As Long
    Dim splitRow As Long
    Dim nextStartRow As Long
    Dim outputPath As String
    Dim workbookPath As String
    Dim hasNewColumn As Boolean
    
    ' "���|�[�g�O���t"�V�[�g��ݒ�
    If WorksheetExists("���|�[�g�O���t") = False Then
        MsgBox "���|�[�g�O���t�V�[�g��������܂���B", vbExclamation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    
    ' NewColumn�̑��݃`�F�b�N
    hasNewColumn = CheckNewColumnExists(ws, lastRow)
    If Not hasNewColumn Then
        MsgBox "NewColumn��������܂���B�����𒆎~���܂��B", vbExclamation
        Exit Sub
    End If
    
    ' ���[�N�u�b�N�̃p�X���擾�ƌ���
    workbookPath = ValidateAndGetWorkbookPath()
    If workbookPath = "" Then Exit Sub
    
    ' �ŏ���NewColumn�s��T��
    pageStartRow = FindFirstNewColumnRow(ws, lastRow)
    
    ' �y�[�W�ԍ��ƍ����̏�����
    pageNum = 1
    currentHeight = 0
    
    ' �ŏ��̃y�[�W�ݒ�
    SetupPageFormat ws
    
    ' �s���������Ȃ���y�[�W����
    For i = pageStartRow To lastRow
        currentHeight = currentHeight + ws.Rows(i).RowHeight
        
        ' �ݐύ�����728�𒴂����ꍇ�̏���
        If currentHeight >= 728 Then
            splitRow = FindSplitRow(ws, i, pageStartRow)
            nextStartRow = splitRow + 1
            
            ' ����͈͂�ݒ�
            ws.PageSetup.PrintArea = "$A$" & pageStartRow & ":$G$" & splitRow
            
            ' �o�͏���
            Select Case outputType
                Case "PDF"
                    outputPath = workbookPath & "Report_" & Format(pageNum, "000") & ".pdf"
                    If Not ExportPageToPDF(ws, outputPath) Then Exit Sub
                Case "Print"
                    If Not ExportPageToPrinter(ws) Then Exit Sub
            End Select
            
            ' �y�[�W�ݒ�����Z�b�g
            ResetPageFormat ws
            
            ' ���̃y�[�W�̏���
            pageNum = pageNum + 1
            pageStartRow = nextStartRow
            i = nextStartRow - 1
            currentHeight = 0
            SetupPageFormat ws
            
        ElseIf i = lastRow Then
            ' ����͈͂�ݒ�
            ws.PageSetup.PrintArea = "$A$" & pageStartRow & ":$G$" & i
            
            ' �o�͏���
            Select Case outputType
                Case "PDF"
                    outputPath = workbookPath & "Report_" & Format(pageNum, "000") & ".pdf"
                    If Not ExportPageToPDF(ws, outputPath) Then Exit Sub
                Case "Print"
                    If Not ExportPageToPrinter(ws) Then Exit Sub
            End Select
        End If
    Next i
    
    MsgBox "�o�͂��������܂����B" & vbNewLine & _
           "�ۑ���: " & workbookPath, vbInformation
    Exit Sub

ErrorHandler:
    HandleError Err.Number, Err.Description
End Sub

' PDF�o�͐�p�̊֐�
Private Function ExportPageToPDF(ws As Worksheet, pdfPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' �t�@�C���̎g�p��Ԃ��`�F�b�N
    If FileExists(pdfPath) Then
        If IsFileInUse(pdfPath) Then
            MsgBox "�t�@�C�� '" & pdfPath & "' �����̃v���Z�X�Ŏg�p���ł��B", vbExclamation
            ExportPageToPDF = False
            Exit Function
        End If
    End If
    
    ' PDF�Ƃ��ĕۑ�
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ExportPageToPDF = True
    Exit Function

ErrorHandler:
    HandleError Err.Number, Err.Description
    ExportPageToPDF = False
End Function

' �v�����^�o�͐�p�̊֐�
Private Function ExportPageToPrinter(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ' �v�����^�֏o��
    ws.PrintOut Copies:=1, Preview:=False, ActivePrinter:=Application.ActivePrinter
    
    ExportPageToPrinter = True
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 1004  ' �v�����^�����Ȃ�
            If InStr(1, Err.Description, "����") > 0 Then
                MsgBox "�v�����^���������Ă��܂���B" & vbNewLine & _
                       "�v�����^�̓d����ڑ����m�F���Ă��������B", vbExclamation
            Else
                MsgBox "�v�����^�G���[���������܂����B" & vbNewLine & _
                       "�v�����^�̏�Ԃ��m�F���Ă��������B", vbExclamation
            End If
        Case Else
            MsgBox "�v�����^�G���[���������܂����B" & vbNewLine & _
                   "�v�����^�̏�Ԃ��m�F���Ă��������B", vbExclamation
    End Select
    ExportPageToPrinter = False
End Function

' �y�[�W�����ʒu��T���֐�
Private Function FindSplitRow(ws As Worksheet, currentRow As Long, startRow As Long) As Long
    Dim j As Long
    
    FindSplitRow = currentRow
    For j = currentRow To startRow Step -1
        If Left(ws.Cells(j, "I").value, 9) = "NewColumn" Then
            FindSplitRow = j - 1
            Exit Function
        End If
    Next j
End Function

' �G���[�����֐�
Private Sub HandleError(ErrorNum As Long, ErrorDesc As String)
    Select Case ErrorNum
        Case 1004  ' �A�v���P�[�V�����܂��͌����G���[
            MsgBox "PDF�t�@�C���̍쐬�������Ȃ����A�܂��͑��̃v���Z�X�Ŏg�p���ł��B" & vbNewLine & _
                   "�G���[�̏ڍ�: " & ErrorDesc, vbCritical
        Case 70, 75  ' �t�@�C���A�N�Z�X�G���[
            MsgBox "PDF�t�@�C���ɃA�N�Z�X�ł��܂���B" & vbNewLine & _
                   "�t�@�C�������̃v���Z�X�ŊJ����Ă��邩�A" & vbNewLine & _
                   "�A�N�Z�X�������Ȃ��\��������܂��B", vbCritical
        Case Else
            MsgBox "�\�����ʃG���[���������܂����B" & vbNewLine & _
                   "�G���[�ԍ�: " & ErrorNum & vbNewLine & _
                   "�G���[�̐���: " & ErrorDesc, vbCritical
    End Select
End Sub

' NewColumn�̑��݃`�F�b�N�֐�
Private Function CheckNewColumnExists(ws As Worksheet, lastRow As Long) As Boolean
    Dim i As Long
    
    CheckNewColumnExists = False
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                CheckNewColumnExists = True
                Exit Function
            End If
        End If
    Next i
End Function

' �ŏ���NewColumn�s��T���֐�
Private Function FindFirstNewColumnRow(ws As Worksheet, lastRow As Long) As Long
    Dim i As Long
    
    FindFirstNewColumnRow = 4  ' �J�n�s��4�s�ڂɕύX
    For i = 4 To lastRow       ' 4�s�ڂ��猟���J�n
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                FindFirstNewColumnRow = i
                Exit Function
            End If
        End If
    Next i
End Function

' ���[�N�u�b�N�p�X�̌��؊֐�
Private Function ValidateAndGetWorkbookPath() As String
    Dim workbookPath As String
    
    workbookPath = ThisWorkbook.Path
    If workbookPath = "" Then
        MsgBox "���[�N�u�b�N���ۑ�����Ă��܂���B��ɕۑ����Ă��������B", vbExclamation
        ValidateAndGetWorkbookPath = ""
        Exit Function
    End If
    
    If Right(workbookPath, 1) <> "\" Then
        workbookPath = workbookPath & "\"
    End If
    
    If Not FolderExists(workbookPath) Then
        MsgBox "�ۑ���t�H���_��������܂���B", vbExclamation
        ValidateAndGetWorkbookPath = ""
        Exit Function
    End If
    
    ValidateAndGetWorkbookPath = workbookPath
End Function

' �V�[�g�̑��݊m�F�֐�
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    WorksheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function
' �t�H���_�̑��݊m�F�֐�
Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

' �t�@�C���̑��݊m�F�֐�
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(filePath) And vbDirectory) <> vbDirectory
    On Error GoTo 0
End Function
' �t�@�C���̎g�p��Ԋm�F�֐�
Private Function IsFileInUse(ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    
    On Error Resume Next
    fileNum = FreeFile()
    Open filePath For Binary Access Read Write Lock Read Write As fileNum
    Close fileNum
    IsFileInUse = (Err.Number <> 0)
    On Error GoTo 0
End Function
' �y�[�W�ݒ���s���T�u�v���V�[�W���i�w�b�_�[�s�ݒ��ǉ��j
Private Sub SetupPageFormat(ws As Worksheet)
    With ws.PageSetup
        ' ��{�ݒ�
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        
        ' �w�b�_�[�s�̐ݒ�i1-2�s�ڂ��J��Ԃ��\���j
        .PrintTitleRows = "$1:$2"
        
        ' �]���ݒ�i�K�v�ɉ����Ē����\�j
        .HeaderMargin = Application.InchesToPoints(0.3)  ' �w�b�_�[�]��
        .TopMargin = Application.InchesToPoints(0.75)    ' ��]��
        .BottomMargin = Application.InchesToPoints(0.75) ' ���]��
        .LeftMargin = Application.InchesToPoints(0.7)    ' ���]��
        .RightMargin = Application.InchesToPoints(0.7)   ' �E�]��
    End With
End Sub

' �y�[�W�ݒ�����Z�b�g����T�u�v���V�[�W��
Private Sub ResetPageFormat(ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ""
        .PrintTitleRows = ""  ' �w�b�_�[�s�ݒ���N���A
        .Zoom = 100
        .FitToPagesWide = False
        .FitToPagesTall = False
    End With
End Sub
