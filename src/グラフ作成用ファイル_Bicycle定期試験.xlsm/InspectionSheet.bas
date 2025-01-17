Attribute VB_Name = "InspectionSheet"
Sub InspectionSheet_Make()
    Call SetupInspectionReport
'    Call TransferDataToAppropriateSheets
'    Call TransferDataToTopImpactTest
'    Call TransferDataToDynamicSheets
'    Call ImpactValueJudgement
'    Call FormatNonContinuousCells
'    Call DistributeChartsToSheets
End Sub

'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�F�����񍐏��V�[�g���쐬���A�V�[�g�����Ǘ�
' �����F�Ȃ�
'*******************************************************************************
Sub SetupInspectionReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row

    Dim groupedData As Object
    Set groupedData = CreateObject("Scripting.Dictionary")
    Dim copiedSheets As Object
    Set copiedSheets = CreateObject("Scripting.Dictionary")
    Dim copiedSheetNames As Collection
    Set copiedSheetNames = New Collection

    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = ws.Cells(i, 2).value

        Dim HelmetData As New HelmetData
        Set HelmetData = ParseHelmetData(cellValue)

'        Dim productNameKey As String
'        productNameKey = HelmetData.GroupNumber & "-" & HelmetData.ProductName

        If Not groupedData.Exists(HelmetData.GroupNumber) Then
            groupedData.Add HelmetData.GroupNumber, New Collection
        End If
        groupedData(HelmetData.GroupNumber).Add HelmetData

        If Not copiedSheets.Exists(HelmetData.productName) Then
            ' 3��ނ̃V�[�g���R�s�[���A�A�ԂŖ��O��ݒ�
            Dim sheetIndex As Long
            Dim sheetName As Variant
            For sheetIndex = 1 To 3
                sheetName = Array("Report01", "Report02", "Report03")(sheetIndex - 1)
                ThisWorkbook.Sheets(sheetName).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                ActiveSheet.Name = CreateUniqueName(HelmetData.productName & "_" & sheetIndex)
                copiedSheetNames.Add ActiveSheet.Name
            Next sheetIndex

            copiedSheets.Add HelmetData.productName, Nothing ' �R�s�[�ς݃t���O�ݒ�
        End If

    Next i

    Debug.Print "Grouped Data:"
    PrintGroupedData groupedData
    SaveCopiedSheetNames copiedSheetNames
End Sub

'*******************************************************************************
' �@�\�F�f�[�^���������͂���HelmetData�I�u�W�F�N�g���쐬
' �����Fvalue - �n�C�t����؂�̃f�[�^������
' �ߒl�FHelmetData - ��͂��ꂽ�f�[�^�I�u�W�F�N�g
'*******************************************************************************
Function ParseHelmetData(value As String) As HelmetData
    Dim parts() As String
    parts = Split(value, "-")
    Dim result As New HelmetData
    
    If UBound(parts) >= 4 Then
        result.GroupNumber = parts(0)
        result.productName = parts(1)
        result.ImpactPosition = parts(2)
        result.ImpactTemp = parts(3)
        result.anvilForm = parts(4)
        result.headModel = parts(5)
    End If
    
    Set ParseHelmetData = result
End Function
'*******************************************************************************
' �@�\�F�d�����Ȃ����j�[�N�ȃV�[�g���𐶐�
' �����FbaseName - ��{�ƂȂ�V�[�g��
' �ߒl�FString - ���j�[�N�ȃV�[�g��
'*******************************************************************************
Function CreateUniqueName(baseName As String) As String
    Dim uniqueName As String
    uniqueName = baseName
    Dim count As Integer
    count = 1
    While SheetExists(uniqueName)
        uniqueName = baseName & count
        count = count + 1
    Wend
    CreateUniqueName = uniqueName ' �������߂�l�̐ݒ�
End Function
'*******************************************************************************
' �@�\�F�w�肳�ꂽ�V�[�g�������݂��邩�m�F
' �����FsheetName - �m�F����V�[�g��
' �ߒl�FBoolean - �V�[�g�����݂���ꍇTrue
'*******************************************************************************
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing ' �������߂�l�̐ݒ�
End Function

'*******************************************************************************
' �@�\�F�O���[�v�����ꂽ�f�[�^���f�o�b�O�E�B���h�E�ɏo��
' �����FgroupedData - �o�͂���f�[�^�I�u�W�F�N�g
'*******************************************************************************
Private Sub PrintGroupedData(groupedData As Object)
    Dim key As Variant, item As HelmetData
    For Each key In groupedData.Keys
        Debug.Print "GroupNumber: " & key
        For Each item In groupedData(key)
            Debug.Print "  ProductName: " & item.productName
            Debug.Print "  ImpactPosition: " & item.ImpactPosition
            Debug.Print "  ImpactTemp: " & item.ImpactTemp
            Debug.Print "  Anvil: " & item.anvilForm
            Debug.Print "  Head: " & item.headModel
            Debug.Print "----------------------------"
        Next item
        Debug.Print "============================"
    Next key
End Sub
'*******************************************************************************
' �@�\�F�R�s�[�����V�[�g����CopiedSheetNames�V�[�g�ɕۑ�
' �����FsheetNames - �ۑ�����V�[�g���̃R���N�V����
'*******************************************************************************
Sub SaveCopiedSheetNames(sheetNames As Collection)
' SetupInspectionReport�̃T�u�v���V�[�W��
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "CopiedSheetNames"
    End If

    ws.Cells.ClearContents

    Dim i As Long
    For i = 1 To sheetNames.count
        ws.Cells(i, 1).value = sheetNames(i)
    Next i
End Sub
'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�F�����L�^�̓]�L�Ɠ��背�R�[�h�̈ړ������s
' �����F�Ȃ�
'*******************************************************************************
Sub ManageInspectionRecords()
    Call TransferDataToInspectionReports
    Call MoveSpecificRecords
End Sub

'*******************************************************************************
' �@�\�FLOG_Bicycle�V�[�g�̃f�[�^�������񍐏��ɓ]�L
' �����F�Ȃ�
'*******************************************************************************
Sub TransferDataToInspectionReports()
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    Dim wsTarget As Worksheet
    Dim i As Long
    Dim productNameKey As String
    Dim dataRange As Range
    Dim targetRow As Long

    ' LOG_Helmet�V�[�g�̊e�s�����[�v���ď������܂�
    For i = 2 To lastRow
        ' GroupNumber��ProductName����productNameKey���\�z���܂�
        Dim parts() As String
        parts = Split(wsSource.Cells(i, "B").value, "-")
        productNameKey = parts(1) & "-" & parts(0)
        Dim productName As String
        productName = parts(1) ' "500" �Ȃ�

        Dim sheetIndex As Long
        Dim numericPart As Long
        numericPart = CLng(Split(productNameKey, "-")(1))

        If numericPart >= 1 And numericPart <= 6 Then ' ���l������1����6�̏ꍇ
            ' �V�[�g�C���f�b�N�X���v�Z (productName-4 ���܂�)
            Select Case numericPart
                Case 1: sheetIndex = 1
                Case 2, 3: sheetIndex = 2
                Case 4, 5, 6: sheetIndex = 3 ' productName-4 �� productName_3 �ɓ]�L
            End Select

            Dim targetSheetName As String
            targetSheetName = productName & "_" & sheetIndex

            TransferData productName, sheetIndex, i

        Else ' ���l������1����6�ȊO�̏ꍇ�͓]�L���Ȃ�
            Debug.Print "productNameKey: " & productNameKey & " �͔͈͊O�̂��ߓ]�L����܂���B"
        End If
    Next i
End Sub
'*******************************************************************************
' �@�\�F�w�肳�ꂽ���i���ƃC���f�b�N�X�Ɋ�Â��f�[�^��]�L
' �����FproductName - ���i��
'       sheetIndex - �V�[�g�C���f�b�N�X
'       sourceRow - �]�L���̍s�ԍ�
'*******************************************************************************
Private Sub TransferData(productName As String, sheetIndex As Long, sourceRow As Long)
    Dim targetSheetName As String
    targetSheetName = productName & "_" & sheetIndex

    On Error Resume Next
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0

    If Not wsTarget Is Nothing Then
        ' �^�[�Q�b�g�V�[�g�Ƀw�b�_�[��]�L���鏈��
        If wsTarget.Range("B30").value = "" Then ' �w�b�_�[�����]�L�ł���Γ]�L
            ThisWorkbook.Sheets("LOG_Bicycle").Range("B1:Z1").Copy Destination:=wsTarget.Range("B30")
        End If

        ' �ŏI�s�������A���̍s����f�[�^�̓]�L���J�n���܂�
        targetRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row + 1
        If targetRow < 31 Then
            targetRow = 31 ' �ŏ��̃f�[�^�]�L�J�n�ʒu��B31�ɐݒ�
        End If

        ThisWorkbook.Sheets("LOG_Bicycle").Range("B" & sourceRow & ":Z" & sourceRow).Copy Destination:=wsTarget.Range("B" & targetRow)

        Set wsTarget = Nothing ' wsTarget �����Z�b�g
    End If
End Sub

'*******************************************************************************
' �@�\�F����ID��'4�����f�[�^��_3�V�[�g����_2�V�[�g�Ɉړ�
' �����F�Ȃ�
'*******************************************************************************
Sub MoveSpecificRecords()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim productName As String
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim headerRow As Long
    Dim sampleIDColumn As Long, testLocationColumn As Long
    Dim i As Long, j As Long ' ���[�v�J�E���^ j ��ǉ�

    headerRow = 30 ' �w�b�_�[�s

    ' �e�V�[�g�����[�v���� (productName_3 �V�[�g��ΏۂƂ���)
    For Each ws In ThisWorkbook.Worksheets
        If Right(ws.Name, 2) = "_3" Then ' �V�[�g���� "_3" �ŏI���V�[�g�̂ݏ���
            productName = Left(ws.Name, Len(ws.Name) - 2) ' productName ���擾

            ' �V�[�g�̗L�����m�F
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Worksheets(productName & "_2")
            Set wsSource = ThisWorkbook.Worksheets(productName & "_3") ' wsSource �������Őݒ�
            On Error GoTo 0
            If wsTarget Is Nothing Then
                MsgBox productName & "_2 �V�[�g�����݂��܂���B", vbCritical
                Exit Sub
            End If
            If wsSource Is Nothing Then
                MsgBox productName & "_3 �V�[�g�����݂��܂���B", vbCritical
                Exit Sub
            End If

            ' "����ID" �� "�����ӏ�" �̗�ԍ����擾
            For j = 1 To wsSource.Cells(headerRow, Columns.count).End(xlToLeft).Column
                If wsSource.Cells(headerRow, j).value = "����ID" Then
                    sampleIDColumn = j
                ElseIf wsSource.Cells(headerRow, j).value = "�����ӏ�" Then
                    testLocationColumn = j
                End If
            Next j

            If sampleIDColumn = 0 Or testLocationColumn = 0 Then
                MsgBox "�u����ID�v�܂��́u�����ӏ��v�̃w�b�_�[��������܂���B", vbCritical
                Exit Sub
            End If

            lastRowSource = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row
            lastRowTarget = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).Row

            ' �]�L���V�[�g�̃f�[�^�����[�v���� (�������Ƀ��[�v)
            For i = lastRowSource To headerRow + 1 Step -1
                If wsSource.Cells(i, sampleIDColumn).value = 4 And _
                   (wsSource.Cells(i, testLocationColumn).value = "�O����" Or wsSource.Cells(i, testLocationColumn).value = "�㓪��") Then

                    ' �f�[�^��]�L��V�[�g�ɃR�s�[
                    wsSource.Rows(i).EntireRow.Copy Destination:=wsTarget.Rows(lastRowTarget + 1)
                    ' �]�L��V�[�g�̍ŏI�s���X�V
                    lastRowTarget = lastRowTarget + 1
                    wsSource.Rows(i).Delete
                End If
            Next i
        End If
    Next ws
End Sub



Sub CustomizeReportProcess()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long, checkRow As Long
    Dim sourceData As String
    Dim parts() As String
    Dim baseSheetName As String
    Dim ws As Worksheet
    Dim foundSheets As Collection
    Dim targetSheet As Worksheet
    Dim isValidData As Boolean
    
    ' �G���[�n���h�����O�̐ݒ�
    On Error GoTo ErrorHandler
    
    ' �\�[�X�V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "D").End(xlUp).Row
    
    ' Excel�̃p�t�H�[�}���X����̂��߂̐ݒ�
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' B��̃f�[�^���擾���ď���
    sourceData = wsSource.Cells(2, "B").value
    parts = Split(sourceData, "-")
    
    ' �G���[�`�F�b�N�F�f�[�^�`���̊m�F
    If UBound(parts) < 4 Then
        MsgBox "�f�[�^�`�����s���ł�: " & sourceData, vbExclamation
        GoTo CleanExit
    End If
    
    ' D��̃f�[�^����
    isValidData = True
    For checkRow = 2 To lastRow
        If wsSource.Cells(checkRow, "D").value <> parts(1) Then
            isValidData = False
            Exit For
        End If
    Next checkRow
    
    ' �f�[�^���،��ʂ̊m�F
    If Not isValidData Then
        MsgBox "�G���[: D��ɈقȂ�l�����݂��܂��B�����𒆎~���܂��B" & vbCrLf & _
               "���Ғl: " & parts(1) & vbCrLf & _
               "�m�F�s: " & checkRow, vbCritical
        GoTo CleanExit
    End If
    
    ' �V�[�g���̃x�[�X�����𐶐�
    baseSheetName = parts(1) & "_1"
    
    ' �Y������V�[�g��T��
    Set foundSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, baseSheetName) > 0 Then
            foundSheets.Add ws
        End If
    Next ws
    
    ' ���������V�[�g���Ȃ��ꍇ�̏���
    If foundSheets.count = 0 Then
        MsgBox "�x��: " & baseSheetName & " �ɊY������V�[�g��������܂���B", vbExclamation
        GoTo CleanExit
    End If
    
    ' ���������e�V�[�g�ɑ΂��ď��������s
    For Each targetSheet In foundSheets
        ' �f�[�^�̓]�L����
        With targetSheet
            .Range("D3").value = wsSource.Cells(2, "D").value
            .Range("D4").value = wsSource.Cells(2, "O").value
            .Range("D5").value = wsSource.Cells(2, "E").value
            .Range("D6").value = wsSource.Cells(2, "Q").value
            .Range("I3").value = wsSource.Cells(2, "F").value
            .Range("I4").value = wsSource.Cells(2, "G").value
        End With
    Next targetSheet
    
CleanExit:
    ' Excel�̐ݒ�����ɖ߂�
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    ' �G���[�������̏���
    MsgBox "�G���[���������܂����B" & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
           "�G���[���e: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

