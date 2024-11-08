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
' �����̃V�[�g���R�s�[���AproductName_1 �Ȃǂ̖��O������B
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
                sheetName = Array("InspectionSheet01", "InspectionSheet02", "InspectionSheet03")(sheetIndex - 1)
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
Function ParseHelmetData(value As String) As HelmetData
' SetupInspectionReport�̃T�u�v���V�[�W��
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

Function CreateUniqueName(baseName As String) As String
' SetupInspectionReport�̃T�u�v���V�[�W��
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
Function SheetExists(sheetName As String) As Boolean
' SetupInspectionReport�̃T�u�v���V�[�W��
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing ' �������߂�l�̐ݒ�
End Function
Private Sub PrintGroupedData(groupedData As Object)
' SetupInspectionReport�̃T�u�v���V�[�W��
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

' �R�s�[�����V�[�g�Ƀw�b�_�[�Ǝ������ʂ�]�L����B
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

Private Sub TransferData(productName As String, sheetIndex As Long, sourceRow As Long)
' TransferDataToInspectionReports�̃T�u�v���V�[�W���B�f�[�^�]�L�������֐���
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

' _4�̃f�[�^�݂̂�]�L����
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






    
    
    '�V�������݂̂̃V�[�g���쐬����B
Sub TransferDataToTopImpactTest()
    '"Log_Helmet"����R�s�[���������[�ɒl��]�L����B
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim firstDashPos As Integer
    Dim secondDashPos As Integer
    Dim matchName As String
    Dim TemperatureCondition As String

    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("Log_Bicycle")

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' C��̒l���琻�i�R�[�h���擾
        firstDashPos = InStr(wsSource.Cells(i, "B").value, "-")
        If firstDashPos > 0 Then
            secondDashPos = InStr(firstDashPos + 1, wsSource.Cells(i, "B").value, "-")
            If secondDashPos > 0 Then
                matchName = Left(wsSource.Cells(i, "B").value, secondDashPos - 1)
            End If
        End If

        ' �e�V�[�g�����[�v���ď����Ɉ�v����V�[�g������
        For Each wsDestination In ThisWorkbook.Sheets
            If wsDestination.Name = matchName Then ' �V�[�g�������i�R�[�h�Ɉ�v���邩�m�F
                ' �����Ɉ�v�����ꍇ�A�]�L�����s
                ' �ȉ��̃R�[�h�͕ύX�Ȃ�
                wsDestination.Range("C2").value = wsSource.Cells(i, 21).value
                wsDestination.Range("F2").value = wsSource.Cells(i, 6).value
                wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                wsDestination.Range("A10").value = "���O�����F" & wsSource.Cells(i, 12).value
                wsDestination.Range("A14").value = "�����ΏۊO"
                wsDestination.Range("A19").value = "�����ΏۊO"
                Exit For ' �]�L��͎��̍s��
            End If
        Next wsDestination
    Next i
End Sub

' productName_1�̃V�[�g�ɓ]�L����B
Sub TransferDataToDynamicSheets()

    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim destinationSheetName As String

    ' �\�[�X�V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row
    
    ' Excel�̃p�t�H�[�}���X����̂��߂̐ݒ�
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSource��C������[�v���ăf�[�^������
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, "B").value
        checkData = wsSource.Cells(i, 5).value
        parts = Split(sourceData, "-")

        ' �V�[�g���̐���
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(0) & "-" & parts(1)

            ' �]�L��V�[�g�̑��݊m�F
            On Error Resume Next
            Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
            On Error GoTo 0

            ' �V�[�g�����݂��A����������v����ꍇ�Ƀf�[�^��]�L
            If Not wsDestination Is Nothing Then
                Select Case parts(2)
                    Case "�V"
                        If checkData = "�V��" Then
                            ' �V�Ɋւ���f�[�^�]�L
                            wsDestination.Range("C2").value = wsSource.Cells(i, 21).value
                            wsDestination.Range("F2").value = wsSource.Cells(i, 6).value
                            wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                            wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                            wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                            wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                            wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                            wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                            wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                            wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                            wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                            wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("A10").value = "���O�����F" & wsSource.Cells(i, 12).value
                        End If
                    Case "�O"
                        If checkData = "�O����" Then
                            ' �O�����Ɋւ���f�[�^�]�L
                            wsDestination.Range("E13").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E14").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E15").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A13").value = "�O����"
                        End If
                    Case "��"
                        If checkData = "�㓪��" Then
                            ' �㓪���Ɋւ���f�[�^�]�L
                            wsDestination.Range("E17").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("E18").value = wsSource.Cells(i, 10).value
                            wsDestination.Range("E19").value = wsSource.Cells(i, 11).value
                            wsDestination.Range("A17").value = "�㓪��"
                        End If
                End Select
            End If
        End If
    Next i
    
    ' Excel�̐ݒ�����ɖ߂�
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub ImpactValueJudgement()
    'CopiedSheetNames�V�[�g��A��Ɋ�Â��Ċe�����[�V�[�g�̏Ռ��l�𔻒肷��
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim resultE11 As Boolean, resultE14 As Boolean, resultE19 As Boolean
    Dim targetSheets As Collection
    
    ' ��������V�[�g�����擾
    Set targetSheets = GetTargetSheetNames()
    
    ' �Ώۂ̃V�[�g���Ɋ�Â��ď������s��
    For i = 1 To targetSheets.count
        sheetName = targetSheets(i)
        ' �Ώۂ̃V�[�g��ݒ�
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        
        ' D11, D14, D19�̒l����ɔ���
        resultE11 = wsTarget.Range("E11").value <= 4.9
        resultE14 = IsEmpty(wsTarget.Range("E13")) Or wsTarget.Range("E13").value <= 9.81
        resultE19 = IsEmpty(wsTarget.Range("E17")) Or wsTarget.Range("E17").value <= 9.81
        
        ' �S�Ă̏�����True�̏ꍇ��"���i"�A����ȊO��"�s���i"��G9�ɋL��
        If resultE11 And resultE14 And resultE19 Then
            wsTarget.Range("H9").value = "���i"
        Else
            wsTarget.Range("H9").value = "�s���i"
        End If
    Next i
End Sub

Function GetTargetSheetNames() As Collection
    ' CopiedSheetNames�V�[�g��A�񂩂�V�[�g�����擾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetNames As New Collection
    
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        sheetNames.Add ws.Cells(i, 1).value
    Next i
    
    Set GetTargetSheetNames = sheetNames
End Function
    ' CopiedSheetNames�V�[�g��A��Ɋ�Â��Č����[�ɏ�����ݒ肷��
Sub FormatNonContinuousCells()
    Dim wsTarget As Worksheet
    Dim i As Long
    Dim sheetName As String
    Dim targetSheets As Collection
    Dim rng As Range
    Dim cell As Range
    
    ' ��������V�[�g�����擾
    Set targetSheets = GetTargetSheetNames()
    
    ' �Ώۂ̃V�[�g���Ɋ�Â��ď������s��
    For i = 1 To targetSheets.count
        sheetName = targetSheets(i)
        
        ' ���[�N�V�[�g�����݂��邩�`�F�b�N
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        ' ���[�N�V�[�g�����݂���΁A�w�肵���Z���͈͂ɏ�����ݒ�
        If Not wsTarget Is Nothing Then
            ' �͈͂Ə����ݒ���֘A�t��
            FormatRange wsTarget.Range("E7"), "������", 12, True
            FormatRange wsTarget.Range("E8"), "������", 12, True
            FormatRange wsTarget.Range("E9"), "������", 12, True

            ' E13�ɒl���Ȃ��ꍇ�AA14:E14��B15:D16���O���[�A�E�g
            If IsEmpty(wsTarget.Range("E13").value) Then
                wsTarget.Range("A13").value = "�����ΏۊO"
                FormatRange wsTarget.Range("A13"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B13:F13, B14:E15"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A13"), "���S�V�b�N", 12, True
                FormatRange wsTarget.Range("E13:E15"), "���S�V�b�N", 10, False, RGB(255, 255, 255)
            End If

            ' E17�ɒl���Ȃ��ꍇ�AA19:E19��B20:D21���O���[�A�E�g
            If IsEmpty(wsTarget.Range("E17").value) Then
                wsTarget.Range("A17").value = "�����ΏۊO"
                FormatRange wsTarget.Range("A17"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B17:F17, B18:E19"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A17"), "���S�V�b�N", 12, True
                FormatRange wsTarget.Range("E17:E19"), "���S�V�b�N", 10, False, RGB(255, 255, 255)
            End If
            
            ' ����̕����ɏ�����K�p
            FormatSpecificEndStrings wsTarget.Range("A10"), "���S�V�b�N", 12, True
            
            ' �Z���̏����ݒ�
            With wsTarget.Range("C2:C4, F2:F4, H2:H4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("F3").NumberFormat = "0.0"" g"""
            wsTarget.Range("H2").NumberFormat = "0"" ��"""
            wsTarget.Range("H3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("E11, E14, E19").NumberFormat = "0.00"" kN"""
            
            ' E14:E15, E18:E19�̒l�ɉ����ď�����ݒ�
            Set rng = wsTarget.Range("E14:E15, E18:E19")
            For Each cell In rng
                If cell.value <= 0.01 Then
                    cell.value = "�\"
                Else
                    cell.NumberFormat = "0.00"" ms"""
                End If
            Next cell
            
            ' ���͈̔͂����l�ɐݒ�\
            ' FormatRange wsTarget.Range("���̑��͈̔�"), "�t�H���g��", �t�H���g�T�C�Y, �������ǂ���, �w�i�F

            Set wsTarget = Nothing
        End If
    Next i
End Sub


Sub FormatSpecificEndStrings(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean)
    ' �Z���̓���̕���(�O����)�ɏ�����K�p����T�u�v���V�[�W��
    Dim cell As Range

    For Each cell In rng
        Dim text As String
        text = cell.value
        Dim textLength As Integer
        textLength = Len(text)

        If textLength >= 2 Then
            If Right(text, 2) = "����" Or Right(text, 2) = "�ቷ" Then
                With cell.Characters(Start:=textLength - 1, Length:=2).Font
                    .Name = fontName
                    .size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "�Z����" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .Name = fontName
                    .size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub

Sub FormatRange(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean, Optional bgColor As Variant)
    ' �͈͂ɏ�����K�p���邽�߂̃T�u�v���V�[�W��
    With rng
        .Font.Name = fontName
        .Font.size = fontSize
        .Font.Bold = isBold
        If Not IsMissing(bgColor) Then
            .Interior.Color = bgColor
        Else
            .Interior.colorIndex = xlColorIndexAutomatic ' �w�i�F�������ɐݒ�
        End If
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub
' �`���[�g���e�V�[�g�ɕ��z����B
Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"�V�[�g��Ώۂɂ���
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"�V�[�g�̃`���[�g�I�u�W�F�N�g���O���[�v����
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartName��"-"�ŕ������AsheetName���擾
        parts = Split(chartObj.Name, "-")
        If UBound(parts) >= 1 Then
            sheetName = parts(0) & "-" & parts(1)
        Else
            sheetName = parts(0)
        End If
        
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' �O���[�v���ƂɃ`���[�g��Ή�����V�[�g�Ɉړ�
    Dim key As Variant
    For Each key In groups.Keys
        ' �V�[�g�̑��݂��m�F
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' �V�[�g�����݂��Ȃ��ꍇ�A�`���[�g���ړ����Ȃ�
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' �`���[�g�̈ړ�
            Dim chart As ChartObject
            For Each chart In groups(key)
                chart.chart.Location Where:=xlLocationAsObject, Name:=targetSheet.Name
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not moved."
        End If
    Next key
End Sub
