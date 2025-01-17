Attribute VB_Name = "Test"
Sub CheckAndMarkRecords()
    Dim wsSource As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, checkRow As Long
    Dim targetLastRow As Long
    Dim foundSheets As Collection
    Dim PLNum As String
    Dim hasFailedRecord As Boolean
    Dim failedRow As Range
    Dim clearRange As Range
    Dim isAllPass As Boolean    ' �S�V�[�g���i�t���O��ǉ�
    
    ' �G���[�n���h�����O�̐ݒ�
    On Error GoTo ErrorHandler
    
    ' �\�[�X�V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "D").End(xlUp).Row
    
    ' Excel�̃p�t�H�[�}���X����̂��߂̐ݒ�
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' �Y������V�[�g��T�����邽�߂̃R���N�V�����쐬
    Set foundSheets = New Collection
    
    ' D�񂩂�ΏۃV�[�g��T��
    For checkRow = 2 To lastRow
        PLNum = wsSource.Cells(checkRow, "D").value
        
        ' ���[�N�u�b�N���̑S�V�[�g���`�F�b�N
        For Each ws In ThisWorkbook.Worksheets
            ' �V�[�g���� "PLNum_����" �̌`���ƈ�v���邩�`�F�b�N
            If ws.Name Like PLNum & "_[0-9]*" Then
                ' �d��������邽�߁A���ɒǉ�����Ă��Ȃ����`�F�b�N
                On Error Resume Next
                foundSheets.Add ws, ws.Name
                On Error GoTo ErrorHandler
            End If
        Next ws
    Next checkRow
    
    ' ���������V�[�g���Ȃ��ꍇ�̏���
    If foundSheets.count = 0 Then
        MsgBox "�ΏۂƂȂ�V�[�g��������܂���B", vbExclamation
        GoTo CleanExit
    End If
    
    ' �S�V�[�g���i�t���O��������
    isAllPass = True
    
    ' �e�V�[�g�ɑ΂��ď��������s
    For Each ws In foundSheets
        hasFailedRecord = False
        targetLastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
        
        ' ������"�s���i"�s���폜
        On Error Resume Next
        ws.Rows(targetLastRow + 1).Delete
        On Error GoTo ErrorHandler
        
        ' �����̐F�t�����N���A
        Set clearRange = ws.Range(ws.Cells(30, "B"), ws.Cells(targetLastRow, "U"))
        clearRange.Interior.ColorIndex = xlNone
        
        ' H18�Z���̓��e���N���A
        ws.Range("H18").value = ""
        
        ' 30�s�ڂ���ŏI�s�܂Ń`�F�b�N
        For checkRow = 30 To targetLastRow
            ' D��̒l��"PLNum"�ƈ�v���郌�R�[�h���`�F�b�N
            If ws.Cells(checkRow, "D").value = PLNum Then
                ' J���L��̒l���擾
                Dim jValue As Variant
                Dim lValue As Variant
                
                jValue = ws.Cells(checkRow, "J").value
                lValue = ws.Cells(checkRow, "L").value
                
                ' J��̐��l�`�F�b�N�Ə�������
                If Not IsNumeric(jValue) Then
                    jValue = 0
                End If
                
                ' L��̐��l�`�F�b�N�Ə�������
                If Not IsNumeric(lValue) Then
                    lValue = 0
                End If
                
                ' �����`�F�b�N
                If CDbl(jValue) >= 300 Or CDbl(lValue) >= 4 Then
                    ' B�񂩂�U���F�t��
                    ws.Range(ws.Cells(checkRow, "B"), _
                            ws.Cells(checkRow, "U")).Interior.Color = RGB(255, 153, 153)
                    hasFailedRecord = True
                End If
            End If
        Next checkRow
        
        ' �����𖞂������R�[�h��1�ł��������ꍇ�A�s���i�����
        If hasFailedRecord Then
            isAllPass = False   ' �s���i���������ꍇ�A�S�V�[�g���i�t���O��false��
            
            ' �ŏI�s���Ď擾
            targetLastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
            
            ' �s���i�s�̐ݒ�
            Set failedRow = ws.Range(ws.Cells(targetLastRow + 1, "A"), _
                                   ws.Cells(targetLastRow + 1, "U"))
            
            With failedRow
                ' �Z���̌���
                .Merge
                ' �s���i�e�L�X�g�̓���
                .value = "�s���i �� J��300�ȏ� �܂��� L��4�ȏ� �̃��R�[�h�����݂��܂�"
                ' �Z���̏����ݒ�
                With .Interior
                    .Color = RGB(255, 153, 153)
                End With
                With .Font
                    .Bold = True
                    .Size = 12
                    .Color = RGB(192, 0, 0)
                End With
                ' �z�u�ݒ�
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' �s�̍����𒲐�
            ws.Rows(targetLastRow + 1).RowHeight = 25
        End If
    Next ws
    
    ' ���ׂẴV�[�g�̏������I�������A�S�V�[�g���i�Ȃ獇�i��\��
    If isAllPass Then
        For Each ws In foundSheets
            With ws.Range("H18")
                .value = "���i"
                With .Font
                    .Bold = True
                    .Size = 12
                    .Color = RGB(0, 176, 80)  ' �ΐF
                End With
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next ws
    End If
    
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


' productName_1�̃V�[�g�̑̍ق𐮂���B
Sub CustomizeReportIntroduction()

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
                    .Size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "�Z����" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .Name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub



