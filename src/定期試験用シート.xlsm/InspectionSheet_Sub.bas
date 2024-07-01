Attribute VB_Name = "InspectionSheet_Sub"


Sub MakeInspectionSheets()
    'Call CreateInspectionSheetIDs
    Call DuplicateAndRenameSheets
    Call TransferDataToTopImpactTest
    Call RenameAndRemoveDuplicateSheets
    Call TransferDataToDynamicSheets
    Call ImpactValueJudgement
    Call FormatNonContinuousCells
    MsgBox "�����[�V�[�g�̍쐬���I�����܂���"
End Sub


Sub DuplicateAndRenameSheets()
    Dim wsLogHelmet As Worksheet, wsTemplate As Worksheet, wsDraft As Worksheet
    Dim i As Long
    Dim part1Result As Boolean
    Dim sheetName As String, value As String, part1 As String, part2 As String

    Const LOG_HELMET As String = "Log_Helmet"
    Const TEMPLATE_SHEET As String = "InspectionSheet"

    Set wsLogHelmet = ThisWorkbook.Sheets(LOG_HELMET)
    Set wsTemplate = ThisWorkbook.Sheets(TEMPLATE_SHEET)

    ' �V�[�g�̕����Ɩ��O�̐ݒ�
    For i = 2 To wsLogHelmet.Cells(wsLogHelmet.Rows.count, 2).End(xlUp).row
        value = wsLogHelmet.Cells(i, 2).value
        
        ' ������ɃL���X�g���Ĉ��S�Ɋ֐��ɓn��
        part1 = CStr(Split(value, "-")(1))
        part2 = CStr(Split(value, "-")(2))
        part1Result = CheckPart1(part1)
        sheetName = ExtractSheetName(value)

        ' �f�o�b�O���̏o��
        Debug.Print "Row: " & i & ", Value: " & value & ", Part1: " & part1 & ", Part2: " & part2
        Debug.Print "Part1Result: " & part1Result & ", SheetName: " & sheetName
        Debug.Print "Should Duplicate: " & (part1Result Or (Not part1Result And part2 = "�V"))
        
        ' part1Result��True���AFalse�ł�part2��"�V"�̏ꍇ�ɃV�[�g�𕡐�
        If part1Result Or (Not part1Result And part2 = "�V") Then
            If sheetName <> "" Then
                wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                Set wsDraft = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                If Not SheetExists(sheetName) Then
                    wsDraft.name = sheetName
                End If
                Debug.Print "Sheet Duplicated: " & sheetName
            Else
                Debug.Print "Sheet not duplicated due to empty sheet name."
            End If
        Else
            Debug.Print "Conditions not met for duplication."
        End If
    Next i
End Sub



Function ExtractSheetName(fullName As String) As String
    Dim parts As Variant
    parts = Split(fullName, "-")
    
    If UBound(parts) >= 2 Then
        ' CheckPart1�̌��ʂɊւ�炸�Apart(2)��"�V"�̏ꍇ�̓V�[�g���𐶐�
        If parts(2) = "�V" Then
            ' parts(1)��"F"���܂ޏꍇ�͂��̕����������ăV�[�g���𐶐�
            Dim cleanPart1 As String
            cleanPart1 = Replace(parts(1), "F", "")  ' "300F"����"F"���폜
            ExtractSheetName = parts(0) & "-" & cleanPart1 & "-" & parts(2)
        Else
            ExtractSheetName = ""  ' �����ɍ��v���Ȃ��ꍇ�͋󕶎����Ԃ�
        End If
    Else
        ExtractSheetName = ""  ' �K�؂ȃp�[�c���Ȃ��ꍇ�͋󕶎����Ԃ�
    End If
End Function



Function CheckPart1(part As String) As Boolean
    ' ������̖�����F�łȂ����True�AF�ł����False��Ԃ�
    CheckPart1 = Not Right(part, 1) = "F"
End Function



Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function

Sub DeleteSheet()
    Dim ws As Worksheet
    On Error Resume Next ' �G���[�����������ꍇ�A���̍s�֐i��
    Set ws = ThisWorkbook.Sheets("ID")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' �폜�̊m�F���b�Z�[�W��\�����Ȃ�
        ws.Delete
        Application.DisplayAlerts = True ' ���b�Z�[�W�\�������ɖ߂�
    End If
    On Error GoTo 0 ' �G���[�n���h�����O�����ɖ߂�
End Sub


Sub TransferDataToTopImpactTest()
    '�V�������݂̂̃V�[�g���쐬����B
    '"Log_Helmet"����R�s�[���������[�ɒl��]�L����B
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dashPosSource As Integer
    Dim dashPosDest As Integer
    Dim matchName As String
    Dim TemperatureCondition As String

    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("Log_Helmet")

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' C���1�����ڂ�"F"�łȂ��s��T��
        If Left(wsSource.Cells(i, 3).value, 1) <> "F" Then
            ' MatchName���擾�iC���1�����ڂ���"-"�܂Łj
            dashPosSource = InStr(wsSource.Cells(i, 3).value, "-")
            If dashPosSource > 0 Then
                matchName = Left(wsSource.Cells(i, 3).value, dashPosSource - 1)

                ' L��̒l�Ɋ�Â��ď�����ݒ�
                Select Case wsSource.Cells(i, 12).value
                    Case "����"
                        TemperatureCondition = "Hot"
                    Case "�ቷ"
                        TemperatureCondition = "Cold"
                    Case "�Z����"
                        TemperatureCondition = "Wet"
                    Case Else
                        TemperatureCondition = ""
                End Select

                ' ���[�N�V�[�g�̖��O�����[�v���ď������`�F�b�N
                For Each wsDestination In ThisWorkbook.Sheets
                    dashPosDest = InStr(wsDestination.name, "-")
                    If dashPosDest > 0 Then
                        If Left(wsDestination.name, dashPosDest - 1) = matchName And InStr(wsDestination.name, TemperatureCondition) > 0 Then
                            ' ���������Ă͂܂�����]�L
                            wsDestination.Range("C2").value = wsSource.Cells(i, 21).value '�������e
                            wsDestination.Range("F2").value = wsSource.Cells(i, 6).value '������
                            wsDestination.Range("H2").value = wsSource.Cells(i, 7).value
                            wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                            wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                            wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                            wsDestination.Range("C4").value = wsSource.Cells(i, 16).value 'Lot
                            wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                            wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                            wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                            wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                            wsDestination.Range("E11").value = wsSource.Cells(i, 8).value
                            wsDestination.Range("A10").value = "���O�����F" & wsSource.Cells(i, 12).value
                            wsDestination.Range("A14").value = "�����ΏۊO"
                            wsDestination.Range("A19").value = "�����ΏۊO"
                        End If
                    End If
                Next wsDestination
            End If
        End If
    Next i
End Sub




Sub RenameAndRemoveDuplicateSheets()
    '�t�B���^�����O���Ղ��悤�ɃV�[�g�������ρuF390F-Cold�v�̌`���ɂ���B
    Dim ws As Worksheet
    Dim parts() As String
    Dim newName As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' �d�����閼�O�����V�[�g����肵�A�폜
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.name, 1) = "F" Then
            parts = Split(ws.name, "-")
            If UBound(parts) >= 2 Then
                newName = parts(0) & "-" & parts(1)
                If dict.Exists(newName) Then
                    Application.DisplayAlerts = False
                    ws.Delete
                    Application.DisplayAlerts = True
                Else
                    dict.Add newName, newName
                End If
            End If
        End If
    Next ws

    ' �d�����폜������A�V�[�g����ύX
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.name, 1) = "F" Or InStr(ws.name, "-") > 0 Then
            parts = Split(ws.name, "-")
            If UBound(parts) >= 2 Then
                newName = parts(0) & "-" & parts(1)
                On Error Resume Next
                ws.name = newName
                On Error GoTo 0
            End If
        End If
    Next ws
End Sub

Sub TransferDataToDynamicSheets()
    'F�t���X�̂̎����[���쐬����B
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim modifiedSourceData As String
    Dim destinationSheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSource��C������[�v
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).value
        checkData = wsSource.Cells(i, 5).value
        parts = Split(sourceData, "-")

        If UBound(parts) >= 2 Then
            ' �f�[�^�ƃV�[�g�����쐬
            modifiedSourceData = parts(0) & "-" & parts(1)
            destinationSheetName = modifiedSourceData

            ' modifiedSourceData �� sourceData �̍ŏ���2�̕�������v����ꍇ�ɂ̂ݓ]�L
            If Left(sourceData, Len(modifiedSourceData)) = modifiedSourceData Then
                ' �V�[�g�����݂��邩�m�F���A���݂���ꍇ�̂ݓ]�L
                If InspectionSheetExists(destinationSheetName) Then
                    Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)

                    ' parts(UBound(parts))�Ɋ�Â��ď����𕪊�
                    Select Case parts(UBound(parts))
                        Case "�V"
                            If checkData = "�V��" Then
                                ' �V�Ɋւ���f�[�^�]�L�̏���
                                wsDestination.Range("C2").value = wsSource.Cells(i, 21).value '�������e
                                wsDestination.Range("F2").value = wsSource.Cells(i, 6).value '������
                                wsDestination.Range("H2").value = wsSource.Cells(i, 7).value '���x
                                wsDestination.Range("C3").value = "No." & wsSource.Cells(i, 4).value & "_" & wsSource.Cells(i, 15).value
                                wsDestination.Range("F3").value = wsSource.Cells(i, 13).value
                                wsDestination.Range("H3").value = wsSource.Cells(i, 14).value
                                wsDestination.Range("C4").value = wsSource.Cells(i, 16).value
                                wsDestination.Range("F4").value = wsSource.Cells(i, 17).value
                                wsDestination.Range("H4").value = wsSource.Cells(i, 18).value
                                wsDestination.Range("H7").value = wsSource.Cells(i, 19).value
                                wsDestination.Range("H8").value = wsSource.Cells(i, 20).value
                                wsDestination.Range("A10").value = "���O�����F" & wsSource.Cells(i, 12).value
                                wsDestination.Range("E11").value = wsSource.Cells(i, 8).value '�Ռ��l
                            End If

                        Case "�O"
                            If checkData = "�O����" Then
                                ' �O�Ɋւ���f�[�^�]�L�̏���
                                wsDestination.Range("E13").value = wsSource.Cells(i, 8).value '�Ռ��l
                                wsDestination.Range("E14").value = wsSource.Cells(i, 10).value '4.90kN
                                wsDestination.Range("E15").value = wsSource.Cells(i, 11).value '7.35kN
                                wsDestination.Range("A13").value = "�O����"
                            End If

                        Case "��"
                            If checkData = "�㓪��" Then
                                ' ��Ɋւ���f�[�^�]�L�̏���
                                wsDestination.Range("E17").value = wsSource.Cells(i, 8).value '�Ռ��l
                                wsDestination.Range("E18").value = wsSource.Cells(i, 10).value '4.90kN
                                wsDestination.Range("E19").value = wsSource.Cells(i, 11).value '7.35kN
                                wsDestination.Range("A17").value = "�㓪��"
                            End If

                        Case Else
                            ' ���̑��̒l�̏ꍇ�̏����i�K�v�ɉ����āj
                    End Select
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' �V�[�g�����݂��邩�ǂ������m�F����֐�
Function InspectionSheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    InspectionSheetExists = Not sheet Is Nothing
End Function


Sub ImpactValueJudgement()
    '�Ռ��z�������̌��ʂ��e�����[�V�[�g�̏Ռ��l���画�肷��B
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim resultE11 As Boolean, resultE14 As Boolean, resultE19 As Boolean

    ' "LOG_Helmet"�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' C��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' C���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        sheetName = wsSource.Cells(i, "C").value
        ' �Ώۂ̃V�[�g����ID�����킹�鏈��
        sheetName = Left(sheetName, Len(sheetName) - 2)

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


Sub FormatNonContinuousCells()
    ' �R�s�[���������[�ɏ�����ݒ肷��B
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmet�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' B��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    ' B��̊e�s�����[�v
    For i = 2 To lastRow
        sheetName = wsSource.Cells(i, 2).value
        ' �Ώۂ̃V�[�g����ID�����킹�鏈��
        sheetName = Left(sheetName, Len(sheetName) - 2)

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
                FormatRange wsTarget.Range("E13:E15"), "���S�V�b�N", 10, False, RGB(255, 255, 255) 'E13:E15�ɒ���
            End If

            ' E17�ɒl���Ȃ��ꍇ�AA19:E19��B20:D21���O���[�A�E�g
            If IsEmpty(wsTarget.Range("E17").value) Then
                wsTarget.Range("A17").value = "�����ΏۊO"
                FormatRange wsTarget.Range("A17"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
                FormatRange wsTarget.Range("B17:F17, B18:E19"), "���S�V�b�N", 10, False, RGB(242, 242, 242)
            Else
                FormatRange wsTarget.Range("A17"), "���S�V�b�N", 12, True
                FormatRange wsTarget.Range("E17:E19"), "���S�V�b�N", 10, False, RGB(255, 255, 255) 'E17:E19�ɒ���
            End If
            FormatSpecificEndStrings wsTarget.Range("A10"), "���S�V�b�N", 12, True '�O������ڗ�������_�����Ƃ��낪�Ȃ��̂ł����ɏ���
            With wsTarget.Range("C2:C4, F2:F4, H2:H4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsTarget.Range("F3").NumberFormat = "0.0"" g"""
            wsTarget.Range("H2").NumberFormat = "0"" ��"""
            wsTarget.Range("H3").NumberFormat = "0.0"" mm"""
            wsTarget.Range("E11, E14, E19").NumberFormat = "0.00"" kN"""
            wsTarget.Range("E14:E15, E18:E19").NumberFormat = "0.00"" ms"""
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
                    .name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            ElseIf textLength >= 3 And Right(text, 3) = "�Z����" Then
                With cell.Characters(Start:=textLength - 2, Length:=3).Font
                    .name = fontName
                    .Size = fontSize
                    .Bold = isBold
                End With
            End If
        End If
    Next cell
End Sub


' �͈͂ɏ�����K�p���邽�߂̃T�u�v���V�[�W��
Sub FormatRange(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean, Optional bgColor As Variant)
    With rng
        .Font.name = fontName
        .Font.Size = fontSize
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

' �Z���̓���̕�����̂�(�����ł͍����̂�)�ɏ�����K�p����v���V�[�W���Z�������Ђ炪�ȂɂȂ������߁AFormatSpecificText�ɂƂ��ĕς��ꂽ�B
Sub FormatLastTwoCharacters(rng As Range, fontName As String, fontSize As Integer, isBold As Boolean)
    Dim cell As Range
    Dim lastTwoChars As String

    For Each cell In rng
        If Len(cell.value) >= 2 Then
            lastTwoChars = Right(cell.value, 2)
            ' ������lastTwoChars�ɑ΂��ē���̏�����K�p����
            ' �������AVBA�ł͕����I�ȃZ���̏����ݒ�͒��ڂł��Ȃ����߁A
            ' ������S�̂ɏ�����K�p���A���̌�ōŌ��2���������ʂ̏�����K�p����
            With cell
                .Font.name = "���S�V�b�N"
                .Font.Size = 10
                .Font.Bold = False
                ' �Ō��2�����ɓ���̏�����K�p����
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.name = "���S�V�b�N"
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.Size = 12
                .Characters(Start:=Len(cell.value) - 1, Length:=2).Font.Bold = True
            End With
        End If
    Next cell
End Sub


Sub PrintFirstPageOfUniqueListedSheets()
    ' �w�肳�ꂽ�����[��1�y�[�W�ڂ��A�d���Ȃ�1�񂸂������v���V�[�W��
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim printedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set printedSheets = New Collection ' ������ꂽ�V�[�g����ǐՂ���R���N�V����

    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 2).value

        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        On Error Resume Next
        ' �R���N�V�����ɓ������O�����ɑ��݂��邩�`�F�b�N
        printedSheets.Add sheetName, sheetName
        If Err.number = 0 Then ' �ǉ������������ꍇ�A�V�[�g�͂܂��������Ă��Ȃ�
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' �V�[�g��1�y�[�W�ڂ݂̂����
            End If
        End If
        On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g

        Set wsTarget = Nothing
    Next i
End Sub



Sub ModifyAndStoreChartTitles()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartTitles() As String
    Dim i As Integer
    Dim parts() As String
    Dim modifiedChartTitle As String

    Set ws = ThisWorkbook.Sheets("LOG_Helmet") ' ���ۂ̃V�[�g���ɒu�������Ă�������

    ReDim chartTitles(1 To ws.ChartObjects.count)

    i = 1
    For Each chartObj In ws.ChartObjects
        ' �`���[�g�^�C�g����"-"�ŕ���
        parts = Split(chartObj.chart.ChartTitle.text, "-")

        ' �ŏ���2�̕�����g�ݍ��킹�ĐV�����^�C�g���𐶐�
        If UBound(parts) >= 1 Then
            modifiedChartTitle = parts(0) & "-" & parts(1)
        Else
            ' �����ł��Ȃ��ꍇ�͌��̃^�C�g�����g�p
            modifiedChartTitle = chartObj.chart.ChartTitle.text
        End If

        ' ���ό�̃^�C�g����z��Ɋi�[
        chartTitles(i) = modifiedChartTitle
        i = i + 1
    Next chartObj

    ' �e�X�g�o��
    For i = 1 To UBound(chartTitles)
        Debug.Print "Chart" & i & ": " & chartTitles(i)
    Next i
End Sub
Sub CopyChartToMatchingSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim chartObj As ChartObject
    Dim parts() As String
    Dim modifiedChartTitle As String
    Dim wb As Workbook
    
    Set wb = ThisWorkbook ' ���݂̃��[�N�u�b�N��ݒ�
    Set wsSource = wb.Sheets("LOG_Helmet") ' �\�[�X�V�[�g�����w��

    ' �\�[�X�V�[�g�̑S�`���[�g�����[�v
    For Each chartObj In wsSource.ChartObjects
        ' �`���[�g�^�C�g����"-"�ŕ���
        parts = Split(chartObj.chart.ChartTitle.text, "-")
        
        ' �ŏ���2�̕�����g�ݍ��킹�ĐV�����^�C�g���𐶐�
        If UBound(parts) >= 1 Then
            modifiedChartTitle = parts(0) & "-" & parts(1)
        Else
            ' �����ł��Ȃ��ꍇ�͌��̃^�C�g�����g�p
            modifiedChartTitle = chartObj.chart.ChartTitle.text
        End If
        
        ' ���[�N�u�b�N�̑S�V�[�g�����[�v
        For Each wsDest In wb.Sheets
            ' �`���[�g�̃^�C�g�����V�[�g���ƈ�v����ꍇ�A�`���[�g���R�s�[���y�[�X�g
            If wsDest.name = modifiedChartTitle Then
                ' �`���[�g���R�s�[
                ' �`���[�g���R�s�[
Dim tryCount As Integer
tryCount = 0
Do
    On Error Resume Next
    chartObj.Copy
    If Err.number = 0 Then Exit Do ' �R�s�[�ɐ��������烋�[�v�𔲂���
    On Error GoTo 0
    tryCount = tryCount + 1
    If tryCount > 5 Then ' 5�񎎍s���ă_���Ȃ�G���[���o��
        MsgBox "�`���[�g�̃R�s�[�Ɏ��s���܂���: " & chartObj.name
        Exit Sub
    End If
    Application.Wait Now + TimeValue("00:00:01") ' 1�b�҂��čĎ��s
Loop
                
                ' �V�[�g�Ƀy�[�X�g
                With wsDest
                    .Activate
                    .Paste
                    ' �\��t�����`���[�g�̈ʒu�𒲐��i��: A1�̈ʒu�ɔz�u�j
                    .Shapes(.Shapes.count).Top = .Range("A1").Top
                    .Shapes(.Shapes.count).Left = .Range("A1").Left
                End With
            End If
        Next wsDest
    Next chartObj
End Sub


'Sub ClearDataFromAllListedSheetsWithMergedCells()
'    '�]�L�������ڂ������v���V�[�W��
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim lastRow As Long
'    Dim i As Long
'    Dim sheetName As String
'
'    ' LOG_Helmet�V�[�g��ݒ�
'    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
'
'    ' B��̍ŏI�s���擾
'    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
'
'    ' B��̊e�s�����[�v
'    For i = 2 To lastRow
'        sheetName = wsSource.Cells(i, 2).value
'
'        If Left(sheetName, 1) = "F" Then
'            sheetName = Left(sheetName, Len(sheetName) - 2)
'        End If
'
'        ' ���[�N�V�[�g�����݂��邩�`�F�b�N
'        On Error Resume Next
'        Set wsTarget = ThisWorkbook.Sheets(sheetName)
'        On Error GoTo 0
'
'        ' ���[�N�V�[�g�����݂���΁A�w�肵�������Z������f�[�^���N���A
'        If Not wsTarget Is Nothing Then
'            ' �����Ō����Z���͈̔͂��w�肵�Ă�������
'            wsTarget.Range("C2:C4", "F2:F4", "H2:H4").ClearContents
'            wsTarget.Range("H7:H9").ClearContents
'            wsTarget.Range("E11:F11").ClearContents
'            wsTarget.Range("E13:E15").ClearContents
'            wsTarget.Range("F13").ClearContents
'            wsTarget.Range("E17:E19").ClearContents
'            wsTarget.Range("F17").ClearContents
'            wsTarget.Range("A10").ClearContents
'            ' �ȉ��A�K�v�Ȕ͈͂ɍ��킹�Ēǉ�
'
'            Set wsTarget = Nothing
'        End If
'    Next i
'End Sub


Sub DeleteAllListedSheets()
    ' �������ꂽ�����[���폜����v���V�[�W��
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' LOG_Helmet�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")

    ' B��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row

    ' B��̊e�s�����[�v
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 2).value

        If Left(sheetName, 1) = "F" Then
            sheetName = Left(sheetName, Len(sheetName) - 2)
        End If

        ' ���[�N�V�[�g�����݂��邩�`�F�b�N
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        If Not wsTarget Is Nothing Then
            Application.DisplayAlerts = False ' �x���̕\�����I�t�ɂ���
            wsTarget.Delete ' �V�[�g���폜
            Application.DisplayAlerts = True ' �x���̕\�����I���ɖ߂�
        End If
        On Error GoTo 0

        Set wsTarget = Nothing
    Next i
End Sub

