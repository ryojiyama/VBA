Attribute VB_Name = "SpecSheet"
 ' ���i�ԁA�����ӏ��Ȃǂɉ�����ID���쐬����
Sub CreateID()
   
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim ID As String

    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("Hel_SpecSheet")

    ' �Ō�̍s���擾
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' �e�s�ɑ΂���ID�𐶐�
    For i = 2 To lastRow ' 1�s�ڂ̓w�b�_�Ɖ���
        ID = GenerateID(ws, i)
        ' B���ID���Z�b�g
        ws.Cells(i, 2).value = ID
    Next i
End Sub

Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' CreateID()�̃T�u�v���V�[�W��
    Dim ID As String

    ' C��: 2���ȉ��̐���
    ID = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    ID = ID & "-" ' C���D��̊Ԃ�"-"
    ' D��̏�����ύX
    ID = ID & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    ID = ID & "-" ' Fm��E��̊Ԃ�"-"
    ID = ID & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E��̏���
    ID = ID & "-" ' Fm��E��̊Ԃ�"-
    ID = ID & GetColumnIValue(ws.Cells(rowIndex, 9).value) ' I��̏���
    ID = ID & "-" ' I���L��̊Ԃ�"-"
    ID = ID & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L��̏���

    GenerateID = ID
End Function
Function ExtractNumberWithF(value As String) As String
' GenerateID�̃T�u�֐�
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' ���K�\���I�u�W�F�N�g�̍쐬
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "�d{3,6}"
    regex.Global = True

    ' ���������𒊏o
    Set matches = regex.Execute(value)
    If matches.Count > 0 Then
        numPart = matches(0).value
    Else
        numPart = "000000" ' �f�t�H���g�l�܂��̓G���[�n���h�����O
    End If

    ' F�̑��݃`�F�b�N
    hasF = InStr(value, "F") > 0

    ' F������ꍇ�͐����̌��F������
    If hasF Then
        ExtractNumberWithF = numPart & "F"
    Else
        ExtractNumberWithF = numPart
    End If
End Function

Function GetColumnCValue(value As Variant) As String
' GenerateID�̃T�u�֐�
    If Len(value) <= 2 Then
        GetColumnCValue = Right("00" & value, 2)
    Else
        GetColumnCValue = "??"
    End If
End Function

Function GetColumnEValue(value As Variant) As String
    ' GenerateID�̃T�u�֐�
    If InStr(value, "�V��") > 0 Then
        GetColumnEValue = "�V"
    ElseIf InStr(value, "�O����") > 0 Then
        GetColumnEValue = "�O"
    ElseIf InStr(value, "�㓪��") > 0 Then
        GetColumnEValue = "��"
    ElseIf InStr(value, "������") > 0 Then
        Dim pos As Integer
        pos = InStr(value, "_")
        If pos > 0 Then
            GetColumnEValue = "��" & Mid(value, pos)
        Else
            GetColumnEValue = "��"
        End If
    Else
        GetColumnEValue = "?"
    End If
End Function



Function GetColumnIValue(value As Variant) As String
' GenerateID�̃T�u�֐�
    Select Case value
        Case "����"
            GetColumnIValue = "Hot"
        Case "�ቷ"
            GetColumnIValue = "Cold"
        Case "�Z����"
            GetColumnIValue = "Wet"
        Case Else
            GetColumnIValue = "?"
    End Select
End Function

Function GetColumnLValue(value As Variant) As String
' GenerateID�̃T�u�֐�
    If value = "��" Then
        GetColumnLValue = "White"
    Else
        GetColumnLValue = "OthClr"
    End If
End Function

' ��SpecSheet�ɓ]�L����v���V�[�W���̖{�́B�A�C�R���ɕR�Â��B
Sub SyncSpecSheetToLogHel()
    ' ���l�����������ꍇ�̓G���[���b�Z�[�W��\�����ď����𒆒f
    If HighlightDuplicateValues Then
        MsgBox "�Ռ��l�œ��l��������܂����B�����_���񌅂ɉe�����o�Ȃ��͈͂ŏC�����Ă��������B", vbCritical
        Exit Sub
    End If

    ' �\�ɋ󗓂�����ꍇ�ɃG���[���b�Z�[�W���o���Ē��f
    If Not LocateEmptySpaces Then
        MsgBox "�󗓂�����܂��B�܂��͂���𖄂߂Ă��������B", vbCritical
        Exit Sub
    End If

    Call ProcessSheetPairs          ' �]�L����������v���V�[�W��
    Call CustomizeSheetFormats      ' �e��ɏ����ݒ������
    Call TransformIDs               ' B���ID���쐬����B
    Call Utlities.FillBlanksWithHyphenInMultipleSheets
End Sub

Function HighlightDuplicateValues() As Boolean
    ' SyncSpecSheetToLogHel�̃T�u�v���V�[�W��
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundDuplicate As Boolean
    foundDuplicate = False ' ���l�������������ǂ����̃t���O��������

    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' �F�̃C���f�b�N�X��������
    Dim colorIndex As Integer
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�

    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            If ws.Cells(i, "H").value = ws.Cells(j, "H").value And ws.Cells(i, "H").value <> "" Then
                ' ���l�����Z�������������ꍇ�A�t���O��True�ɐݒ肵�A�Z���ɐF��h��
                foundDuplicate = True
                ws.Cells(i, "H").Interior.colorIndex = colorIndex
                ws.Cells(j, "H").Interior.colorIndex = colorIndex
                ws.Cells(i, "H").Interior.colorIndex = colorIndex ' ���l�����������Z���ɐF��h��
            End If
        Next j
        ' ���l�����������ꍇ�A���̐F�ɕύX
        If foundDuplicate And ws.Cells(i, "H").Interior.colorIndex <> xlNone Then
            colorIndex = colorIndex + 1
            ' �F�C���f�b�N�X�̍ő�l�𒴂��Ȃ��悤�Ƀ`�F�b�N
            If colorIndex > 56 Then colorIndex = 3 ' �F�C���f�b�N�X�����Z�b�g
        End If
    Next i

    ' ���l�����������Ȃ������ꍇ�AH��̃Z���̐F�𔒂ɐݒ�
    If Not foundDuplicate Then
        For i = 2 To lastRow
            ws.Cells(i, "H").Interior.Color = xlNone
        Next i
    End If

    ' ���l�������������ǂ����Ɋ�Â��Č��ʂ�Ԃ�
    HighlightDuplicateValues = foundDuplicate
End Function

Function LocateEmptySpaces() As Boolean
    ' SyncSpecSheetToLogHel�̃T�u�v���V�[�W��
    Dim sheetName As String
    sheetName = "Hel_SpecSheet"

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim cell As Range
    Dim errorMsg As String

    ' �G���[���b�Z�[�W�p�̕������������
    errorMsg = ""

    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' �ŏI���"M"(�����敪)�ɌŒ�
    Dim lastCol As Long
    lastCol = ws.Columns("M").Column

    ' �w��͈͂����[�v
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' �󔒂̃`�F�b�N
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "�󔒃Z��: " & cell.Address(False, False) & vbNewLine
            End If

            ' ��G�AH�AJ�AK�Ő��l�̊m�F
            If j = Columns("G").Column Or j = Columns("H").Column Or j = Columns("J").Column Or j = Columns("K").Column Then
                If Not IsNumeric(cell.value) Then
                    ' �Z���̏�����W���ɐݒ�
                    cell.NumberFormat = "General"

                    ' ���l�ɕϊ�
                    If IsNumeric(cell.value) Then
                        cell.value = CDbl(cell.value)
                    Else
                        cell.value = 0
                    End If
                    cell.Interior.colorIndex = 6 ' ���F�ɐF�t��
                    errorMsg = errorMsg & "���l�ɕϊ������Z��: " & cell.Address(False, False) & vbNewLine
                End If
                cell.NumberFormat = "General"
            End If

            ' ��N�AO�AP�ŕ�����̊m�F
            If j = Columns("N").Column Or j = Columns("O").Column Or j = Columns("P").Column Then
                If Not VarType(cell.value) = vbString Then
                    ' ������ɕϊ�
                    cell.value = CStr(cell.value)
                    cell.Interior.colorIndex = 6 ' ���F�ɐF�t��
                    errorMsg = errorMsg & "������ɕϊ������Z��: " & cell.Address(False, False) & vbNewLine
                End If
            End If
        Next j
    Next i

    ' �G���[���b�Z�[�W������Ε\�����AFalse��Ԃ�
    If Len(errorMsg) > 0 Then
        LocateEmptySpaces = False
        MsgBox errorMsg, vbCritical
    Else
        LocateEmptySpaces = True
    End If
End Function

' �]�L����������v���V�[�W��
Sub ProcessSheetPairs()
    Dim sheetPairs As Variant
    Dim logSheetName As String
    Dim specSheetName As String
    Dim pair As Variant

    ' �V�[�g�y�A���`
    sheetPairs = Array( _
        Array("LOG_Helmet", "Hel_SpecSheet"), _
        Array("LOG_FallArrest", "FallArr_SpecSheet"), _
        Array("LOG_Bicycle", "Bic_SpecSheet"), _
        Array("LOG_BaseBall", "Base_SpecSheet") _
    )

    ' �e�V�[�g�y�A��T�����ď���
    For Each pair In sheetPairs
        logSheetName = pair(0)
        specSheetName = pair(1)
'        Debug.Print logSheetName
'        Debug.Print specSheetName
        ' �V�[�g�y�A�����݂��邩�`�F�b�N
        If SheetExists(logSheetName) And SheetExists(specSheetName) Then
            ' �V�[�g�y�A�����������ꍇ�ɏ��������s
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
            Debug.Print logSheetName
        End If
    Next pair
End Sub

Function SheetExists(sheetName As String) As Boolean
    'ProcessSheetPairs�̃T�u�v���V�[�W��
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Sub CopyDataBasedOnCondition(sheetNameLog As String, sheetNameSpec As String)
    'ProcessSheetPairs�̃T�u�v���V�[�W��
    Dim logSheet As Worksheet
    Dim helSpec As Worksheet
    Dim lastRowLog As Long
    Dim lastRowSpec As Long
    Dim i As Long, j As Long
    Dim matchCount As Long
    Dim columnsToCopy As Collection
    Dim colPair As Variant
    Dim logHeader As Range
    Dim helSpecHeader As Range
    Dim col As Range
    Dim colLog As Range

    ' ���[�N�V�[�g���Z�b�g
    Set logSheet = ThisWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ThisWorkbook.Worksheets(sheetNameSpec)

    ' LOG�V�[�g�̍ŏI�s���擾
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).row
    ' Spec�V�[�g�̍ŏI�s���擾
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).row

    ' �w�b�_�[�s���擾
    Set logHeader = logSheet.Rows(1)
    Set helSpecHeader = helSpec.Rows(1)

    ' �]�L�����̃y�A���R���N�V�����ɒ�`
    Set columnsToCopy = New Collection

    ' �y�A�ƂȂ�w�b�_�[�����擾
    colPair = GetHeaderPairs(sheetNameLog, sheetNameSpec)

    ' �y�A���������擾����Ă��邩�m�F
    If UBound(colPair) = -1 Then
        MsgBox "�w�b�_�[�̃y�A��������܂���ł���: " & sheetNameLog & " �� " & sheetNameSpec
        Exit Sub
    End If

    ' �e�w�b�_�[�s�𑖍����Ĉ�v����w�b�_�[��������
    Dim pair As Variant
    For Each pair In colPair
        Dim logCol As Long
        Dim helSpecCol As Long
        logCol = 0
        helSpecCol = 0
        For Each col In logHeader.Cells
            If col.value = pair(0) Then
                logCol = col.Column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.value = pair(1) Then
                helSpecCol = col.Column
                Exit For
            End If
        Next col
        If logCol > 0 And helSpecCol > 0 Then
            columnsToCopy.Add Array(logCol, helSpecCol)
        Else
            MsgBox "�w�b�_�[��������܂���ł���: " & pair(0) & " �܂��� " & pair(1)
        End If
    Next pair

    ' �l���r���ē]�L
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").value = helSpec.Cells(j, "H").value Then
                ' H��̒l����v�����ꍇ�A�e��̓��e��]�L
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.Count
                    logSheet.Cells(i, columnsToCopy(k)(0)).value = helSpec.Cells(j, columnsToCopy(k)(1)).value
                Next k
            End If
        Next j

        ' ��v�����l���������݂���ꍇ�A�����𑾎��ɂ���
        If matchCount > 1 Then
            Dim l As Long
            For l = 1 To columnsToCopy.Count
                logSheet.Cells(i, columnsToCopy(l)(0)).Font.Bold = True
            Next l
        End If
    Next i

    ' �]�L���s��ꂽ���Ƃ��m�F
    MsgBox "�]�L���������܂���: " & sheetNameLog & " ���� " & sheetNameSpec
End Sub


Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    'ProcessSheetPairs�̃T�u�v���V�[�W��
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
            headerPairs = Array( _
                Array("����ID", "����ID(C)"), _
                Array("�i��", "�i��(D)"), _
                Array("�������e", "�������e(E)"), _
                Array("������", "������(F)"), _
                Array("���x", "���x(G)"), _
                Array("�O����", "�O����(L)"), _
                Array("�d��", "�d��(M)"), _
                Array("�V��������", "�V��������(N)"), _
                Array("�X�̐F", "�X�̐F(O)"), _
                Array("�����敪", "�����敪(U)") _
            )
    ElseIf sheetNameLog = "LOG_FallArrest" And sheetNameSpec = "FallArr_SpecSheet" Then
        headerPairs = Array( _
            Array("�ʂ̍ő�l", "�ʂ̏Ռ��l"), _
            Array("�ʂ�D�w�b�_�[��", "�ʂ�D�w�b�_�[��"), _
            Array("�ʂ�E�w�b�_�[��", "�ʂ�E�w�b�_�[��"), _
            Array("�ʂ�F�w�b�_�[��", "�ʂ�F�w�b�_�[��"), _
            Array("�ʂ�G�w�b�_�[��", "�ʂ�G�w�b�_�[��"), _
            Array("�ʂ�L�w�b�_�[��", "�ʂ�I�w�b�_�[��"), _
            Array("�ʂ�M�w�b�_�[��", "�ʂ�J�w�b�_�[��"), _
            Array("�ʂ�N�w�b�_�[��", "�ʂ�K�w�b�_�[��"), _
            Array("�ʂ�O�w�b�_�[��", "�ʂ�L�w�b�_�[��"), _
            Array("�ʂ�U�w�b�_�[��", "�ʂ�M�w�b�_�[��") _
        )
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bic_SpecSheet" Then
        headerPairs = Array( _
            Array("�ʂ̍ő�l", "�ʂ̏Ռ��l"), _
            Array("�ʂ�D�w�b�_�[��", "�ʂ�D�w�b�_�[��"), _
            Array("�ʂ�U�w�b�_�[��", "�ʂ�M�w�b�_�[��") _
        )
    ElseIf sheetNameLog = "LOG_BaseBall" And sheetNameSpec = "Base_SpecSheet" Then
        headerPairs = Array( _
            Array("�ʂ̍ő�l", "�ʂ̏Ռ��l"), _
            Array("�ʂ�D�w�b�_�[��", "�ʂ�D�w�b�_�[��"), _
            Array("�ʂ�U�w�b�_�[��", "�ʂ�M�w�b�_�[��") _
        )
    Else
        headerPairs = Array()
    End If

    GetHeaderPairs = headerPairs
End Function

' �e��ɏ����ݒ������
Sub CustomizeSheetFormats()

    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            ' Determine the data type based on the column header and set the format accordingly
            Select Case True
                Case InStr(cell.value, "������") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "yyyy-mm-dd"
                Case InStr(cell.value, "���x") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "�ő�l(kN)") > 0, InStr(cell.value, "�d��") > 0, _
                     InStr(cell.value, "�V��������") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "�ő�l���L�^��������") > 0, _
                     InStr(cell.value, "4.9kN�̌p������") > 0, _
                     InStr(cell.value, "7.3kN�̌p������") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "0.00"
                Case InStr(cell.value, "ID") > 0, InStr(cell.value, "����ID") > 0, _
                     InStr(cell.value, "�i��") > 0, InStr(cell.value, "�����ʒu") > 0, _
                     InStr(cell.value, "�O����") > 0, InStr(cell.value, "�X�̐F") > 0, _
                     InStr(cell.value, "���i���b�g") > 0, InStr(cell.value, "�X�̃��b�g") > 0, _
                     InStr(cell.value, "�������b�g") > 0, InStr(cell.value, "�\������") > 0, _
                     InStr(cell.value, "�ђʌ���") > 0, InStr(cell.value, "�����敪") > 0
                    Set rng = ws.Range(cell, ws.Cells(ws.Rows.Count, cell.Column).End(xlUp))
                    rng.NumberFormat = "@"
            End Select
        Next cell
    Next sheet


End Sub

Sub CustomizeSheetFormats_Old()

    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim col As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "�ő�l(kN)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.00 "
            ElseIf InStr(1, cell.value, "�ő�l(G)") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0 "
            ElseIf InStr(1, cell.value, "����") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "���x") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "�d��") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            ElseIf InStr(1, cell.value, "���b�g") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "@"
            ElseIf InStr(1, cell.value, "�V��������") > 0 Then
                Set rng = ws.Range(cell, ws.Cells(Rows.Count, cell.Column).End(xlUp))
                rng.NumberFormat = "0.0 "
            End If
        Next cell
    Next sheet
End Sub
' B���ID���쐬����B
Sub TransformIDs()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim newID As String
    
    ' LOG_Helmet�V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' 2�s�ڂ���ŏI�s�܂Ń��[�v�i1�s�ڂ̓w�b�_�[�Ɖ���j
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "C").value
        
        ' ID��ϊ�
        newID = GenerateNewID(cellValue)
        
        ' �V����ID���Z���ɃZ�b�g
        ws.Cells(i, "B").value = newID
    Next i
End Sub

Function GenerateNewID(cellValue As String) As String
    'TransformIDs�̃T�u�v���V�[�W��
    Dim numPart As String
    Dim otherPart As String
    Dim newID As String
    Dim matches As Object
    Dim reNum As Object
    Dim reOther As Object
    Dim startIndex As Long
    
    ' ���l�����̐��K�\���I�u�W�F�N�g���쐬
    Set reNum = CreateObject("VBScript.RegExp")
    reNum.Global = False
    reNum.IgnoreCase = False
    reNum.Pattern = "�d{3,5}F?"
    
    ' ���l�����𒊏o
    If reNum.Test(cellValue) Then
        Set matches = reNum.Execute(cellValue)
        numPart = ExtractNumberPart(matches(0).value)
        newID = numPart
        
        ' ����̕�����ɑ��������𒊏o
        otherPart = ExtractOtherPart(cellValue, reNum.Execute(cellValue)(0).FirstIndex + 1)
        
        ' �f�o�b�O�p�̏o��
        Debug.Print numPart
        Debug.Print otherPart
        
        ' �V����ID������
        GenerateNewID = newID & otherPart
    Else
        ' ���l������������Ȃ��ꍇ�͌��̒l��Ԃ�
        GenerateNewID = cellValue
    End If
End Function

Function ExtractNumberPart(numPart As String) As String
        'TransformIDs�̃T�u�v���V�[�W��
    Dim hasF As Boolean
    ' ���������̖�����F�̏ꍇ
    hasF = Right(numPart, 1) = "F"
    If hasF Then
        ' ������F���������Đ��l�������擾
        numPart = Left(numPart, Len(numPart) - 1)
        ' �V����ID�𐶐��i�O���F��ǉ��j
        ExtractNumberPart = "F" & numPart & "F"
    Else
        ' ������F���Ȃ��ꍇ�͂��̂܂܎g�p
        ExtractNumberPart = numPart
    End If
End Function

Function ExtractOtherPart(cellValue As String, startIndex As Long) As String
    'TransformIDs�̃T�u�v���V�[�W��
    Dim reOther As Object
    Dim matches As Object
    Dim otherPart As String
    Dim endIndex As Long
    
    ' ����̕�����ɑ��������𒊏o���邽�߂̐��K�\��
    Set reOther = CreateObject("VBScript.RegExp")
    reOther.Global = False
    reOther.IgnoreCase = False
    reOther.Pattern = "-(�V|�O|��|��)"
    
    If reOther.Test(cellValue) Then
        startIndex = reOther.Execute(cellValue)(0).FirstIndex + 1
        otherPart = Mid(cellValue, startIndex)
        
        ' �Ō��'-'�ȍ~�̕�������菜��
        endIndex = InStrRev(otherPart, "-")
        If endIndex > 0 Then
            otherPart = Left(otherPart, endIndex - 1)
        End If
        ExtractOtherPart = otherPart
    Else
        ExtractOtherPart = ""
    End If
End Function







