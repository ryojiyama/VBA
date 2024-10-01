Attribute VB_Name = "SpecSheet"
 ' ���i�ԁA�����ӏ��Ȃǂɉ�����ID���쐬����
Sub createID()

    Dim ws As Worksheet
    Dim i As Long
    Dim id As String
    Dim lastRow As Long

    ' "Bicycle_SpecSheet" ���܂ރV�[�g����������
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Bicycle_SpecSheet", vbTextCompare) > 0 Then
            
            ' �Ō�̍s���擾
            lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
            
            ' �e�s�ɑ΂���ID�𐶐�
            For i = 2 To lastRow ' 1�s�ڂ̓w�b�_�Ɖ���
                id = GenerateID(ws, i)
                ' B���ID���Z�b�g
                ws.Cells(i, 2).value = id
            Next i
            
        End If
    Next ws

End Sub



Function GenerateID(ws As Worksheet, rowIndex As Long) As String
    ' CreateID()�̃T�u�v���V�[�W��
    Dim id As String

    ' C��: 2���ȉ��̐���
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C���D��̊Ԃ�"-"
    
    ' D��̏�����ύX
    id = id & ExtractNumber(ws.Cells(rowIndex, 4).value)
    id = id & "-" ' D���E��̊Ԃ�"-"
    
    ' N��i14��ځj�̏���
    id = id & GetColumnNValue(ws.Cells(rowIndex, 14).value)
    id = id & "-" ' E���M��̊Ԃ�"-"
    
    ' M��i13��ځj�̏���
    id = id & GetColumnMValue(ws.Cells(rowIndex, 13).value)
    id = id & "-" ' M���O��̊Ԃ�"-"
    
    ' O��i15��ځj�̏���
    id = id & GetColumnOValue(ws.Cells(rowIndex, 15).value)
    id = id & "-" ' O���P��̊Ԃ�"-"
    
    ' P��i16��ځj�̏���
    id = id & GetColumnPValue(ws.Cells(rowIndex, 16).value)
    
    ' ��������ID��Ԃ�
    GenerateID = id
End Function

Function ExtractNumber(value As String) As String
    ' value �𕶎���ɕϊ����ĕԂ�
    ExtractNumber = CStr(value)
End Function

Function GetColumnCValue(value As Variant) As String
    ' GenerateID�̃T�u�֐�
    If Len(value) <= 2 Then
        GetColumnCValue = Right("00" & value, 2)
    Else
        GetColumnCValue = "??"
    End If
End Function

Function GetColumnNValue(value As Variant) As String
    ' Value�� "�O����" ���܂܂�Ă���ꍇ�� "�O" ��Ԃ�
    If InStr(value, "�O����") > 0 Then
        GetColumnNValue = "�O"
    ' Value�� "�㓪��" ���܂܂�Ă���ꍇ�� "��" ��Ԃ�
    ElseIf InStr(value, "�㓪��") > 0 Then
        GetColumnNValue = "��"
    ' Value�� "��������" ���܂܂�Ă���ꍇ�� "��" ��Ԃ�
    ElseIf InStr(value, "��������") > 0 Then
        GetColumnNValue = "��"
    ' Value�� "�E������" ���܂܂�Ă���ꍇ�� "�E" ��Ԃ�
    ElseIf InStr(value, "�E������") > 0 Then
        GetColumnNValue = "�E"
    ' ����ȊO�̏ꍇ�� "?" ��Ԃ�
    Else
        GetColumnNValue = "??"
    End If
End Function


Function GetColumnMValue(value As Variant) As String
    ' GenerateID�̃T�u�֐�
    Select Case value
        Case "����"
            GetColumnMValue = "Hot"
        Case "�ቷ"
            GetColumnMValue = "Cold"
        Case "�Z����"
            GetColumnMValue = "Wet"
        Case Else
            GetColumnMValue = "?"
    End Select
End Function


Function GetColumnOValue(value As Variant) As String
    ' Value�� "��" �̏ꍇ�� "��" ��Ԃ�
    If value = "��" Then
        GetColumnOValue = "��"
    ' Value�� "��" �̏ꍇ�� "��" ��Ԃ�
    ElseIf value = "��" Then
        GetColumnOValue = "��"
    ' ����ȊO�̏ꍇ�� "���̑�" ��Ԃ�
    Else
        GetColumnOValue = "���̑�"
    End If
End Function


Function GetColumnPValue(value As Variant) As String
    ' Value�� "A", "E", "J", "M", "O" �̏ꍇ�͂��̂܂ܕԂ�
    If value = "A" Then
        GetColumnPValue = "A"
    ElseIf value = "E" Then
        GetColumnPValue = "E"
    ElseIf value = "J" Then
        GetColumnPValue = "J"
    ElseIf value = "M" Then
        GetColumnPValue = "M"
    ElseIf value = "O" Then
        GetColumnPValue = "O"
    ' ����ȊO�̏ꍇ�� "���̑�" ��Ԃ�
    Else
        GetColumnPValue = "���̑�"
    End If
End Function




' ��SpecSheet�ɓ]�L����v���V�[�W���̖{�́B
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
    Call createID              ' B���ID���쐬����B
    Call ProcessSheetPairs          ' �]�L����������v���V�[�W��

End Sub
Function HighlightDuplicateValues() As Boolean
    ' SyncSpecSheetToLogHel�̃T�u�v���V�[�W��
    Dim sheetName As String
    sheetName = "Bicycle_SpecSheet"

    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundDuplicate As Boolean
    foundDuplicate = False ' ���l�������������ǂ����̃t���O��������

    ' �V�[�g�I�u�W�F�N�g��ݒ�
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row

    ' �ȑO�̐F���N���A
    For i = 2 To lastRow
        ws.Cells(i, "J").Interior.colorIndex = xlNone
        ws.Cells(i, "K").Interior.colorIndex = xlNone
    Next i

    ' �F�̃C���f�b�N�X��������
    Dim colorIndex As Integer
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�

    ' J+K���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        For j = i + 1 To lastRow
            ' J���K��̒l��g�ݍ��킹�Ĕ�r
            If ws.Cells(i, "J").value & ws.Cells(i, "K").value = ws.Cells(j, "J").value & ws.Cells(j, "K").value And ws.Cells(i, "J").value <> "" And ws.Cells(i, "K").value <> "" Then
                ' ���l�����Z�������������ꍇ�A�t���O��True�ɐݒ肵�A�Z���ɐF��h��
                foundDuplicate = True
                ws.Cells(i, "J").Interior.colorIndex = colorIndex
                ws.Cells(j, "J").Interior.colorIndex = colorIndex
                ws.Cells(i, "K").Interior.colorIndex = colorIndex
                ws.Cells(j, "K").Interior.colorIndex = colorIndex
            End If
        Next j
        ' ���l�����������ꍇ�A���̐F�ɕύX
        If foundDuplicate And ws.Cells(i, "J").Interior.colorIndex <> xlNone Then
            colorIndex = colorIndex + 1
            ' �F�C���f�b�N�X�̍ő�l�𒴂��Ȃ��悤�Ƀ`�F�b�N
            If colorIndex > 56 Then colorIndex = 3 ' �F�C���f�b�N�X�����Z�b�g
        End If
    Next i

    ' ���l�����������Ȃ������ꍇ�AJ���K��̃Z���̐F���N���A
    If Not foundDuplicate Then
        For i = 2 To lastRow
            ws.Cells(i, "J").Interior.Color = xlNone
            ws.Cells(i, "K").Interior.Color = xlNone
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
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row

    ' �ŏI���"M"(�����敪)�ɌŒ�
    Dim lastCol As Long
    lastCol = ws.Columns("M").column

    ' �w��͈͂����[�v
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' �󔒂̃`�F�b�N
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "�󔒃Z��: " & cell.Address(False, False) & vbNewLine
            End If

            ' ��G�AH�AJ�AK�Ő��l�̊m�F
            If j = Columns("G").column Or j = Columns("H").column Or j = Columns("J").column Or j = Columns("K").column Then
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
            End If

            ' ��N�AO�AP�ŕ�����̊m�F
            If j = Columns("N").column Or j = Columns("O").column Or j = Columns("P").column Then
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
        Array("LOG_Bicycle", "Bicycle_SpecSheet"), _
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
    Dim structureCol As Long
    Dim penetrationCol As Long
    Dim zairyoCol As Long
    Dim logSum As Double
    Dim specSum As Double

    ' ���[�N�V�[�g���Z�b�g
    Set logSheet = ThisWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ThisWorkbook.Worksheets(sheetNameSpec)

    ' LOG�V�[�g�̍ŏI�s���擾
    lastRowLog = logSheet.Cells(logSheet.Rows.count, "J").End(xlUp).row
    ' Spec�V�[�g�̍ŏI�s���擾
    lastRowSpec = helSpec.Cells(helSpec.Rows.count, "J").End(xlUp).row

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
                logCol = col.column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.value = pair(1) Then
                helSpecCol = col.column
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
        logSum = logSheet.Cells(i, "J").value + logSheet.Cells(i, "K").value

        For j = 2 To lastRowSpec
            specSum = helSpec.Cells(j, "J").value + helSpec.Cells(j, "K").value

            ' J+K�̍��v����v����ꍇ�ɓ]�L�������s��
            If logSum = specSum Then
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.count
                    logSheet.Cells(i, columnsToCopy(k)(0)).value = helSpec.Cells(j, columnsToCopy(k)(1)).value
                Next k
            End If
        Next j

        ' ��v�����l���������݂���ꍇ�A�����𑾎��ɂ���
        If matchCount > 1 Then
            Dim l As Long
            For l = 1 To columnsToCopy.count
                logSheet.Cells(i, columnsToCopy(l)(0)).Font.Bold = True
            Next l
        End If
    Next i

    ' �ǉ��@�\: �u�\��_�������ʁv�Ɓu�ϊђ�_�������ʁv�̗�Ɂu���i�v�����
    structureCol = FindHeaderColumn(logHeader, "�O�ό���")
    penetrationCol = FindHeaderColumn(logHeader, "�����Ђ�����")
    zairyoCol = FindHeaderColumn(logHeader, "�ޗ��E�t���i����")

    If structureCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, structureCol).value = "���i"
        Next i
    Else
        MsgBox "�w�b�_�[�u�\��_�������ʁv��������܂���ł����B"
    End If

    If penetrationCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, penetrationCol).value = "���i"
        Next i
    Else
        MsgBox "�w�b�_�[�u�ϊђ�_�������ʁv��������܂���ł����B"
    End If
    
    If zairyoCol > 0 Then
        For i = 2 To lastRowLog
            logSheet.Cells(i, zairyoCol).value = "���i"
        Next i
    Else
        MsgBox "�w�b�_�[�u�ޗ��E�t���i�����v��������܂���ł����B"
    End If

    ' �]�L���s��ꂽ���Ƃ��m�F
    MsgBox "�]�L���������܂���: " & sheetNameLog & " ���� " & sheetNameSpec
End Sub


' �w�肵���w�b�_�[��������̔ԍ����擾����֐�
Function FindHeaderColumn(headerRow As Range, headerName As String) As Long
    Dim col As Range
    For Each col In headerRow.Cells
        If col.value = headerName Then
            FindHeaderColumn = col.column
            Exit Function
        End If
    Next col
    FindHeaderColumn = -1 ' �w�b�_�[��������Ȃ������ꍇ
End Function

Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    'ProcessSheetPairs�̃T�u�v���V�[�W��
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
            headerPairs = Array( _
                Array("����ID", "����ID(C)"), _
                Array("�i��", "�i��(D)"), _
                Array("�������e", "�����ʒu(E)"), _
                Array("������", "������(F)"), _
                Array("���x", "���x(G)"), _
                Array("�ő�l(kN)", "�Ռ��l(H)"), _
                Array("�O����", "�O����(L)"), _
                Array("�d��", "�d��(M)"), _
                Array("�V��������", "�V��������(N)"), _
                Array("�X�̐F", "�X�̐F(O)"), _
                Array("���b�gNo.", "�������b�g(P)"), _
                Array("�X�̃��b�g", "�X�̃��b�g(Q)"), _
                Array("�������b�g", "�������b�g(R)"), _
                Array("�\��_��������", "�\��/����(S)"), _
                Array("�ϊђ�_��������", "�ϊђ�/����(U)"), _
                Array("�����敪", "�������e(U)") _
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
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bicycle_SpecSheet" Then
        headerPairs = Array( _
            Array("ID", "����ID(C)"), _
            Array("����ID", "����ID"), _
            Array("�i��", "�i��"), _
            Array("���b�g�ԍ�", "���b�g�ԍ�"), _
            Array("������", "������"), _
            Array("���x", "���x"), _
            Array("���x", "���x"), _
            Array("�d��", "�d��"), _
            Array("�O����", "�O����"), _
            Array("�����ӏ�", "�����ӏ�"), _
            Array("�A���r��", "�A���r��"), _
            Array("�l���͌^", "�l���͌^") _
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










