Attribute VB_Name = "SpecSheet"
 ' ���i�ԁA�����ӏ��Ȃǂɉ�����ID���쐬����
Sub CreateID(sheetName As String)
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim id As String
    
    ' �����œn���ꂽ�V�[�g�����g�p
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    
    ' �Ō�̍s���擾
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' �e�s�ɑ΂���ID�𐶐�
    For i = 2 To lastRow ' 1�s�ڂ̓w�b�_�Ɖ���
        id = GenerateID(ws, i)
        ' B���ID���Z�b�g
        ws.Cells(i, 2).value = id
    Next i
End Sub
Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' CreateID()�̃T�u�v���V�[�W��
    Dim id As String

    ' C��: 2���ȉ��̐���
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C���D��̊Ԃ�"-"
    ' D��̏�����ύX
    id = id & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    id = id & "-" ' Fm��E��̊Ԃ�"-"
    id = id & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E��̏���
    id = id & "-" ' Fm��E��̊Ԃ�"-
    id = id & GetColumnIValue(ws.Cells(rowIndex, 9).value) ' I��̏���
    id = id & "-" ' I���L��̊Ԃ�"-"
    id = id & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L��̏���

    GenerateID = id
End Function
Function ExtractNumberWithF(value As String) As String
' GenerateID�̃T�u�֐�
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' ���K�\���I�u�W�F�N�g�̍쐬
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{3,6}"
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
    ElseIf InStr(value, "����") > 0 Then
        Dim Parts() As String
        Parts = Split(value, "_")
        
        If UBound(Parts) >= 1 Then
            Dim angle As String
            Dim direction As String
            
            ' �p�x�𒊏o
            angle = Replace(Parts(0), "����", "")
            
            ' �����𒊏o�Ɛ��`
            direction = Parts(1)
            direction = Replace(direction, "�O", "�O")
            direction = Replace(direction, "��", "��")
            direction = Replace(direction, "��", "��")
            direction = Replace(direction, "�E", "�E")
            
            GetColumnEValue = "��" & angle & direction
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
        Case "�퉷"
            GetColumnIValue = "Nrml"
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
' �Ռ��l�݂̂�LOG�V�[�g�ɓ]�L����B
Sub TransferValuesBetweenSheets()

    ' �V�[�g�y�A�̔z����쐬
    Dim sheetPairs As Variant
    sheetPairs = Array( _
        Array("Hel_SpecSheet", "LOG_Helmet"), _
        Array("Bicycle_SpecSheet", "LOG_Bicycle"), _
        Array("Fall_SpecSheet", "LOG_FallArrest"), _
        Array("BaseBall_SpecSheet", "LOG_BaseBall"))
    
    Dim i As Long
    Dim specSheet As Worksheet
    Dim logSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim cell As Range
    
    ' �V�[�g�y�A�����[�v���ď���
    For i = LBound(sheetPairs) To UBound(sheetPairs)
        ' SpecSheet �� LOG_ �V�[�g��ݒ�
        On Error Resume Next
        Set specSheet = ActiveWorkbook.Sheets(sheetPairs(i)(0))
        Set logSheet = ActiveWorkbook.Sheets(sheetPairs(i)(1))
        On Error GoTo 0
        
        ' �V�[�g�����݂���ꍇ�ɏ��������s
        If Not specSheet Is Nothing And Not logSheet Is Nothing Then
            ' SpecSheet��H��̍ŏI�s���擾
            lastRow = specSheet.Cells(specSheet.Rows.Count, "H").End(xlUp).row
            
            ' H��̃f�[�^��]�L����͈͂�ݒ�
            Set dataRange = specSheet.Range("H2:H" & lastRow) ' H2����ŏI�s�܂�
                       
            ' SpecSheet����LOG_�V�[�g�֒l��]�L
            logSheet.Range("H2").Resize(dataRange.Rows.Count).value = dataRange.value
        Else
            ' �V�[�g�����݂��Ȃ��ꍇ�̃f�o�b�O�o��
            Debug.Print "�V�[�g��������܂���ł���: " & sheetPairs(i)(0) & " �܂��� " & sheetPairs(i)(1)
        End If
        
        ' �I�u�W�F�N�g�̃N���A
        Set specSheet = Nothing
        Set logSheet = Nothing
    Next i

End Sub





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

    Call UpdateCrownClearance ' �V�������܂𒲐�
    Call ProcessSheetPairs   ' �]�L����������v���V�[�W��

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
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row

    ' �F�̃C���f�b�N�X��������
    Dim colorIndex As Integer
    colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�

    ' H���2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' M��̒l���`�F�b�N���A"�˗�"���܂܂��ꍇ�̓t���O��False�ɐݒ�
        If InStr(ws.Cells(i, "M").value, "�˗�") > 0 Then
            foundDuplicate = False
        Else
            For j = i + 1 To lastRow
                If ws.Cells(i, "H").value = ws.Cells(j, "H").value And ws.Cells(i, "H").value <> "" Then
                    ' ���l�����Z�������������ꍇ�A�t���O��True�ɐݒ肵�A�Z���ɐF��h��
                    foundDuplicate = True
                    ws.Cells(i, "H").Interior.colorIndex = colorIndex
                    ws.Cells(j, "H").Interior.colorIndex = colorIndex
                End If
            Next j
            ' ���l�����������ꍇ�A���̐F�ɕύX
            If foundDuplicate And ws.Cells(i, "H").Interior.colorIndex <> xlNone Then
                colorIndex = colorIndex + 1
                ' �F�C���f�b�N�X�̍ő�l�𒴂��Ȃ��悤�Ƀ`�F�b�N
                If colorIndex > 56 Then colorIndex = 3 ' �F�C���f�b�N�X�����Z�b�g
            End If
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
    Set ws = ActiveWorkbook.Sheets(sheetName)

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' �ŏI���"M"(�����敪)�ɌŒ�
    Dim lastCol As Long
    lastCol = ws.columns("M").column

    ' �w��͈͂����[�v
    For i = 2 To lastRow
        For j = 2 To lastCol
            Set cell = ws.Cells(i, j)

            ' �󔒂̃`�F�b�N
            If IsEmpty(cell.value) Then
                errorMsg = errorMsg & "�󔒃Z��: " & cell.Address(False, False) & vbNewLine
            End If

            ' ��G(���x)�AH(�Ռ��l)�AJ(�d��)�AK(�V��������)�Ő��l�̊m�F
            If j = columns("G").column Or j = columns("H").column Or j = columns("J").column Or j = columns("K").column Then
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

            ' ��N(�������b�g)�AO(�X�̃��b�g)�AP(�������b�g)�ŕ�����̊m�F
            If j = columns("N").column Or j = columns("O").column Or j = columns("P").column Then
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
' �V�������Ԃ�"Setting"�V�[�g�̃f�[�^�ɍ��킹�Ē�������B
Sub UpdateCrownClearance()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim colGenshoNoSukima As Integer
    Dim colKaisu As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim i As Long
    Dim skipCount As Long

    ' �V�[�g���Z�b�g
    Set wsHelSpec = ActiveWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ActiveWorkbook.Sheets("Setting")

    ' �w�b�_�[�̗�ԍ����擾
    colHinban = GetColumnIndex(wsHelSpec, "�i��(D)")
    colBoutai = GetColumnIndex(wsSetting, "�X��No.")
    colTencho = GetColumnIndex(wsHelSpec, "�V������")
    colTenchoSukima = GetColumnIndex(wsHelSpec, "�V��������(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "���肷����")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "�V������")
    colGenshoNoSukima = GetColumnIndex(wsHelSpec, "�����̂�����")
    colKaisu = GetColumnIndex(wsHelSpec, "��")

    ' �K�v�ȗ񂪌������������m�F
    If colHinban = 0 Or colBoutai = 0 Or colTencho = 0 Or _
       colTenchoSukima = 0 Or colSokuteiSukima = 0 Or _
       colTenchoNikui = 0 Or colGenshoNoSukima = 0 Or _
       colKaisu = 0 Then
        MsgBox "�K�v�ȗ񂪌�����܂���B�w�b�_�[���m�F���Ă��������B", vbCritical
        Exit Sub
    End If

    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row

    ' "�i��(D)" ��̒l��T�����A�]�L
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell

    skipCount = 0

    ' "�V��������(N)" �̒l�� "���肷����" �ɃR�s�[���A�l���v�Z
    For i = 2 To lastRowHelSpec
        ' �񐔂��L���ς݂̏ꍇ�̓X�L�b�v���A�J�E���g�𑝂₷
        If wsHelSpec.Cells(i, colKaisu).value <> "" Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' "�����̂�����" ���󗓂̏ꍇ�̂݃R�s�[ (�񐔗�̏�ԂɊւ�炸���s)
        If wsHelSpec.Cells(i, colGenshoNoSukima).value = "" Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = wsHelSpec.Cells(i, colTenchoSukima).value
            wsHelSpec.Cells(i, colGenshoNoSukima).value = wsHelSpec.Cells(i, colTenchoSukima).value
        End If

        ' �e�Z���̒l���擾 (�����̂����Ԃ̒l���擾)
        tenchoSukimaValue = wsHelSpec.Cells(i, colGenshoNoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value

        ' "�����̂�����"�̒l����"�V������"�̒l������
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If

        ' �񐔂ɍς���
        wsHelSpec.Cells(i, colKaisu).value = "��"

        ' Q���R���"���i"�̒l����
        wsHelSpec.Cells(i, 17).value = "���i" ' Q���17�Ԗڂ̗�
        wsHelSpec.Cells(i, 18).value = "���i" ' R���18�Ԗڂ̗�

NextRow:
    Next i

    ' ���b�Z�[�W�̕\��
    If skipCount > 0 Then
        MsgBox "�C���͂��łɍs���܂����B�i" & skipCount & "�s�X�L�b�v����܂����j"
    Else
        MsgBox "�V�������Ԃ����������`�F�b�N�����肢���܂��B"
    End If

End Sub

' ��ԍ����擾����֐�
Function GetColumnIndex(targetSheet As Worksheet, headerName As String) As Integer
    Dim headerRange As Range
    Set headerRange = targetSheet.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerRange Is Nothing Then
        GetColumnIndex = headerRange.column
    Else
        GetColumnIndex = 0
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
        If sheetExists(logSheetName) And sheetExists(specSheetName) Then
            ' �V�[�g�y�A�����������ꍇ�ɏ��������s
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
            Debug.Print logSheetName
        End If
    Next pair
End Sub

Function sheetExists(sheetName As String) As Boolean
    'ProcessSheetPairs�̃T�u�v���V�[�W��
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    sheetExists = Not ws Is Nothing
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

    ' ���[�N�V�[�g���Z�b�g
    Set logSheet = ActiveWorkbook.Worksheets(sheetNameLog)
    Set helSpec = ActiveWorkbook.Worksheets(sheetNameSpec)

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
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").value = helSpec.Cells(j, "H").value Then
                ' H��̒l����v�����ꍇ�A�e��̓��e��]�L
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.Count
                    logSheet.Cells(i, columnsToCopy(k)(0)).value = helSpec.Cells(j, columnsToCopy(k)(1)).value
                Next k
                ' C��̒l��B��ɃR�s�[
                logSheet.Cells(i, "B").value = logSheet.Cells(i, "C").value
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

    ' �ǉ��@�\: �u�\��_�������ʁv�Ɓu�ϊђ�_�������ʁv�̗�Ɂu���i�v�����
    structureCol = FindHeaderColumn(logHeader, "�\��_��������")
    penetrationCol = FindHeaderColumn(logHeader, "�ϊђ�_��������")

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
                Array("�����敪", "�������e(U)"), _
                Array("�X�g���C�J����", "�X�g���C�J����(V)"), _
                Array("�������", "�������(W)"), _
                Array("�O��������", "�O��������(X)"), _
                Array("���l", "���l(Z)") _
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

Sub HighlightMismatchedRows()
    Dim sheetPairs As Variant
    Dim logSheet As Worksheet
    Dim specSheet As Worksheet
    Dim logLastRow As Long, specLastRow As Long
    Dim i As Long, j As Long
    Dim logValue1 As Variant, specValue1 As Variant
    Dim logValue2 As Variant, specValue2 As Variant
    Dim logValue3 As Variant, specValue3 As Variant
    Dim mismatchFound As Boolean ' �S�̂̕s��v���m�t���O
    
    ' �y�A���Ƃ̔�r��̒�`��ʂ̔z��Őݒ�
    Dim helmetColumns As Variant
    Dim fallArrestColumns As Variant
    Dim bicycleColumns As Variant
    Dim baseBallColumns As Variant
    
    ' �V�[�g�y�A�̒�`
    sheetPairs = Array( _
        Array("LOG_Helmet", "Hel_SpecSheet"), _
        Array("LOG_FallArrest", "FallArr_SpecSheet"), _
        Array("LOG_Bicycle", "Bicycle_SpecSheet"), _
        Array("LOG_BaseBall", "BaseBall_SpecSheet") _
    )
    
    ' �e�V�[�g�y�A�ɑΉ������̒�`
    helmetColumns = Array("J", "I", "O", "J", "I", "O")
    fallArrestColumns = Array("K", "L", "M", "K", "L", "M")
    bicycleColumns = Array("J", "I", "O", "J", "I", "O")
    baseBallColumns = Array("N", "O", "P", "N", "O", "P")
    
    mismatchFound = False ' ������
    
    ' �e�V�[�g�y�A�����[�v
    For j = LBound(sheetPairs) To UBound(sheetPairs)
        Dim logCol1 As String, logCol2 As String, logCol3 As String
        Dim specCol1 As String, specCol2 As String, specCol3 As String
        Dim columns As Variant
        
        ' �y�A�ɉ����đΉ�������I��
        Select Case sheetPairs(j)(0)
            Case "LOG_Helmet"
                columns = helmetColumns
            Case "LOG_FallArrest"
                columns = fallArrestColumns
            Case "LOG_Bicycle"
                columns = bicycleColumns
            Case "LOG_BaseBall"
                columns = baseBallColumns
            Case Else
                Debug.Print "�y�A��������܂���ł���: " & sheetPairs(j)(0)
        End Select
        
        ' ��̊��蓖��
        If Not IsEmpty(columns) Then
            logCol1 = columns(0)
            logCol2 = columns(1)
            logCol3 = columns(2)
            specCol1 = columns(3)
            specCol2 = columns(4)
            specCol3 = columns(5)
            
            ' �V�[�g�̑��݊m�F
            On Error Resume Next
            Set logSheet = ActiveWorkbook.Sheets(sheetPairs(j)(0))
            Set specSheet = ActiveWorkbook.Sheets(sheetPairs(j)(1))
            On Error GoTo 0
            
            If Not logSheet Is Nothing And Not specSheet Is Nothing Then
                logLastRow = logSheet.Cells(logSheet.Rows.Count, "C").End(xlUp).row
                specLastRow = specSheet.Cells(specSheet.Rows.Count, "C").End(xlUp).row

                ' LOG�V�[�g��2�s�ڈȍ~�����[�v
                For i = 2 To logLastRow
                    If i <= specLastRow Then
                        ' �ʂ̃v���V�[�W���Ŕ�r���s��
                        If CompareRows(logSheet, specSheet, i, logCol1, logCol2, logCol3, specCol1, specCol2, specCol3) Then
                            logSheet.Range(logSheet.Cells(i, "D"), logSheet.Cells(i, "O")).Interior.Color = RGB(255, 0, 0)
                            mismatchFound = True
                            Debug.Print "�s��v�s: " & i & " (�V�[�g: " & logSheet.Name & ")"
                        End If
                    End If
                Next i
            End If
        End If
    Next j
    
    ' �s��v�s���Ȃ��ꍇ�Ƀ��b�Z�[�W�{�b�N�X�ƃn�C���C�g���Z�b�g�����s
    If Not mismatchFound Then
        MsgBox "�s��v�s�͌�����܂���ł����B", vbInformation, "����"
        
        ' �n�C���C�g���Z�b�g: ���ׂẴV�[�g�̃n�C���C�g���N���A
        For j = LBound(sheetPairs) To UBound(sheetPairs)
            Set logSheet = ActiveWorkbook.Sheets(sheetPairs(j)(0))
            On Error Resume Next ' �V�[�g�����݂��Ȃ��ꍇ���l��
            logLastRow = logSheet.Cells(logSheet.Rows.Count, "C").End(xlUp).row
            logSheet.Range(logSheet.Cells(2, "D"), logSheet.Cells(logLastRow, "O")).Interior.colorIndex = xlNone
        Next j
    End If
End Sub

' ��r���W�b�N��ʂ̃v���V�[�W���ɕ���
Function CompareRows(logSheet As Worksheet, specSheet As Worksheet, row As Long, _
                     logCol1 As String, logCol2 As String, logCol3 As String, _
                     specCol1 As String, specCol2 As String, specCol3 As String) As Boolean
    Dim logValue1 As Variant, specValue1 As Variant
    Dim logValue2 As Variant, specValue2 As Variant
    Dim logValue3 As Variant, specValue3 As Variant
    
    ' LOG�V�[�g��SpecSheet�̒l���擾
    logValue1 = logSheet.Cells(row, logCol1).value
    logValue2 = logSheet.Cells(row, logCol2).value
    logValue3 = logSheet.Cells(row, logCol3).value
    specValue1 = specSheet.Cells(row, specCol1).value
    specValue2 = specSheet.Cells(row, specCol2).value
    specValue3 = specSheet.Cells(row, specCol3).value
    
    ' �f�o�b�O�o��
'    Debug.Print "LOG�V�[�g�s: " & row & " - " & logValue1 & ", " & logValue2 & ", " & logValue3
'    Debug.Print "SpecSheet�s: " & row & " - " & specValue1 & ", " & specValue2 & ", " & specValue3

    ' ��r����
    If logValue1 <> specValue1 Or logValue2 <> specValue2 Or logValue3 <> specValue3 Then
        CompareRows = True
    Else
        CompareRows = False
    End If
End Function










