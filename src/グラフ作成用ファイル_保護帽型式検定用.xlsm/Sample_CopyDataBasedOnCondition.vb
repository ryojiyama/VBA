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

        ' �V�[�g�y�A�����݂��邩�`�F�b�N
        If SheetExists(logSheetName) And SheetExists(specSheetName) Then
            ' �V�[�g�y�A�����������ꍇ�ɏ��������s
            Call CopyDataBasedOnCondition(logSheetName, specSheetName)
        End If
    Next pair
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Sub CopyDataBasedOnCondition(sheetNameLog As String, sheetNameSpec As String)
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
    lastRowLog = logSheet.Cells(logSheet.Rows.Count, "H").End(xlUp).Row
    ' Spec�V�[�g�̍ŏI�s���擾
    lastRowSpec = helSpec.Cells(helSpec.Rows.Count, "H").End(xlUp).Row

    ' �w�b�_�[�s���擾
    Set logHeader = logSheet.Rows(1)
    Set helSpecHeader = helSpec.Rows(1)

    ' �]�L�����̃y�A���R���N�V�����ɒ�`
    Set columnsToCopy = New Collection

    ' �y�A�ƂȂ�w�b�_�[�����擾
    colPair = GetHeaderPairs(sheetNameLog, sheetNameSpec)

    ' �e�w�b�_�[�s�𑖍����Ĉ�v����w�b�_�[��������
    Dim pair As Variant
    For Each pair In colPair
        Dim logCol As Long
        Dim helSpecCol As Long
        logCol = 0
        helSpecCol = 0
        For Each col In logHeader.Cells
            If col.Value = pair(0) Then
                logCol = col.Column
                Exit For
            End If
        Next col
        For Each col In helSpecHeader.Cells
            If col.Value = pair(1) Then
                helSpecCol = col.Column
                Exit For
            End If
        Next col
        If logCol > 0 And helSpecCol > 0 Then
            columnsToCopy.Add Array(logCol, helSpecCol)
        End If
    Next pair

    ' �l���r���ē]�L
    For i = 2 To lastRowLog
        matchCount = 0
        For j = 2 To lastRowSpec
            If logSheet.Cells(i, "H").Value = helSpec.Cells(j, "H").Value Then
                ' H��̒l����v�����ꍇ�A�e��̓��e��]�L
                matchCount = matchCount + 1
                Dim k As Long
                For k = 1 To columnsToCopy.Count
                    logSheet.Cells(i, columnsToCopy(k)(0)).Value = helSpec.Cells(j, columnsToCopy(k)(1)).Value
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
End Sub

Function GetHeaderPairs(sheetNameLog As String, sheetNameSpec As String) As Variant
    Dim headerPairs As Variant

    If sheetNameLog = "LOG_Helmet" And sheetNameSpec = "Hel_SpecSheet" Then
        headerPairs = Array( _
            Array("����ID(C)", "�Ռ��l(H)"), _
            Array("�i��(D)", "�i��"), _
            Array("�������e(E)", "�������e"), _
            Array("������(F)", "������"), _
            Array("���x(G)", "���x"), _
            Array("�O����(L)", "�O����"), _
            Array("�d��(M)", "�d��"), _
            Array("�V��������(N)", "�V��������"), _
            Array("�X�̐F(O)", "�X�̐F") _
            Array("�����敪(U)", "�����敪") _
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
    ' �K�v�ɉ����đ��̃V�[�g�y�A��ǉ�
    ElseIf sheetNameLog = "LOG_Bicycle" And sheetNameSpec = "Bic_SpecSheet" Then
        headerPairs = Array( _
            Array("�l1", "�l2"), _
            Array("�w�b�_�[1", "�w�b�_�[2") _
            ' ���̃y�A��ǉ�
        )
    ElseIf sheetNameLog = "LOG_BaseBall" And sheetNameSpec = "Base_SpecSheet" Then
        headerPairs = Array( _
            Array("�lA", "�lB"), _
            Array("�w�b�_�[A", "�w�b�_�[B") _
            ' ���̃y�A��ǉ�
        )
    Else
        headerPairs = Array()
    End If

    GetHeaderPairs = headerPairs
End Function
