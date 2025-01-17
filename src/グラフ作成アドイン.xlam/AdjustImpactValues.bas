Attribute VB_Name = "AdjustImpactValues"
'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�FLOG_�V�[�g�̏Ռ��l�𐻕i��ʂɉ����Ē������A�t�H�[�}�b�g��ݒ�
' �����F�Ȃ�
' �⑫�F�e���i�̏Ռ��l�ɍs�ԍ��ɉ����������l�����Z���ďd����h��
'*******************************************************************************
Sub AdjustImpactValuesWithCustomFormatForAllLOGSheets()
    Dim ws As Worksheet
    Dim impactCol As Long
    Dim impactValue As Double
    Dim rowNum As Long
    Dim lastRow As Long
    Dim i As Long
    Dim backupCol As Long
    Dim adjustmentFactor As Double
    Dim displayFormat As String

    ' �o�b�N�A�b�v��ۑ������iX��24��ځj
    backupCol = 24

    ' ���ׂẴV�[�g�����ɒT��
    For Each ws In ActiveWorkbook.Sheets
        ' �V�[�g����"LOG_"���܂ރV�[�g�ɑ΂��Ă̂ݏ������s��
        If InStr(ws.Name, "LOG_") > 0 Then

            ' �w�b�_�[�s����"�ő�l("���܂ޗ��������
            For i = 1 To ws.Cells(1, ws.columns.Count).End(xlToLeft).column
                If InStr(ws.Cells(1, i).value, "�ő�l(") > 0 Then
                    impactCol = i
                    Exit For
                End If
            Next i

            ' "�ő�l"�񂪌�����Ȃ������ꍇ�͎��̃V�[�g��
            If impactCol = 0 Then
                MsgBox ws.Name & " �V�[�g�ɍő�l�񂪌�����܂���B"
                GoTo NextSheet
            End If

            ' �ŏI�s���擾
            lastRow = ws.Cells(ws.Rows.Count, impactCol).End(xlUp).row

            ' �V�[�g���ɉ����Ē����W����ݒ�
            Select Case ws.Name
                Case "LOG_Helmet", "LOG_FallArrest"
                    adjustmentFactor = 0.000001
                    displayFormat = "0.000000" ' �����_�ȉ�6���܂ŕ\��
                Case "LOG_BaseBall", "LOG_Bicycle"
                    adjustmentFactor = 0.01
                    displayFormat = "0.00" ' �����_�ȉ�2���܂ŕ\��
                Case Else
                    MsgBox ws.Name & " �V�[�g�����K�؂ł͂���܂���B�������X�L�b�v���܂��B"
                    GoTo NextSheet
            End Select

            ' 2�s�ڈȍ~�̃Z���ɑ΂��ď������s��
            For rowNum = 2 To lastRow
                impactValue = ws.Cells(rowNum, impactCol).value

                ' ���� impactValue �� X��Ƀo�b�N�A�b�v
                ws.Cells(rowNum, backupCol).value = impactValue

                ' �v�Z����K�p
                impactValue = impactValue + (rowNum * adjustmentFactor)

                ' �v�Z���ʂ����̗�ɑ��
                ws.Cells(rowNum, impactCol).value = impactValue

                ' �Z���̕\���`����ݒ�
                ws.Cells(rowNum, impactCol).NumberFormat = displayFormat
            Next rowNum

NextSheet:
            impactCol = 0 ' ���̃V�[�g�̂��߂Ƀ��Z�b�g

        End If
    Next ws
    Call HighlightDuplicateValues
End Sub
'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�Ռ��l�̏d�����`�F�b�N���A�d���l�ɐF�t�����s��
' �����F�Ȃ�
' �⑫�F���i��ʂ��ƂɑΏۗ��ς��ďd���`�F�b�N�����{
'*******************************************************************************
Sub HighlightDuplicateValues()
    ' �ΏۃV�[�g���̃��X�g
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    
    ' �ϐ��錾
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim valueToFind As Variant
    Dim colorIndex As Integer
    Dim sheetName As Variant
    
    ' �V�[�g���Ƃɏ���
    For Each sheetName In sheetNames
        ' �V�[�g�I�u�W�F�N�g��ݒ�i�G���[�n���h�����O��ǉ��j
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' �V�[�g�����݂���ꍇ�̂ݏ��������s
        If Not ws Is Nothing Then
            ' �V�[�g�ɉ����đΏۗ��ݒ�
            Dim targetColumn As String
            Select Case CStr(sheetName)
                Case "LOG_Helmet"
                    targetColumn = "H"  ' �w�����b�g�̃��O�� H ��
                Case "LOG_FallArrest"
                    targetColumn = "H"  ' �ė����~�p���̃��O�� I ��
                Case "LOG_Bicycle"
                    targetColumn = "J"  ' ���]�Ԃ̃��O�� J ��
                Case "LOG_BaseBall"
                    targetColumn = "H"  ' �싅�p��̃��O�� K ��
            End Select
            
            ' �ŏI�s���擾
            lastRow = ws.Cells(ws.Rows.Count, targetColumn).End(xlUp).row
            
            ' �Ώ۔͈͂̐F���N���A
            ws.Range(targetColumn & "2:" & targetColumn & lastRow).Interior.colorIndex = xlNone
            
            ' �F�̃C���f�b�N�X��������
            colorIndex = 3 ' Excel�̐F�C���f�b�N�X��3����n�܂�
            
            ' �V�[�g���Ƃ̏d���`�F�b�N
            For i = 2 To lastRow
                ' ���݂̃Z���̒l���擾
                valueToFind = ws.Cells(i, targetColumn).value
                
                ' �l����łȂ����Ƃ��m�F
                If Not IsEmpty(valueToFind) Then
                    ' �����l�����Z�������ɐF�t������Ă��Ȃ����`�F�b�N
                    If ws.Cells(i, targetColumn).Interior.colorIndex = xlNone Then
                        Dim duplicateFound As Boolean
                        duplicateFound = False
                        
                        For j = i + 1 To lastRow
                            ' �V�[�g���Ƃ̏d���`�F�b�N���W�b�N
                            Select Case CStr(sheetName)
                                Case "LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall"
                                    ' ���ׂĊ��S��v�Ń`�F�b�N
                                    If ws.Cells(j, targetColumn).value = valueToFind Then
                                        duplicateFound = True
                                    End If
                            End Select
                            
                            ' �d�������������ꍇ�A�F��t����
                            If duplicateFound Then
                                ws.Cells(i, targetColumn).Interior.colorIndex = colorIndex
                                ws.Cells(j, targetColumn).Interior.colorIndex = colorIndex
                                duplicateFound = False  ' ���̃`�F�b�N�̂��߂Ƀ��Z�b�g
                            End If
                        Next j
                        
                        ' �F�C���f�b�N�X���X�V
                        colorIndex = colorIndex + 1
                        If colorIndex > 56 Then colorIndex = 3
                    End If
                End If
            Next i
            
            ' �I�u�W�F�N�g�̃N���A
            Set ws = Nothing
        Else
            ' �V�[�g��������Ȃ��ꍇ�̃f�o�b�O�o��
            Debug.Print "�V�[�g��������܂���ł���: " & sheetName
        End If
    Next sheetName
End Sub


