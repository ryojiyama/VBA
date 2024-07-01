Attribute VB_Name = "Test"
Sub TransferDataToTopImpactTest()
    '�V�������݂̂̃V�[�g���쐬����B
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
    Set wsSource = ThisWorkbook.Sheets("Log_Helmet")

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' C��̒l���琻�i�R�[�h���擾
        firstDashPos = InStr(wsSource.Cells(i, 3).value, "-")
        If firstDashPos > 0 Then
            secondDashPos = InStr(firstDashPos + 1, wsSource.Cells(i, 3).value, "-")
            If secondDashPos > 0 Then
                matchName = Left(wsSource.Cells(i, 3).value, secondDashPos - 1)
            End If
        End If

        ' �e�V�[�g�����[�v���ď����Ɉ�v����V�[�g������
        For Each wsDestination In ThisWorkbook.Sheets
            If wsDestination.name = matchName Then ' �V�[�g�������i�R�[�h�Ɉ�v���邩�m�F
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


Sub TransferDataToDynamicSheets()
    ' �X�̂̎������ʂ�Ή�����V�[�g�ɓ]�L����B
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String, checkData As String
    Dim parts() As String
    Dim destinationSheetName As String

    ' �\�[�X�V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row
    
    ' Excel�̃p�t�H�[�}���X����̂��߂̐ݒ�
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSource��C������[�v���ăf�[�^������
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, 3).value
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
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

