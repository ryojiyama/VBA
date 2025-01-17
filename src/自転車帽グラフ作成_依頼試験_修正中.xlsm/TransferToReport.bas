Attribute VB_Name = "TransferToReport"
' ���|�[�g�{���̕\�Ɍ��ʂ�}������B
Sub TransferDataWithMappingAndFormatting()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim i As Long
    Dim destRow As Long
    Dim transferredRows As Long
    Const MAX_ROWS As Long = 24 ' �ő�]�L�s����12�ɐݒ�
    Dim startRow As Long
    Dim mappingDict As Object ' �}�b�s���O�p�̃f�B�N�V���i��

    ' �V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle") ' �]�L���V�[�g
    Set wsDest = ThisWorkbook.Sheets("���|�[�g�{��") ' �]�L��V�[�g

    ' �]�L���̍ŏI�s���擾
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).row

    ' �]�L��̊J�n�s��ݒ�i9�s�ڂ���J�n�j
    destRow = 9
    startRow = destRow ' �V�����ǉ������s�̊J�n�s���L�^
    transferredRows = 0 ' �]�L�����s�����J�E���g

    ' �}�b�s���O�p�̃f�B�N�V���i�����擾
    Set mappingDict = GetMappingDictionary()

    ' �]�L����2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRowSource
        ' �]�L�����s��12�s�ɒB�����璆�~
        If transferredRows >= MAX_ROWS Then
            MsgBox "�]�L�͍ő�" & MAX_ROWS & "�s�܂łɐ�������Ă��܂��B�����𒆎~���܂����B", vbExclamation
            Exit For
        End If

        ' �]�L��̍s��ǉ��idestRow�̈ʒu�ɐV�����s��}���j
        wsDest.Rows(destRow).Insert Shift:=xlDown

        ' �]�L�����s�i�f�B�N�V���i���Ɋ�Â��]�L�j
        Call TransferMappedValues(wsSource, wsDest, i, destRow, mappingDict)

        ' ���̓]�L��s�ɐi��
        destRow = destRow + 1
        transferredRows = transferredRows + 1 ' �]�L�����s�����J�E���g
    Next i

    ' �V�����ǉ����ꂽ�s�Ƀt�H�[�}�b�g��K�p
    Call ApplyFormattingToNewRows(wsDest, startRow, destRow - 1)

    ' �S�s���]�L���ꂽ�ꍇ�̓��b�Z�[�W��\��
    If transferredRows < MAX_ROWS Then
        MsgBox "�f�[�^�̓]�L���������܂����B", vbInformation
    End If

End Sub

Private Function GetMappingDictionary() As Object
    ' TransferDataWithMappingAndFormatting�̃T�u�v���V�[�W���B�]�L���Ɠ]�L��̃}�b�s���O���f�B�N�V���i���Őݒ�
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' �]�L���̗� �� �]�L��̗��񖼂Ƃ��Ė����I�ɋL�q
    dict.Add "D", "B" ' �]�L����D�� �� �]�L���B��
    dict.Add "N", "C" ' �]�L����E�� �� �]�L���C��
    dict.Add "M", "D" ' �]�L����L�� �� �]�L���D��
    dict.Add "J", "E" ' �]�L����H�� �� �]�L���E��
    dict.Add "Q", "F" ' �]�L����M�� �� �]�L���F��
    dict.Add "P", "G" ' �]�L����N�� �� �]�L���G��
    dict.Add "R", "H" ' �]�L����N�� �� �]�L���H��
    dict.Add "V", "I" ' �]�L����V�� �� �]�L���I��
    ' �K�v�ɉ����đ��̗�̃}�b�s���O���ǉ�

    Set GetMappingDictionary = dict
End Function

Private Sub TransferMappedValues(wsSource As Worksheet, wsDest As Worksheet, sourceRow As Long, destRow As Long, mappingDict As Object)
    Dim key As Variant
    
    For Each key In mappingDict.Keys
        ' �f�o�b�O�p�̃��O�o��
        Debug.Print "�]�L��: ��" & key & " �s" & sourceRow & " �l:" & wsSource.Cells(sourceRow, key).value
        Debug.Print "�]�L��: ��" & mappingDict(key) & " �s" & destRow
        
        ' �l�̓]�L�iRange.Cells �̗�Q�Ƃ𕶎���Œ��ڎw��j
        wsDest.Cells(destRow, mappingDict(key)).value = wsSource.Cells(sourceRow, key).value
    Next key
End Sub
Private Sub ApplyFormattingToNewRows(ws As Worksheet, startRow As Long, endRow As Long)
    ' TransferDataWithMappingAndFormatting�̃T�u�v���V�[�W���B�V�����ǉ����ꂽ�s�Ƀt�H�[�}�b�g��K�p���AI��Ɉ������
    Dim currentRow As Long
    Dim targetRange As Range
    Dim eRange As Range, fRange As Range, gRange As Range, iRange As Range
    
    ' 1�s������
    For currentRow = startRow To endRow
        ' ���݂̍s�͈̔͂��擾
        Set targetRange = ws.Range("B" & currentRow & ":I" & currentRow)
        
        ' �t�H�[�}�b�g��K�p
        With targetRange
            .Font.Name = "���S�V�b�N" ' �t�H���g����ݒ�
            .Font.ThemeFont = xlThemeFontMinor ' Light�E�F�C�g�ɂ���i�e�[�}�t�H���g�j
            .Font.Bold = False ' ����������
            .Font.Color = RGB(0, 0, 0) ' �t�H���g�̐F�����ɐݒ�
            
            ' �w�i�F���s���ƂɕύX
            If currentRow Mod 2 = 0 Then
                ' �����s�F�����F
                .Interior.Color = RGB(220, 230, 241)
            Else
                ' ��s�F�����D�F
                .Interior.Color = RGB(255, 255, 255)
            End If
            
            .Borders.LineStyle = xlContinuous ' �r����ݒ�
        End With
        
        ' E��� 0.00 "kN" �̏����ݒ�
        Set eRange = ws.Range("E" & currentRow)
        eRange.NumberFormat = "0 ""G"""
        eRange.HorizontalAlignment = xlRight ' �E��

        ' J��� "Insert + �s�ԍ�" �̈��t����
        Set iRange = ws.Range("L" & currentRow)
        iRange.value = "Insert " & currentRow
    Next currentRow
End Sub

' "���|�[�g�{��"�̕\�����Ƀe�L�X�g��}������B
Sub InsertTextToReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim insertTextA As String
    Dim insertTextB As String
    Dim foundOpen As Boolean

    ' ��`
    Set ws = ThisWorkbook.Sheets("���|�[�g�{��")
    insertTextA = "�p�����Ԃ͋K�i�l�𖞂����Ă��܂����B"
    insertTextB = "�A���r���Փˎ��ɊJ�����X�̂�����܂�"
    foundOpen = False

    ' L��̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).row
    Debug.Print lastRow
    ' L���"Insert"���܂܂��s��T�����AH���"�J"���܂܂�Ă��邩�m�F
    For i = 1 To lastRow
        If InStr(ws.Cells(i, "L").value, "Insert") > 0 Then
            If InStr(ws.Cells(i, "H").value, "�J") > 0 Then
                foundOpen = True
                Exit For
            End If
        End If
    Next i

    ' �\�̉����Ƀe�L�X�g��}��
    ws.Cells(lastRow + 1, "B").value = insertTextA
    If foundOpen Then
        ws.Cells(lastRow + 1, "B").value = insertTextA & " " & insertTextB
    End If
End Sub


