Attribute VB_Name = "TransferToReport"
' ���|�[�g�{���̕\�Ɍ��ʂ�}������B�w�����b�g�̂��̂����̂܂܃R�s�[�C�����K�v
Sub TransferDataWithMappingAndFormatting()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim i As Long
    Dim destRow As Long
    Dim transferredRows As Long
    Const MAX_ROWS As Long = 12 ' �ő�]�L�s����12�ɐݒ�
    Dim startRow As Long
    Dim mappingDict As Object ' �}�b�s���O�p�̃f�B�N�V���i��

    ' �V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet") ' �]�L���V�[�g
    Set wsDest = ThisWorkbook.Sheets("���|�[�g�{��") ' �]�L��V�[�g

    ' �]�L���̍ŏI�s���擾
    lastRowSource = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).row

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
    ' �]�L���Ɠ]�L��̃}�b�s���O���f�B�N�V���i���Őݒ�
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' �]�L���̗� �� �]�L��̗��񖼂Ƃ��Ė����I�ɋL�q
    dict.Add "D", "B" ' �]�L����D�� �� �]�L���B��
    dict.Add "E", "C" ' �]�L����E�� �� �]�L���C��
    dict.Add "L", "D" ' �]�L����L�� �� �]�L���D��
    dict.Add "H", "E" ' �]�L����H�� �� �]�L���E��
    dict.Add "M", "F" ' �]�L����M�� �� �]�L���F��
    dict.Add "N", "G" ' �]�L����N�� �� �]�L���G��
    ' �K�v�ɉ����đ��̗�̃}�b�s���O���ǉ�

    Set GetMappingDictionary = dict
End Function

Private Sub TransferMappedValues(wsSource As Worksheet, wsDest As Worksheet, sourceRow As Long, destRow As Long, mappingDict As Object)
    ' �}�b�s���O�Ɋ�Â��Ēl��]�L����
    Dim key As Variant

    ' �}�b�s���O�f�B�N�V���i�������[�v���ē]�L�����s
    For Each key In mappingDict.Keys
        wsDest.Cells(destRow, Columns(mappingDict(key)).column).value = wsSource.Cells(sourceRow, Columns(key).column).value
    Next key
End Sub
Private Sub ApplyFormattingToNewRows(ws As Worksheet, startRow As Long, endRow As Long)
    ' �V�����ǉ����ꂽ�s�Ƀt�H�[�}�b�g��K�p���AI��Ɉ������
    Dim currentRow As Long
    Dim targetRange As Range
    Dim eRange As Range, fRange As Range, gRange As Range, iRange As Range
    
    ' 1�s������
    For currentRow = startRow To endRow
        ' ���݂̍s�͈̔͂��擾
        Set targetRange = ws.Range("B" & currentRow & ":G" & currentRow)
        
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
        eRange.NumberFormat = "0.00 ""kN"""
        eRange.HorizontalAlignment = xlRight ' �E��

        ' F��� 0.0 "g" �̏����ݒ�
        Set fRange = ws.Range("F" & currentRow)
        fRange.NumberFormat = "0.0 ""g"""
        fRange.HorizontalAlignment = xlRight ' �E��

        ' G��� 0.0 "mm" �̏����ݒ�
        Set gRange = ws.Range("G" & currentRow)
        gRange.NumberFormat = "0.0 ""mm"""
        gRange.HorizontalAlignment = xlRight ' �E��

        ' I��� "Insert + �s�ԍ�" �̈��t����
        Set iRange = ws.Range("I" & currentRow)
        iRange.value = "Insert " & currentRow
    Next currentRow
End Sub





