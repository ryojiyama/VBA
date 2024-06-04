Attribute VB_Name = "CloseAndSave"
Sub CloseAndSave()
    ' �m�F�̃��b�Z�[�W�{�b�N�X��\��
    Dim Response As VbMsgBoxResult
    Response = MsgBox("�ǂݍ��񂾃V�[�g�Ƃ��ׂẴO���t����������܂��B", vbOKCancel + vbQuestion, "�m�F")

    ' OK�������ꂽ�ꍇ�ɏ��������s
    If Response = vbOK Then
        Call DeleteAllChartsAndSheets
        Call SetRowHeightAndColumnWidth
        MsgBox "�������������܂����B", vbInformation, "���슮��"
    End If
End Sub


Sub DeleteAllChartsAndSheets()
    ' �V�[�g���̃O���t�Ɨ]�v�ȃV�[�g���폜
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String

    ' �ی삷��V�[�g�̃��X�g
    Dim protectSheets As Variant
    protectSheets = Array("LOG_Helmet", "Setting", "Hel_SpecSheet", "Penetration", "Impact_Top", "Impact_Front", "Impact_Back", "Impact_Side")

    ' �x���\���𖳌���
    Application.DisplayAlerts = False

    ' �e�V�[�g�ɑ΂��ď��������s
    For Each sheet In ThisWorkbook.Sheets
        sheetName = sheet.name
        ' �ی삷��V�[�g�ȊO�̏ꍇ�A�V�[�g���폜
        If Not IsInArray(sheetName, protectSheets) Then
            sheet.Delete
        End If
    Next sheet

    ' �x���\�������ɖ߂�
    Application.DisplayAlerts = True

    ' �u�b�N��ۑ�
    ThisWorkbook.Save
End Sub



' DeleteAllChartsAndSheets_�z����ɓ���̒l�����݂��邩�`�F�b�N����֐�
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub SetRowHeightAndColumnWidth()
    ' A1�̕��ƍ�����20�ɂ���B
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant

    ' �ݒ��K�p����V�[�g���̃��X�g���`����
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    ' �V�[�g���̔z������[�v����
    For Each sheetName In sheetNames
        ' �V�[�g�������̃��[�N�u�b�N�ɑ��݂���ꍇ�A�s�̍����Ɨ�̕���ݒ肷��
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Range("A1").RowHeight = 20
            ws.Range("A1").ColumnWidth = 20
            Set ws = Nothing
        End If
    Next sheetName
End Sub

