Attribute VB_Name = "ArrangementData"
Option Explicit
' Impact_Top�̃f�[�^����ג����B
Public Sub Rearrangement_Top()
    Dim copyDict As Object
    Dim cellValue As String
    Dim pseudoSpace As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Set copyDict = CreateObject("Scripting.Dictionary")
    
    ' �R�s�[���ƃR�s�[��̑Ή��֌W��ݒ�
    copyDict.Add "D16", "B2"
    copyDict.Add "H16", "C6"
    copyDict.Add "H17", "C8"
    copyDict.Add "H18", "C10"
    copyDict.Add "H19", "E6"
    copyDict.Add "H20", "E8"
    copyDict.Add "H21", "E10"
    copyDict.Add "H22", "G6"
    copyDict.Add "H23", "G8"
    copyDict.Add "H24", "G10"

    ' B16�̒l�����邩�ǂ������m�F
    If Not IsEmpty(ws.Range("B16").value) Then
        Dim key As Variant
        For Each key In copyDict.Keys
            ws.Range(copyDict(key)).value = ws.Range(key).value
        Next key
    End If

    ' B2�Z���ɒl�����邩�ǂ������m�F���A�l��ύX
    If Not IsEmpty(ws.Range("B2").value) Then
        pseudoSpace = ChrW(12288)
        cellValue = ws.Range("B2").value
        
        ' ���� "No." �� " AB" ���܂܂�Ă��Ȃ����m�F
        If Left(cellValue, 3) <> "No." Or Right(cellValue, 3) <> " AB" Then
            ws.Range("B2").value = "No." & cellValue & pseudoSpace & " AB"
        End If
    End If
    
    ' �f�B�N�V���i���̉��
    Set copyDict = Nothing
End Sub


' Impact_Front, Back�̃f�[�^����ג����B
Public Sub Rearrangement_FrontAndBack()
    Dim copyDict As Object
    Dim cellValue As String
    Dim pseudoSpace As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Set copyDict = CreateObject("Scripting.Dictionary")
    
    ' �i�Ԃ̐ݒ�
    copyDict.Add "D16", "B2"
    ' �Ռ��l�̐ݒ�
    copyDict.Add "H16", "C6"
    copyDict.Add "H17", "C9"
    copyDict.Add "H18", "C12"
    copyDict.Add "H19", "E6"
    copyDict.Add "H20", "E9"
    copyDict.Add "H21", "E12"
    copyDict.Add "H22", "G6"
    copyDict.Add "H23", "G9"
    copyDict.Add "H24", "G12"
    ' 4.9kN���̌p�����Ԃ̐ݒ�
    copyDict.Add "J16", "D6"
    copyDict.Add "J17", "D9"
    copyDict.Add "J18", "D12"
    copyDict.Add "J19", "F6"
    copyDict.Add "J20", "F9"
    copyDict.Add "J21", "F12"
    copyDict.Add "J22", "H6"
    copyDict.Add "J23", "H9"
    copyDict.Add "J24", "H12"
    ' 7.3kN���̌p�����Ԃ̐ݒ�
    copyDict.Add "K16", "D7"
    copyDict.Add "K17", "D10"
    copyDict.Add "K18", "D13"
    copyDict.Add "K19", "F7"
    copyDict.Add "K20", "F10"
    copyDict.Add "K21", "F13"
    copyDict.Add "K22", "H7"
    copyDict.Add "K23", "H10"
    copyDict.Add "K24", "H13"

    ' B16�̒l�����邩�ǂ������m�F
    If Not IsEmpty(ws.Range("B16").value) Then
        Dim key As Variant
        For Each key In copyDict.Keys
            ws.Range(copyDict(key)).value = ws.Range(key).value
        Next key
    End If
    
    ' B2�Z���ɒl�����邩�ǂ������m�F���A�l��ύX
    If Not IsEmpty(ws.Range("B2").value) Then
        pseudoSpace = ChrW(12288)
        cellValue = ws.Range("B2").value
        
        ' ���� "No." �� " AB" ���܂܂�Ă��Ȃ����m�F
        If Left(cellValue, 3) <> "No." Or Right(cellValue, 3) <> " AB" Then
            ws.Range("B2").value = "No." & cellValue & pseudoSpace & " AB"
        End If
    End If
    
    ' �f�B�N�V���i���̉��
    Set copyDict = Nothing
End Sub

'���ג�������̏����ݒ�
Public Sub ApplyFormattingCondition()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' B16�̒l�����邩�ǂ������m�F
    If Not IsEmpty(ws.Range("B16").value) Then
        ' �͈�A1:H13�̃t�H���g��ݒ�
        ws.Range("A1:H13").Font.name = "������"
        
        ' �w��Z����NumberFormat�ƃt�H���g�T�C�Y��ݒ�
        ws.Range("C6, C9, C12, E6, E9, E12, G6, G9, G12").NumberFormat = "0.00 ""kN"""
        ws.Range("C6, C9, C12, E6, E9, E12, G6, G9, G12").Font.size = 10
        
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").NumberFormat = """    4.90kN   ""0.0 ""ms"""
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").Font.size = 8
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").HorizontalAlignment = xlLeft
        
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").NumberFormat = """    7.30kN   ""0.0 ""ms"""
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").Font.size = 8
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").HorizontalAlignment = xlLeft
        
        ' �l��0�̏ꍇ��NumberFormat��ݒ�
        Dim rng As Range
        For Each rng In ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13")
            If rng.value = 0 Then
                rng.NumberFormat = """    7.30kN    �\ """
            End If
        Next rng
    End If
End Sub


