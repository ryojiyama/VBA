Attribute VB_Name = "Hel_RequestedTestArrange"


' �˗������p��LOG_Helmet�ɐV����ID���쐬����B
Sub GenereteRequestsID()

    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim id As String

    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row

    ' �e�s�ɑ΂���ID�𐶐�
    For i = 2 To lastRow ' 1�s�ڂ̓w�b�_�Ɖ���
        id = GenerateID(ws, i)
        ' B���ID���Z�b�g
        ws.Cells(i, 2).value = id
    Next i
End Sub

Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' GenereteRequestsID()�̃T�u�v���V�[�W��
    Dim id As String

    ' C��: 2���ȉ��̐���
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C���D��̊Ԃ�"-"
    ' D��̏�����ύX
    id = id & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    id = id & "-"
    id = id & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E��̏���
    id = id & "-"
    id = id & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L��̏���

    GenerateID = id
End Function
Function ExtractNumberWithF(value As String) As String
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' ���K�\���I�u�W�F�N�g�̍쐬
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{3,6}" ' 1���ȏ�̐����𒊏o
    regex.Global = True

    ' ���������𒊏o
    Set matches = regex.Execute(value)
    If matches.count > 0 Then
        numPart = matches(0).value ' �ŏ��Ɍ��������������擾
    Else
        numPart = "000000" ' �f�t�H���g�l�܂��̓G���[�n���h�����O
    End If

    ' "F"�̑��݃`�F�b�N
    hasF = InStr(value, "F") > 0

    ' F������ꍇ�͐����̌��"F"������
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
        Dim parts() As String
        parts = Split(value, "_")

        If UBound(parts) >= 1 Then
            Dim angle As String
            Dim direction As String

            ' �p�x�𒊏o
            angle = Replace(parts(0), "����", "")

            ' �����𒊏o�Ɛ��`
            direction = parts(1)
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

Function GetColumnLValue(value As Variant) As String
' GenerateID�̃T�u�֐�
    Select Case value
        Case "����"
            GetColumnLValue = "Hot"
        Case "�ቷ"
            GetColumnLValue = "Cold"
        Case "�Z����"
            GetColumnLValue = "Wet"
        Case Else
            GetColumnLValue = "?"
    End Select
End Function

' �O���[�v���ƂɐF����_�O���[�v�����m�ɂł��Ă��邩���m�F����
Sub ColorGroups()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentColorIndex As Long
    Dim i As Long
    Dim currentGroup As String
    Dim previousGroup As String
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("LOG_Helmet") ' �V�[�g����K�v�ɉ����ĕύX
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
    
    ' �����ݒ�
    currentColorIndex = 1
    previousGroup = ""
    
    Dim colorArray(1 To 20) As Long
    colorArray(1) = RGB(204, 255, 255) ' Light Cyan
    colorArray(2) = RGB(255, 204, 204) ' Light Red
    colorArray(3) = RGB(204, 255, 204) ' Light Green
    colorArray(4) = RGB(255, 255, 204) ' Light Yellow
    colorArray(5) = RGB(204, 204, 255) ' Light Blue
    colorArray(6) = RGB(255, 229, 204) ' Light Orange
    colorArray(7) = RGB(204, 255, 229) ' Light Aqua
    colorArray(8) = RGB(229, 204, 255) ' Light Purple
    colorArray(9) = RGB(255, 204, 229) ' Light Pink
    colorArray(10) = RGB(255, 255, 153) ' Light Yellow 2
    colorArray(11) = RGB(204, 255, 153) ' Light Lime
    colorArray(12) = RGB(153, 204, 255) ' Light Sky Blue
    colorArray(13) = RGB(255, 204, 153) ' Light Peach
    colorArray(14) = RGB(204, 153, 255) ' Light Lavender
    colorArray(15) = RGB(255, 153, 204) ' Light Rose
    colorArray(16) = RGB(204, 255, 255) ' Light Mint
    colorArray(17) = RGB(255, 255, 204) ' Light Cream
    colorArray(18) = RGB(204, 229, 255) ' Light Denim
    colorArray(19) = RGB(255, 204, 255) ' Light Fuchsia
    colorArray(20) = RGB(255, 204, 229) ' Light Rose 2

    ' �O���[�v���ƂɐF������
    For i = 2 To lastRow
        currentGroup = ws.Cells(i, 3).value
        
        ' �O���[�v���ς�����玟�̐F�ɐ؂�ւ�
        If currentGroup <> previousGroup Then
            currentColorIndex = currentColorIndex + 1
            If currentColorIndex > UBound(colorArray) Then
                currentColorIndex = 1
            End If
        End If
        
        ' B�񂩂�E��ɐF��ݒ�
        ws.Range(ws.Cells(i, 2), ws.Cells(i, 5)).Interior.color = colorArray(currentColorIndex)
        
        ' ���݂̃O���[�v���L�^
        previousGroup = currentGroup
    Next i
End Sub

