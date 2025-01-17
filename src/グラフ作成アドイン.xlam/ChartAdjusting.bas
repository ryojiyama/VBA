Attribute VB_Name = "ChartAdjusting"
'�O���t��Y���̍ő�l�𒲐�����B
Sub UniformizeLineGraphAxes()

    On Error GoTo ErrorHandler

    ' Display input dialog to set the maximum value for the axes
    Dim MaxValue As Variant
    MaxValue = InputBox("Y���̍ő�l����͂��Ă��������B(����)", "�ő�l�����")

    ' Check if the user pressed Cancel
    If MaxValue = False Then
        MsgBox "���삪�L�����Z������܂����B", vbInformation
        Exit Sub
    End If

    ' Validate the input
    If Not IsNumeric(MaxValue) Or MaxValue <= 0 Then
        MsgBox "�L���Ȑ��l����͂��Ă��������B", vbExclamation
        Exit Sub
    End If

    MaxValue = CDbl(MaxValue)

    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.Count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                With chartObj.chart.Axes(xlValue)
                    ' Set the Y-axis maximum value
                    .MaximumScale = MaxValue
                    
                    ' Set the MajorUnit based on MaxValue
                    If MaxValue <= 5 Then
                        .MajorUnit = 1#
                    ElseIf MaxValue > 5 And MaxValue <= 25 Then
                        .MajorUnit = 2#
                    ElseIf MaxValue > 25 And MaxValue <= 100 Then
                        .MajorUnit = 10#
                    ElseIf MaxValue > 100 And MaxValue <= 300 Then
                        .MajorUnit = 50#
                    ElseIf MaxValue > 300 Then
                        .MajorUnit = 100#
                    End If
                End With
            Next chartObj
        End If
    Next ws

    MsgBox "���ׂẴV�[�g�̃O���t��Y���̍ő�l�� " & MaxValue & " �ɐݒ肵�A�K�؂Ȗڐ���Ԋu��ݒ肵�܂����B", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical

End Sub




' �O���t�̏c���䗦��ύX����B
Sub AdjustChartAspectRatio()
    Dim userChoice As Variant

    ' ���[�U�[�ɃO���t�̔䗦��I��������
    userChoice = MsgBox("�O���t�̔䗦���A" & vbCrLf & _
                        "�ǂ̂悤�ɕύX���܂����H" & vbCrLf & vbCrLf & _
                        "[�͂�] - 2��p�ɕύX����" & vbCrLf & _
                        "[������] - 3��p�ɕύX����", vbYesNo + vbQuestion, "�䗦�̑I��")

    ' �I���ɉ����ăv���V�[�W�������s
    Select Case userChoice
        Case vbYes
            Call SetChartRatio129
        Case vbNo
            Call SetChartRatio1110
        Case Else
            Exit Sub ' �L�����Z�����ꂽ�ꍇ�͏������I��
    End Select
End Sub

' "Impact" ���܂ރV�[�g���̃O���t�䗦�� 480:360 �ɂ���v���V�[�W��
Sub SetChartRatio129()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    ' "Impact" ���܂ރV�[�g�����[�v����
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' �V�[�g���̃O���t�����[�v����
            For Each chartObj In ws.ChartObjects
                chartObj.Width = 480
                chartObj.Height = 360
            Next chartObj
        End If
    Next ws
End Sub

' "Impact" ���܂ރV�[�g���̃O���t�䗦�� 440:400 �ɂ���v���V�[�W��
Sub SetChartRatio1110()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    ' "Impact" ���܂ރV�[�g�����[�v����
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' �V�[�g���̃O���t�����[�v����
            For Each chartObj In ws.ChartObjects
                chartObj.Width = 400
                chartObj.Height = 440
            Next chartObj
        End If
    Next ws
End Sub

