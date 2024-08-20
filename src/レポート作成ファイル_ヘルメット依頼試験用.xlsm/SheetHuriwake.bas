Attribute VB_Name = "SheetHuriwake"
Sub TransferDataBasedOnID()
    Call Utlities.DeleteRowsBelowHeader

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim maxValue As Double, duration49kN As Double, duration73kN As Double
    Dim nextRow As Long
    Dim tempArray As Variant
    Dim data As Collection
    Dim dataItem As Variant
    
    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row

    ' �e�s�����[�v����
    For i = 1 To lastRow
        ' ID�𕪊�
        idParts = Split(wsSource.Cells(i, 3).value, "-")
        If UBound(idParts) >= 2 Then
            ' �O���[�v���i���ʁj���擾
            group = idParts(2)
            
            ' �O���[�v���Ɋ�Â��ăV�[�g����ݒ�
            Select Case group
                Case "�V"
                    targetSheetName = "Impact_Top"
                Case "�O"
                    targetSheetName = "Impact_Front"
                Case "��"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' �Ή�����O���[�v���Ȃ��ꍇ�̓X�L�b�v
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).value
                    GoTo NextIteration
            End Select
            
            groupName = "Group:" & idParts(0) & group
            maxValue = wsSource.Range("H" & i).value
            duration49kN = wsSource.Range("J" & i).value
            duration73kN = wsSource.Range("K" & i).value

            ' �O���[�v���ƃV�[�g���̑Ή����m�F
'            Debug.Print "Group: " & groupName & "; Sheet: " & targetSheetName
'            Debug.Print "Max Value: " & Format(maxValue, "0.00") & " 49kN Duration: " & Format(duration49kN, "0.00") & " 73kN Duration: " & Format(duration73kN, "0.00")

            ' �f�[�^���R���N�V�����ɒǉ�
            tempArray = Array( _
            groupName, _
            targetSheetName, _
            Format(maxValue, "0.00"), _
            Format(duration49kN, "0.00"), _
            Format(duration73kN, "0.00") _
            )
            data.Add tempArray
        End If
NextIteration:
    Next i
    
    ' �R���N�V��������e�V�[�g�Ƀf�[�^��]�L
    For Each dataItem In data
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)
        ' �ړI�̃V�[�g���쐬
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsDest.name = targetSheetName
        End If
        On Error GoTo 0
        
        ' �w�b�_�[�s��ݒ�i14�s�ځj
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "Max"
            wsDest.Range("C14").value = "4.9kN"
            wsDest.Range("D14").value = "7.3kN"
        End If
        nextRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        
        '�f�[�^��]�L
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = maxValue
        wsDest.Range("C" & nextRow).value = duration49kN
        wsDest.Range("D" & nextRow).value = duration73kN
    Next dataItem

    ' ���\�[�X�����
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub



