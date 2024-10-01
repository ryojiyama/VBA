Attribute VB_Name = "ArrangeData"
Option Explicit


' �f�[�^�U�蕪���̊m�F���f�o�b�O�E�C���h�E�ōs��
Sub ConsolidateData()

  ' �V�[�g���̕ϐ����`
  Dim wsName As String: wsName = "Impact" ' �V�[�g���̈ꕔ���w��
  Dim wsResult As Worksheet
  Dim i As Long, j As Long
  Dim groupInsert As Variant, groupResults As Variant
  Dim insertNum As Long, resultNum As Long
  Dim dict As Object

  ' "Impact"���܂ރV�[�g�����ɏ���
  For Each wsResult In ThisWorkbook.Worksheets
    If InStr(wsResult.Name, wsName) > 0 Then

      ' �ŏI�s���擾 (A���I��̗����Œl�����͂���Ă���Ō�̍s���擾)
      Dim lastRow As Long: lastRow = wsResult.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

      ' �O���[�v���ƂɃf�[�^���i�[���鎫����������
      Set dict = CreateObject("Scripting.Dictionary")

      ' I���A��̒l�Ɋ�Â��ăO���[�v�����A�����Ɋi�[
      For i = 2 To lastRow
        groupInsert = GetGroupNumber(wsResult.Cells(i, "I").value)
        groupResults = GetGroupNumber(wsResult.Cells(i, "A").value)

        ' �f�o�b�O: I���A��̒l�A�O���[�v�ԍ���\��
        Debug.Print "Sheet: " & wsResult.Name & ", Row: " & i & ", I Column: " & wsResult.Cells(i, "I").value & ", GroupInsert: " & groupInsert
        Debug.Print "Sheet: " & wsResult.Name & ", Row: " & i & ", A Column: " & wsResult.Cells(i, "A").value & ", GroupResults: " & groupResults

        ' �O���[�v����v���A����GroupInsert��GroupResults�������Ƃ�Null�łȂ��ꍇ�A�����Ƀf�[�^��ǉ�
        If Not IsNull(groupInsert) And Not IsNull(groupResults) And groupInsert = groupResults Then
          If Not dict.Exists(groupInsert) Then
            dict.Add groupInsert, New Collection
          End If
          dict(groupInsert).Add wsResult.Cells(i, "D").value
        End If
      Next i

      ' �f�o�b�O: �O���[�v���Ƃ�D��̒l��\��
      Debug.Print "Sheet: " & wsResult.Name & ", Grouped Data:"
      For Each groupInsert In dict.Keys
        Debug.Print "Group: " & groupInsert & ", Values: ";
        ' Collection�̗v�f�����[�v�������邽�߂̕ϐ����`
        Dim item As Variant
        For Each item In dict(groupInsert)
          Debug.Print item & " ";
        Next item
        Debug.Print
      Next groupInsert

    End If
  Next wsResult

End Sub

' �o���オ����"impact"�V�[�g�Ɋe�l��z�u����
Sub ArrangeDataByGroup()
    Dim wsName As String: wsName = "Impact" ' �V�[�g���Ɋ܂܂�镔��������
    Dim wsResult As Worksheet
    Dim lastRow As Long

    ' "Impact"���܂ނ��ׂẴ��[�N�V�[�g�����[�v
    For Each wsResult In ThisWorkbook.Worksheets
        If InStr(wsResult.Name, wsName) > 0 Then
            ' ���[�N�V�[�g�̍ŏI�g�p�s���擾
            lastRow = wsResult.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

            ' ��I�Ɋ�Â��ăO���[�v������
            ProcessGroupsInColumnI wsResult, lastRow
        End If
    Next wsResult
    Call InsertTextInMergedCells
End Sub

Private Sub ProcessGroupsInColumnI(ws As Worksheet, lastRow As Long)
    'ArrangeDataByGroup�̃T�u�v���V�[�W���BI��̒l�ŃO���[�v���쐬
    Dim firstRow As Long: firstRow = 2
    Dim groupInsert As Variant
    Dim i As Long
    Dim groupRange As Range

    ' ��I�̃O���[�v����肷�邽�߂Ɋe�s�����[�v
    Do While firstRow <= lastRow
        groupInsert = GetGroupNumber(ws.Cells(firstRow, "I").value)

        ' �O���[�v�ԍ����󔒂△���̏ꍇ�͎��̍s�ɐi��
        If Not IsNull(groupInsert) And groupInsert <> "" Then

            ' ���݂̃O���[�v�̍ŏI�s��������
            For i = firstRow + 1 To lastRow
                ' I��̒l���󔒂̏ꍇ�A���[�v���I��
                If ws.Cells(i, "I").value = "" Then Exit For
                ' I��̒l�����̃O���[�v�ɕς�����烋�[�v���I��
                If GetGroupNumber(ws.Cells(i, "I").value) <> groupInsert Then Exit For
            Next i

            ' �f�o�b�O: �O���[�v�͈̔͂��o��
            ' Debug.Print "�O���[�v�͈�: A" & firstRow & ":G" & i - 1

            ' ���݂̃O���[�v�͈̔͂�ݒ�
            Set groupRange = ws.Range("A" & firstRow & ":G" & i - 1)

            ' ��A�Ɋ�Â��ăO���[�v������
            ProcessGroupsInColumnA ws, groupInsert, groupRange

            ' ���̃O���[�v��
            firstRow = i
        Else
            ' groupInsert��Null�܂��͋�̏ꍇ�A���̍s��
            firstRow = firstRow + 1
        End If
    Loop
End Sub


Private Sub ProcessGroupsInColumnA(ws As Worksheet, groupInsert As Variant, groupRange As Range)
    ' ArrangeDataByGroup�̃T�u�v���V�[�W���B��I�Ɋ�Â��ăO���[�v������
    Dim groupFirstRow As Long: groupFirstRow = 2
    Dim groupResults As Variant
    Dim j As Long
    Dim lastRowA As Long

    ' ��A�̍ŏI�g�p�s���擾
    lastRowA = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    ' ��A�̃O���[�v����肷�邽�߂Ɋe�s�����[�v
    Do While groupFirstRow <= lastRowA
        groupResults = GetGroupNumber(ws.Cells(groupFirstRow, "A").value)

        ' �O���[�v�ԍ����󔒂łȂ��ꍇ�ɏ������s��
        If Not IsNull(groupResults) And groupResults <> "" Then
            ' ���݂̃O���[�v�̍ŏI�s��������
            For j = groupFirstRow + 1 To lastRowA + 1
                If j > lastRowA Or GetGroupNumber(ws.Cells(j, "A").value) <> groupResults Then Exit For
            Next j

            ' �O���[�v�̃T�C�Y���v�Z
            Dim groupSize As Long
            groupSize = j - groupFirstRow

            ' �O���[�v�T�C�Y�ɉ����ăf�[�^��z�u
            If groupResults = groupInsert Then
                Select Case groupSize
                    Case 2 ' �O���[�v��2�̃��R�[�h������ꍇ
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
                        groupRange.Cells(3, 2).value = ws.Cells(groupFirstRow, "E").value
                        groupRange.Cells(3, 5).value = ws.Cells(j - 1, "E").value
                        groupRange.Cells(1, 2).value = ws.Cells(groupFirstRow, "B").value & ws.Cells(groupFirstRow, "C").value
                        groupRange.Cells(1, 5).value = ws.Cells(j - 1, "B").value & ws.Cells(j - 1, "C").value
                    Case 3 ' �O���[�v��3�̃��R�[�h������ꍇ
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
                        groupRange.Cells(3, 2).value = ws.Cells(groupFirstRow, "E").value
                        groupRange.Cells(3, 4).value = ws.Cells(groupFirstRow + 1, "E").value
                        groupRange.Cells(3, 6).value = ws.Cells(j - 1, "E").value
                        groupRange.Cells(1, 2).value = ws.Cells(groupFirstRow, "B").value & ws.Cells(groupFirstRow, "C").value
                        groupRange.Cells(1, 4).value = ws.Cells(groupFirstRow + 1, "B").value & ws.Cells(groupFirstRow + 1, "C").value
                        groupRange.Cells(1, 6).value = ws.Cells(j - 1, "B").value & ws.Cells(j - 1, "C").value
                End Select

                ' �Z���̏�����ݒ�
                Call FormatGroupRange(groupRange)
            End If

            ' ���̃O���[�v��
            groupFirstRow = j
        Else
            ' groupResults��Null�܂��͋󔒂̏ꍇ�A���̍s��
            groupFirstRow = groupFirstRow + 1
        End If
    Loop
End Sub


Private Sub FormatGroupRange(groupRange As Range)
    ' �Z���̏����ݒ���s��ProcessGroupsInColumnA�̃T�u���[�`��
    Dim ws As Worksheet
    Set ws = groupRange.Worksheet

    Dim headerRange As Range
    Dim leftColumnRange As Range
    Dim fontRange As Range

    ' ���[�N�V�[�g��̐�ΓI�ȃZ���͈͂��擾
    With ws
        ' groupRange�̊J�n�s��1��ڂ���7��ڂ܂ł͈̔�
        Set headerRange = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row, 7))
        ' groupRange�̊J�n�s����2�s���܂ł�1��ڂ͈̔�
        Set leftColumnRange = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 2, 1))
    End With

    ' ��L�͈̔͂�����
    Set fontRange = Union(headerRange, leftColumnRange)

    ' �Z���F��RGB(48,84,150)�ɐݒ�
    With headerRange.Interior
        .color = RGB(48, 84, 150)
    End With

    With leftColumnRange.Interior
        .color = RGB(48, 84, 150)
    End With

    ' �t�H���g��"UDEV Gothic"�ɐݒ肵�A�t�H���g�̐F�𔒂ɐݒ�
    With fontRange.Font
        .Name = "UDEV Gothic"
        .color = RGB(255, 255, 255) ' �t�H���g�̐F�𔒂ɐݒ�
    End With

End Sub

Private Function GetGroupNumber(cellValue As String) As Variant
    ' ArrangeDataByGroup�̃T�u�v���V�[�W���BI��̒l�����̃O���[�v�ɕς�����烋�[�v���I��
    Dim regex As Object, matches As Object
    Dim result As String

    ' �󔒃Z����Group�̏���
    If cellValue = "" Or cellValue = "Group" Then
        GetGroupNumber = Null
        Exit Function
    End If

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "\D" ' �����ȊO�̕����Ƀ}�b�`����p�^�[��
    End With

    ' �����ȊO�̕������󕶎��ɒu��
    result = regex.Replace(cellValue, "")

    ' ���ʂ��󕶎��łȂ��ꍇ�A������Ԃ�
    If result <> "" Then
        GetGroupNumber = result
    Else
        GetGroupNumber = Null ' ������������Ȃ��ꍇ��Null��Ԃ�
    End If
End Function




' �e�V�[�g�ɕ\�������
Sub InsertTextInMergedCells()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim textToInsert As String
    Dim sheetDict As Object

    ' �V�[�g���ƑΉ�����e�L�X�g�̎������쐬
    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "Impact_Top", "�V�����Ռ�����"
    sheetDict.Add "Impact_Front", "�O�����Ռ�����"
    sheetDict.Add "Impact_Back", "�㓪���Ռ�����"
    sheetDict.Add "Impact_Side", "�������Ռ�����"

    ' �e�V�[�g�����[�v
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        
        ' �V�[�g����"Impact"���܂܂�Ă���ꍇ�̂ݏ���
        If InStr(sheetName, "Impact") > 0 Then
            ' �V�[�g���Ɋ�Â��đ}������e�L�X�g������
            If sheetDict.Exists(sheetName) Then
                textToInsert = sheetDict(sheetName)
                
                ' Cells(1,2)~Cells(1,7)������
                With ws.Range(ws.Cells(1, 2), ws.Cells(1, 7))
                    .Merge
                    .value = textToInsert ' �V�[�g���ɑΉ�����e�L�X�g��}��
                    .HorizontalAlignment = xlCenter ' �e�L�X�g�𒆉�����
                    .VerticalAlignment = xlCenter   ' �e�L�X�g���c��������
                    .Font.Name = "���S�V�b�N"         ' �t�H���g��"���S�V�b�N"�ɐݒ�
                    .Font.size = 20                   ' �t�H���g�T�C�Y��20�ɐݒ�
                    .Font.Bold = True
                End With
                
                ' �s�̍�����50�ɐݒ�
                ws.Rows(1).RowHeight = 50
            End If
        End If
    Next ws
End Sub

' �`���[�g���e�V�[�g�ɕ��z����B
Sub DistributeChartsToRequestedSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
    Dim key As Variant
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"�V�[�g��Ώۂɂ���
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"�V�[�g�̃`���[�g�I�u�W�F�N�g���O���[�v����
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.Text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartName��"-"�ŕ������AsheetName���擾
        parts = Split(chartObj.Name, "-")
        If UBound(parts) >= 2 Then
            ' sheetName�����ۂ̃V�[�g���ɕϊ�
            Select Case parts(2)
                Case "�V"
                    sheetName = "Impact_Top"
                Case "�O"
                    sheetName = "Impact_Front"
                Case "��"
                    sheetName = "Impact_Back"
                Case Else
                    sheetName = parts(2) ' ����ȊO�̏ꍇ�͂��̂܂�
            End Select
        Else
            sheetName = parts(0)
        End If
        
        ' ���ۂ̃V�[�g�����L�[�Ƃ��ăf�B�N�V���i���ɒǉ�
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' �O���[�v���ƂɃ`���[�g��Ή�����V�[�g�ɃR�s�[
    For Each key In groups.Keys
        ' �V�[�g�̑��݂��m�F
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' �V�[�g�����݂��Ȃ��ꍇ�A�`���[�g���R�s�[���Ȃ�
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' �`���[�g�̃R�s�[
            Dim chart As ChartObject
            Dim newChart As ChartObject
            For Each chart In groups(key)
                ' �`���[�g���R�s�[
                chart.Copy
                WaitHalfASecond ' 0.5�b�ҋ@
                ' �R�s�[�����`���[�g��\��t���A�߂�l�Ƃ��ĐV�����`���[�g�I�u�W�F�N�g���擾
                targetSheet.Paste
                Set newChart = targetSheet.ChartObjects(targetSheet.ChartObjects.count)
                
                ' ���̃`���[�g�̈ʒu�Ɋ�Â��āA�E���ɑ��ΓI�Ɉړ�
                With newChart
                    .Top = chart.Top + 50  ' ���̃`���[�g�̈ʒu����50�|�C���g���Ɉړ�
                    .Left = chart.Left + 100 ' ���̃`���[�g�̈ʒu����100�|�C���g�E�Ɉړ�
                End With
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not copied."
        End If
    Next key
End Sub

Sub WaitHalfASecond()
    Dim start As Single
    start = Timer
    Do While Timer < start + 0.4
        DoEvents ' ���\�[�X���
    Loop
End Sub

