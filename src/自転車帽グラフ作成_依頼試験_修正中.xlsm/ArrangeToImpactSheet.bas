Attribute VB_Name = "ArrangeToImpactSheet"
Option Explicit


' �o���オ����"���|�[�g�O���t"�V�[�g�Ɋe�l��z�u����
Sub ArrangeDataByGroup()
    Dim wsName As String: wsName = "���|�[�g�O���t" ' �V�[�g���Ɋ܂܂�镔��������
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
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

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
'            Debug.Print "FirstRow:" & groupFirstRow
'            Debug.Print "Size:" & groupSize

            ' �O���[�v�T�C�Y�ɉ����ăf�[�^��z�u
            If groupResults = groupInsert Then
                Select Case groupSize
                ' groupRange.Cells:�}�������\��, ws.Cells:Group�ȉ��̌��ʈꗗ�\
                        Case 4  ' �O���[�v��2�̃��R�[�h������ꍇ
                            With groupRange
                                ' �w�b�_�[���i���ڊԂ�4�̃X�y�[�X��ݒ�j
                                .Cells(1, 1).value = "����No." & ws.Cells(groupFirstRow, "A").value & "    " & _
                                                     ws.Cells(groupFirstRow, "C").value & "    " & _
                                                     "�y�O�����z" & ws.Cells(groupFirstRow, "G").value & "    " & _
                                                     "�y���l�z" & ws.Cells(groupFirstRow, "H").value
                                
                                ' ����̃f�[�^
                                .Cells(1, 3).value = ws.Cells(groupFirstRow, "D").value       ' �����ӏ�
                                .Cells(2, 3).value = ws.Cells(groupFirstRow, "F").value       ' �A���r��
                                .Cells(1, 4).value = ws.Cells(groupFirstRow, "B").value       ' �Ռ��l
                                
                                ' �E��̃f�[�^
                                .Cells(1, 6).value = ws.Cells(groupFirstRow + 1, "D").value   ' �����ӏ�
                                .Cells(2, 6).value = ws.Cells(groupFirstRow + 1, "F").value   ' �A���r��
                                .Cells(1, 7).value = ws.Cells(groupFirstRow + 1, "B").value   ' �Ռ��l
                                
                                ' �����̃f�[�^
                                .Cells(4, 3).value = ws.Cells(groupFirstRow + 2, "D").value   ' �����ӏ�
                                .Cells(5, 3).value = ws.Cells(groupFirstRow + 2, "F").value   ' �A���r��
                                .Cells(4, 4).value = ws.Cells(groupFirstRow + 2, "B").value   ' �Ռ��l
                                
                                ' �E���̃f�[�^
                                .Cells(4, 6).value = ws.Cells(j - 1, "D").value              ' �����ӏ�
                                .Cells(5, 6).value = ws.Cells(j - 1, "F").value              ' �A���r��
                                .Cells(4, 7).value = ws.Cells(j - 1, "B").value              ' �Ռ��l
                                
                                'j��_�̈ʒu���킹�B�O�̈וۗ�
                                'groupRange.Cells(2, 3).Value = ws.Cells(j - 1, "C").Value & ws.Cells(j - 1, "D").Value
                            End With
                        
                        Case 3  ' �����̊g���p�ɗ\��
                        groupRange.Cells(2, 1).value = ws.Cells(groupFirstRow, "A").value
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
    ' groupRange�̊J�n�s���f�o�b�O�o��
    Debug.Print "groupRange�̊J�n�s: " & groupRange.row
    
    Dim ws As Worksheet
    Set ws = groupRange.Worksheet

    Dim rowIndex As Range
    Dim headerRange1 As Range
    Dim headerRange2 As Range
    Dim headerInput1 As Range
    Dim headerInput2 As Range
    Dim impactValue As Range
    Dim fontRange As Range

    ' ���[�N�V�[�g��̐�ΓI�ȃZ���͈͂��擾
    With ws
        ' groupRange�̊J�n�s��1��ڂ���7��ڂ܂ł͈̔�
        Set rowIndex = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 5, 1))
        Set headerRange1 = .Range(.Cells(groupRange.row, 1), .Cells(groupRange.row + 2, 7))
        Set headerRange2 = .Range(.Cells(groupRange.row + 3, 1), Cells(groupRange.row + 5, 7))
    
        ' headerInput1: 2��ڂ���6��ڂ̗�S��
        Set headerInput1 = .Columns("B:F")
        ' headerInput2: 2��ڂ�5��ڂ̒P�Ƃ̗�S��
        Set headerInput2 = Union(.Columns("B"), .Columns("E"))
        Set impactValue = Union(.Cells(groupRange.row, 4), .Cells(groupRange.row, 7), .Cells(groupRange.row + 3, 4), .Cells(groupRange.row + 3, 7))
    End With


    ' �͈͍쐬�T���v��_�ۗ����Ă����B
'    Set fontRange = Union(headerRange, leftColumnRange1)

    ' fontRange1 �ɑ΂��ď����ݒ�
    With headerRange1.Font
        .Name = "UDEV Gothic"
        .Color = RGB(60, 60, 60) ' �t�H���g�̐F�𔒂ɐݒ�
    End With
    With headerRange2.Font
        .Name = "UDEV Gothic"
        .Color = RGB(60, 60, 60) ' �t�H���g�̐F�𔒂ɐݒ�
    End With
    
    With headerInput1
        .HorizontalAlignment = xlCenter ' ���������̒�������
        .VerticalAlignment = xlCenter ' ���������̒�������
    End With
    With headerInput2
        .HorizontalAlignment = xlLeft ' ���������̒�������
        .VerticalAlignment = xlCenter ' ���������̒�������
    End With
    
    With impactValue
        .NumberFormat = "0"" G""" ' ���l�t�H�[�}�b�g 0.0"G" ��ǉ�
    End With
    
    With rowIndex.Font
        .Color = RGB(230, 230, 230)
    End With
    With rowIndex.Interior
        .Color = RGB(48, 84, 150) ' �w�i�F��ɐݒ�
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
' �w�b�_�[�ݒ�p�̃T�u�v���V�[�W��


' �w�b�_�[�ݒ�p�̓Ɨ������v���V�[�W��
Sub SetupSheetHeader()
    Dim ws As Worksheet
    
    ' "���|�[�g�O���t"�V�[�g�̑��݊m�F
    If WorksheetExists("���|�[�g�O���t") = False Then
        MsgBox "���|�[�g�O���t�V�[�g��������܂���B", vbExclamation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' �����̌����Z�����N���A�i�G���[�h�~�̂��߁j
    On Error Resume Next
    ws.Range("A1:B2").UnMerge
    ws.Range("C1:E2").UnMerge
    On Error GoTo ErrorHandler
    
    With ws
        .Rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' �Z���̌���
        .Range("A1:B2").Merge
        .Range("C1:E2").Merge
        
        ' ���e�̋L��
        .Cells(1, "A").value = "�\��"
        .Cells(1, "F").value = "�쐬��"
        .Cells(2, "F").value = "�쐬��"
        .Cells(1, "G").value = Format(Date, "yyyy/mm/dd")
        
        ' ��{�̏����ݒ�
        With .Range("A1:G2")
            .HorizontalAlignment = xlCenter    ' ��������
            .VerticalAlignment = xlCenter      ' �㉺��������
            .WrapText = True                   ' �܂�Ԃ��ĕ\��
            
            ' �t�H���g�ݒ�
            With .Font
                .Name = "���S�V�b�N"
                .Size = 10
            End With
        End With
        
        ' �r���ݒ�
        With .Range("A1:G2")
            ' ���ׂĂ̌r������U�N���A
            .Borders.LineStyle = xlNone
            
            ' �O�g�̌r���i�א��j
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            
            ' �����̌r���i�ɍא��j
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlHairline
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlHairline
            End With
        End With
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "�w�b�_�[�̐ݒ肪�������܂����B", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "�w�b�_�[�̐ݒ蒆�ɃG���[���������܂����B" & vbNewLine & _
           "�G���[�̏ڍ�: " & Err.Description, vbCritical
End Sub

' �V�[�g�̑��݊m�F�֐�
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    WorksheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function


' "LOG_Bicycel"�V�[�g�̃`���[�g��"���|�[�g�O���t"�V�[�g�Ɉړ�����B
' �`���[�g�̏o���ʒu�̓T�u���[�`���Őݒ肵�Ă���B
Sub MoveChartsFromLOGToReport()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim chartObj As ChartObject
    Dim groupCell As Range
    Dim targetTop As Double
    Dim targetLeft As Double
    Dim offsetX As Double
    Dim offsetY As Double
    Dim i As Integer
    Dim recordID As String
    Dim idNumber As Integer ' ID�̍ŏ��̐��l����
    Dim chartHeight As Double
    Dim chartWidth As Double
    Dim previousLeft As Double
    previousLeft = 0

    ' �V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Set wsTarget = ThisWorkbook.Sheets("���|�[�g�O���t")

    ' A���"Group"�Ƃ����l��T��
    Set groupCell = wsTarget.Columns("A").Find(What:="Group", LookIn:=xlValues, LookAt:=xlWhole)
    If groupCell Is Nothing Then
        MsgBox "���|�[�g�O���t�V�[�g��A���'Group'��������܂���B", vbExclamation
        Exit Sub
    End If

    ' �`���[�g�̐ݒu��̃I�t�Z�b�g�i�s�N�Z���P�ʁj
    offsetY = 30 ' �������30�s�N�Z��
    offsetX = 10 ' �e�`���[�g���E������10�s�N�Z�������炷

    ' �c����1:2�ɐݒ肷�邽�߂̃T�C�Y
    chartHeight = 200 ' ������200�s�N�Z���ɐݒ�
    chartWidth = chartHeight * 2 ' ���͍�����2�{�ɐݒ�

    ' �`���[�g�̈ړ�
    i = 0
    For Each chartObj In wsSource.ChartObjects
        ' �`���[�g�̃^�C�g������ID���擾
        If chartObj.chart.HasTitle Then
            recordID = Replace(chartObj.chart.chartTitle.Text, "ID: ", "") ' "ID: "����������ID�̂ݎ擾
            Debug.Print "recordID: " & recordID ' �C�~�f�B�G�C�g�E�B���h�E��ID���o��

            ' ID�̍ŏ��̐��l�����𒊏o
            idNumber = CInt(Split(recordID, "-")(0)) ' recordID�̍ŏ��̕����𐔒l��

            ' �`���[�g�̈ʒu���T�u�v���V�[�W���Őݒ�
            Call SetChartPosition(idNumber, i, groupCell.Left, targetTop, targetLeft, previousLeft)
            ' �`���[�g���R�s�[
            chartObj.Copy

            ' �R�s�[�̃^�C�����O���쐬
            WaitHalfASecond

            ' ���|�[�g�O���t�V�[�g���A�N�e�B�u�ɂ��āA�`���[�g��\��t��
            wsTarget.Activate
            wsTarget.Paste

            ' �\��t����ꂽ�`���[�g�̃I�u�W�F�N�g���擾
            With wsTarget.ChartObjects(wsTarget.ChartObjects.Count)
                ' �`���[�g�̈ʒu��ݒ�
                .Top = targetTop
                .Left = targetLeft

                ' �`���[�g�̃T�C�Y��ݒ� (�c���� 1:2)
                .Height = chartHeight
                .Width = chartWidth
                ' �`���[�g�̈ʒu��ݒ�
                Call SetChartPosition(idNumber, i, groupCell.Left, targetTop, targetLeft, previousLeft)
                previousLeft = targetLeft
            End With

            ' ���̃`���[�g�ʒu���E�ɂ��炷
            i = i + 1
        End If
    Next chartObj

    MsgBox "�`���[�g�̈ړ����������܂����B", vbInformation

    ' ���������
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set chartObj = Nothing
    Set groupCell = Nothing
End Sub

Sub SetChartPosition(ByVal idNumber As Integer, ByVal chartIndex As Integer, ByVal groupLeft As Double, ByRef targetTop As Double, ByRef targetLeft As Double, ByVal previousLeft As Double)
    ' idNumber �Ɋ�Â��č����𓮓I�Ɍv�Z
    targetTop = 100 + (idNumber - 1) * 200 ' idNumber�������邲�Ƃ�200�s�N�Z����������ς���
    
    ' �������̈ʒu��chartIndex �Ɋ�Â��Ē���
    If targetLeft <= previousLeft Then
        targetLeft = previousLeft + 15 ' ���������̏ꍇ�ł����X�ɉ������ɂ��炷
    Else
        targetLeft = groupLeft + 400 + (chartIndex Mod 9) * 5 ' 9�̃`���[�g���Ƃɉ��ɂ��炷
    End If
    
    Debug.Print "Index:"; chartIndex & " targetTop:"; targetTop & " targetLeft:"; targetLeft
End Sub

Sub WaitHalfASecond()
    Dim start As Single
    start = Timer
    Do While Timer < start + 0.4
        DoEvents ' ���\�[�X���
    Loop
End Sub




' �f�[�^�U�蕪���̊m�F���f�o�b�O�E�C���h�E�ōs���B�ڍs�����r���[�ŏI����Ă���B
Sub ConsolidateData()

  ' �V�[�g���̕ϐ����`
  Dim wsName As String: wsName = "���|�[�g�O���t" ' �V�[�g���̈ꕔ���w��
  Dim wsResult As Worksheet
  Dim i As Long, j As Long
  Dim groupInsert As Variant, groupResults As Variant
  Dim insertNum As Long, resultNum As Long
  Dim dict As Object

  ' "���|�[�g�O���t"���܂ރV�[�g�����ɏ���
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
          
          ' �f�o�b�O: �����ɒǉ������f�[�^���m�F
          Debug.Print "Adding value to group: " & groupInsert & ", Value: " & wsResult.Cells(i, "C").value
          
          dict(groupInsert).Add wsResult.Cells(i, "C").value
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
