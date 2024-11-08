Attribute VB_Name = "TransferDraftDatatoSheet"
' "���|�[�g�O���t"�Ȃǂ̃V�[�g���쐬���A�l��]�L����v���V�[�W��
Sub TransferDataBasedOnID()
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim preProcess As String ' �O����
    Dim anvilType As String ' �A���r��
    Dim dummyHead As String '�l���͌^
    Dim testPoint As String '�����ӏ�
    Dim sampleName As String
    Dim maxValue As Double
    Dim tempArray As Variant
    Dim data As Collection
    
    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    Set data = New Collection

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
    
    ' �e�s�����[�v����
    For i = 1 To lastRow
        ' ID�𕪊����A�K�v�ȏ����擾
        idParts = Split(wsSource.Cells(i, "B").value, "-")
        If UBound(idParts) >= 3 Then
            group = idParts(0)
        Else
            ' ID�`�����s���ȏꍇ�͎��̃��[�v��
            GoTo NextIteration
        End If
        
        ' �V�[�g�����O���[�v�Ɋ�Â��Č���
        targetSheetName = GetTargetSheetName(group)
        If targetSheetName = "" Then
            GoTo NextIteration
        End If
        
        ' �f�[�^���R���N�V�����ɒǉ�
        maxValue = wsSource.Range("J" & i).value
        sampleName = wsSource.Range("D" & i).value
        testPoint = wsSource.Range("N" & i).value
        dummyHead = wsSource.Range("P" & i).value
        anvilType = wsSource.Range("O" & i).value
        preProcess = wsSource.Range("M" & i).value
        tempArray = Array( _
            idParts(0), _
            targetSheetName, _
            maxValue, _
            sampleName, _
            testPoint, _
            dummyHead, _
            anvilType, _
            preProcess _
        )
        data.Add tempArray
        
NextIteration:
    Next i
    
    ' �f�[�^���e�V�[�g�ɓ]�L
    TransferDataToSheets data
    
    ' ���\�[�X�����
    Set wsSource = Nothing
    Set data = Nothing
End Sub

Function GetTargetSheetName(ByVal group As String) As String
'TransferDataBasedOnID�̃T�u�֐��B�O���[�v�Ɋ�Â��ă^�[�Q�b�g�V�[�g�����擾����
    Select Case group
        Case "�V"
            GetTargetSheetName = "Sub1"
        Case "�O"
            GetTargetSheetName = "Sub2"
        Case "��"
            GetTargetSheetName = "Sub3"
        Case Else
            GetTargetSheetName = "���|�[�g�O���t"
    End Select
End Function

Sub TransferDataToSheets(ByVal data As Collection)
' �f�[�^���e�V�[�g�ɓ]�L����TransferDataBasedOnID�̃T�u�v���V�[�W��
    Dim wsDest As Worksheet
    Dim dataItem As Variant
    Dim nextRow As Long
    
    For Each dataItem In data
        ' �ϐ��̊��蓖��
        Dim groupName As String
        Dim targetSheetName As String
        Dim preProcess As String
        Dim topGap As String
        Dim testPoint As String
        Dim maxValue As Double, duration49kN As Double, duration73kN As Double
        Dim sampleName As String
        
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        sampleName = dataItem(3)
        testPoint = dataItem(4)
        dummyHead = dataItem(5)
        anvilType = dataItem(6)
        preProcess = dataItem(7)
    
        ' �ړI�̃V�[�g���擾�܂��͍쐬
        Set wsDest = GetOrCreateSheet(targetSheetName)
        
        ' �w�b�_�[�s��ݒ�i14�s�ځj
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "�ő�l"
            wsDest.Range("C14").value = "�X��No."
            wsDest.Range("D14").value = "�����ӏ�"
            wsDest.Range("E14").value = "�l���͌^"
            wsDest.Range("F14").value = "�A���r��"
            wsDest.Range("G14").value = "�O����"
        End If
        
        ' ���̋�s���擾���f�[�^��]�L
        nextRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = maxValue
        wsDest.Range("C" & nextRow).value = sampleName
        wsDest.Range("D" & nextRow).value = testPoint
        wsDest.Range("E" & nextRow).value = dummyHead
        wsDest.Range("F" & nextRow).value = anvilType
        wsDest.Range("G" & nextRow).value = preProcess
    Next dataItem
End Sub

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
' TransferDataToSheets�̃T�u�֐��B�w�肳�ꂽ�V�[�g���擾�܂��͍쐬
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
    GetOrCreateSheet.Visible = xlSheetVisible
    On Error GoTo 0
End Function


' "���|�[�g�O���t"�V�[�g��"�e���v���[�g"�V�[�g����s���R�s�[���ăO���[�v���̂ݑ}��
Sub ProcessImpactSheets()
    ' �ϐ��̐錾
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim startRow As Long ' �T���J�n�s
    Dim currentGroup As String ' ���݂̃O���[�v�̒l
    Dim previousGroup As String ' �O�̃O���[�v�̒l
    Dim groupCount As Long ' �O���[�v�̌����J�E���g����ϐ�
    Dim groupSize As Long ' �e�O���[�v���̃f�[�^��
    Dim insertRowOffset As Long ' �}���s�̃I�t�Z�b�g
    Dim lastRow As Long ' A��̍ŏI�s�ԍ�
    Dim rowsToInsert As Collection ' �}���s�ԍ��̃R���N�V����
    Dim groupIndices As Collection ' �O���[�v�ԍ��̃R���N�V����
    Dim groupIndex As Long ' �O���[�v�ԍ�
    Dim insertRows As Long
    Dim row As Variant ' Collection�̃��[�v�p�ϐ�
    Dim index As Variant ' Collection�̃O���[�v�ԍ��p�ϐ�
    Dim groupCounter As Long ' �O���[�v�J�E���^�[�i�}���p�j

    ' "��������"�V�[�g��ϐ��Ɋi�[
    Set wsResult = Sheets("Bicycle_�e���v���[�g")

    ' �V�[�g����"���|�[�g�O���t"���܂ރV�[�g�����[�v����
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then

            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate

            ' A���"Group"����n�܂�s��T���J�n�s�Ƃ���
            startRow = Application.WorksheetFunction.Match("Group", ws.Range("A:A"), 0)

            ' startRow�̒l���m�F
            Debug.Print "startRow (Group�̍s): " & startRow

            ' A��̍ŏI�s���擾
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

            ' �T���͈͂��m�F
            Debug.Print "�T���͈�: A" & startRow & ":A" & lastRow

            ' A���"Group"�ȉ��̃f�[�^�������ɕ��בւ���
            With ws.Sort
                .SortFields.Clear
                .SortFields.Add key:=ws.Range("A" & startRow + 1 & ":A" & lastRow), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SetRange ws.Range("A" & startRow & ":Z" & lastRow)
                .Header = xlYes ' �w�b�_�[�����邱�Ƃ��w��
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With

            ' �}���\��̍s�ԍ���ێ�����R���N�V����
            Set rowsToInsert = New Collection
            Set groupIndices = New Collection

            ' �f�[�^��1�s���m�F���A�O���[�v�𔻕�
            previousGroup = "" ' �ŏ��̃��[�v�ł͑O�̃O���[�v�͑��݂��Ȃ����ߋ󕶎����ݒ�
            groupCount = 0 ' �O���[�v�̃J�E���g��������
            groupIndex = 1 ' �O���[�v�ԍ��̏�����
            groupSize = 0 ' �����O���[�v�̃T�C�Y��������

            For i = startRow + 1 To lastRow
                ' Trim�ŃX�y�[�X���������Ēl���擾
                currentGroup = Trim(ws.Cells(i, "A").value)

                ' �󔒃Z���𖳎�����
                If currentGroup <> "" Then
                    ' A��̒l���̂܂܂��o��
                    Debug.Print "�s�ԍ�: " & i & ", A��̒l: '" & currentGroup & "'"

                    ' �����O���[�v���ǂ����𔻒�
                    If currentGroup = previousGroup Then
                        groupSize = groupSize + 1 ' �����O���[�v���Ȃ̂ŃJ�E���g�𑝂₷
                    Else
                        ' �O�̃O���[�v�̏����������������_�ŁA�O���[�v�̃T�C�Y���m�F
                        If groupSize > 0 Then
                            ' �O���[�v��4�𒴂��Ă���ꍇ�̓G���[���o���Ė���
                            If groupSize > 4 Then
                                Debug.Print "�G���[: �O���[�v" & previousGroup & "��4�𒴂��Ă��܂��B�����𖳎����܂��B"
                            Else
                                ' �O���[�v�̃T�C�Y��4�ȉ��̏ꍇ�͑}���Ώ�
                                rowsToInsert.Add i - groupSize - 1 ' �O���[�v�̊J�n�s��ۑ�
                                groupIndices.Add groupIndex ' �O���[�v�ԍ���ۑ�
                            End If
                        End If
                        
                        ' �V�����O���[�v�Ɉڂ邽�߁A�J�E���g�����Z�b�g
                        groupSize = 1
                        groupIndex = groupIndex + 1 ' �V�����O���[�v�ԍ��ɐi��
                    End If

                    ' ���̃O���[�v����̂��߂Ɍ��݂̃O���[�v��ۑ�
                    previousGroup = currentGroup
                End If
            Next i

            ' �Ō�̃O���[�v���m�F
            If groupSize > 0 Then
                If groupSize > 4 Then
                    Debug.Print "�G���[: �O���[�v" & previousGroup & "��4�𒴂��Ă��܂��B�����𖳎����܂��B"
                Else
                    rowsToInsert.Add lastRow - groupSize + 1 ' �Ō�̃O���[�v�̊J�n�s��ۑ�
                    groupIndices.Add groupIndex ' �Ō�̃O���[�v�ԍ���ۑ�
                End If
            End If

            ' �s�}���������܂Ƃ߂Ď��s
            insertRowOffset = 0
            groupCounter = 1 ' �O���[�v�J�E���^�[��������
            For Each row In rowsToInsert
                ' �O���[�v�ԍ����擾
                index = groupIndices.item(groupCounter)
                
                insertRows = wsResult.Range("A2:A7").Rows.Count
                ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                With wsResult
                    .Range(.Cells(2, "A"), .Cells(7, "G")).Copy
                End With
                ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                ' index �� -1 ���āA������ groupIndex ��ݒ�
                ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & (index - 1)

                ' �}���s�̃I�t�Z�b�g���X�V
                insertRowOffset = insertRowOffset + insertRows
                groupCounter = groupCounter + 1 ' �O���[�v�J�E���^�[���X�V
            Next row

            ' ���[�v�̍Ō�ɃJ�E���^�����Z�b�g
            groupCount = 0
            insertRows = 0
            insertRowOffset = 0
        End If
    Next ws
End Sub
' "���|�[�g�O���t"�V�[�g�̗�/�s�̃T�C�Y�𐮂���
Sub SetCellDimensions()
    ' ProcessImpactSheets�̃T�u���[�`���B�V�[�g����"Impact"���܂ރV�[�g�����[�v����
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then
            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate

            ' I���"Insert" + �����������Ă���s�����[�v����
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
            Dim i As Long
            Dim currentInsertGroup As String
            Dim groupStartRow As Long

            For i = 2 To lastRow ' 2�s�ڂ���J�n
                ' I��̒l�� "Insert" �Ŏn�܂�A���̌�ɐ����������ꍇ
                If ws.Cells(i, "I").value Like "Insert[0-9]*" Then
                    ' �V�����O���[�v�̏ꍇ
                    If ws.Cells(i, "I").value <> currentInsertGroup Then
                        ' �O�̃O���[�v�̃Z�������E����ݒ� (�ŏ��̃O���[�v�ȊO)
                        If currentInsertGroup <> "" Then
                            SetGroupCellDimensions ws, groupStartRow, i - 1
                        End If

                        ' �V�����O���[�v�̊J�n�s���L�^
                        currentInsertGroup = ws.Cells(i, "I").value
                        groupStartRow = i
                    End If
                End If
            Next i

            ' �Ō�̃O���[�v�̃Z�������E����ݒ�
            If currentInsertGroup <> "" Then
                SetGroupCellDimensions ws, groupStartRow, lastRow
            End If
        End If
    Next ws

End Sub

Sub SetGroupCellDimensions(ws As Worksheet, startRow As Long, endRow As Long)
    ' SetCellDimensions�̃T�u���[�`���B�O���[�v�̃Z�������E����ݒ肷��
    ' A�񂩂�G��̕��Ɗe�s�̍������w�肳�ꂽ�����ɍ��킹�Đݒ�
    Debug.Print "startRow: " & startRow
    ' A��̕���񕝒P�ʂŎw��
    ws.Columns(1).ColumnWidth = 2.3

    ' B���E��̕���񕝒P�ʂŎw��
    ws.Columns(2).ColumnWidth = 11.8
    ws.Columns(5).ColumnWidth = 11.8

    ' C���F��̕���񕝒P�ʂŎw��
    ws.Columns(3).ColumnWidth = 11
    ws.Columns(6).ColumnWidth = 11

    ' D���G��̕���񕝒P�ʂŎw��
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(7).ColumnWidth = 16

    ' �s�̍������s�N�Z�����Z�̃|�C���g�Őݒ�
    Dim i As Long
    For i = startRow To endRow
        Select Case (i - startRow + 1) Mod 6
            Case 1, 2, 4, 5
                ws.Rows(i).RowHeight = 18
            Case 3, 6
                ws.Rows(i).RowHeight = 127.8
            Case 0
                ws.Rows(i).RowHeight = 127.8
        End Select
    Next i
End Sub
' "���|�[�g�O���t"�V�[�g�Ƀw�b�_�[��ǉ�����B(�r��)
Sub AddHeaderToReportSheets()
    Dim ws As Worksheet
    Dim lastCol As Long

    ' ���O��"���|�[�g�O���t"���܂܂��V�[�g������
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then
            ' 1�s�ڂɍs��}��
            ws.Rows(1).Insert Shift:=xlDown

            ' �ŏI����擾
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            
            ' �^�񒆂̗���v�Z���āA�_�~�[�e�L�X�g��}��
            ws.Cells(1, Application.RoundUp(lastCol / 2, 0)).value = "�_�~�[�e�L�X�g"
        End If
    Next ws
End Sub






