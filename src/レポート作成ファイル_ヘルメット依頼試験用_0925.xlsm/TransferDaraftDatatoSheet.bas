Attribute VB_Name = "TransferDaraftDatatoSheet"
' "Impact_Top"�Ȃǂ̃V�[�g���쐬���A�l��]�L����v���V�[�W��
Sub TransferDataBasedOnID()
    Dim wsSource As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim preProcess As String
    Dim topGap As String
    Dim testPoint As String
    Dim sampleName As String
    Dim MaxValue As Double, duration49kN As Double, duration73kN As Double
    Dim tempArray As Variant
    Dim data As Collection
    
    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row
    
    ' �e�s�����[�v����
    For i = 1 To lastRow
        ' ID�𕪊����A�K�v�ȏ����擾
        idParts = Split(wsSource.Cells(i, "B").value, "-")
        If UBound(idParts) >= 3 Then
            group = idParts(2)
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
        MaxValue = wsSource.Range("H" & i).value
        duration49kN = wsSource.Range("J" & i).value
        duration73kN = wsSource.Range("K" & i).value
        preProcess = wsSource.Range("L" & i).value
        topGap = wsSource.Range("N" & i).value
        testPoint = wsSource.Range("E" & i).value
        sampleName = wsSource.Range("D" & i).value
        tempArray = Array( _
            idParts(0), _
            targetSheetName, _
            MaxValue, _
            duration49kN, _
            duration73kN, _
            preProcess, _
            topGap, _
            testPoint, _
            sampleName _
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
            GetTargetSheetName = "Impact_Top"
        Case "�O"
            GetTargetSheetName = "Impact_Front"
        Case "��"
            GetTargetSheetName = "Impact_Back"
        Case Else
            GetTargetSheetName = ""
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
        Dim MaxValue As Double, duration49kN As Double, duration73kN As Double
        Dim sampleName As String
        
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        MaxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)
        preProcess = dataItem(5)
        topGap = dataItem(6)
        testPoint = dataItem(7)
        sampleName = dataItem(8)
    
        ' �ړI�̃V�[�g���擾�܂��͍쐬
        Set wsDest = GetOrCreateSheet(targetSheetName)
        
        ' �w�b�_�[�s��ݒ�i14�s�ځj
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "�X��No."
            wsDest.Range("C14").value = "�O����"
            wsDest.Range("D14").value = "�����ʒu"
            wsDest.Range("E14").value = "MAX"
            wsDest.Range("F14").value = "�V��������"
            wsDest.Range("G14").value = "4.9kN"
            wsDest.Range("H14").value = "7.3kN"
        End If
        
        ' ���̋�s���擾���f�[�^��]�L
        nextRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = sampleName
        wsDest.Range("C" & nextRow).value = preProcess
        wsDest.Range("D" & nextRow).value = testPoint
        wsDest.Range("E" & nextRow).value = MaxValue
        wsDest.Range("F" & nextRow).value = topGap
        wsDest.Range("G" & nextRow).value = duration49kN
        wsDest.Range("H" & nextRow).value = duration73kN
    Next dataItem
End Sub

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
' TransferDataToSheets�̃T�u�֐��B�w�肳�ꂽ�V�[�g���擾�܂��͍쐬
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetOrCreateSheet.Name = sheetName
    End If
    GetOrCreateSheet.Visible = xlSheetVisible
    On Error GoTo 0
End Function

' "��������"�V�[�g����e���v���[�g���R�s�[���Ă���v���V�[�W��
Sub ProcessImpactSheets()
    ' �ϐ��̐錾
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim count As Integer
    Dim insertRows As Long
    Dim startRow As Long ' �T���J�n�s
    Dim currentGroup As Variant ' ���݂̃O���[�v�̒l
    Dim previousGroup As Variant ' �O�̃O���[�v�̒l
    Dim groupCount As Long ' �O���[�v�̌����J�E���g����ϐ�
    Dim groupValues As Object ' �e�O���[�v�̒l�ƃJ�E���g���i�[����Dictionary
    Dim insertRowOffset As Long ' �}���s�̃I�t�Z�b�g

    ' "��������"�V�[�g��ϐ��Ɋi�[
    Set wsResult = Sheets("��������")

    ' �V�[�g����"Impact"���܂ރV�[�g�����[�v����
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
        
            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate

            ' A���"Group"����n�܂�s��T���J�n�s�Ƃ���
            startRow = Application.WorksheetFunction.Match("Group", ws.Range("A:A"), 0)

            ' A���"Group"����n�܂�s����ŏI�s�܂Ń��[�v����
            count = 0 ' ���[�v�̍ŏ��� count �����Z�b�g
            previousGroup = "" ' �ŏ��̃��[�v�ł͑O�̃O���[�v�͑��݂��Ȃ����ߋ󕶎����ݒ�
            groupCount = 0 ' �O���[�v�̌���������
            Set groupValues = CreateObject("Scripting.Dictionary") ' Dictionary�I�u�W�F�N�g���쐬
            insertRowOffset = 0
            insertCount = 0
            
            For i = startRow To ws.Cells(ws.Rows.count, "A").End(xlUp).row
                currentGroup = ws.Cells(i, "A").value

                ' ���݂̃O���[�v�ƑO�̃O���[�v�������ꍇ�A�J�E���g�𑝂₷
                If currentGroup = previousGroup Then
                    count = count + 1
                Else
                    ' ���݂̃O���[�v�ƑO�̃O���[�v���قȂ�ꍇ�A�J�E���g�����Z�b�g
                    ' �V�����O���[�v���J�n���ꂽ���߁A�O�̃O���[�v�̒l�ƃJ�E���g��Dictionary�Ɋi�[
                    If previousGroup <> "" Then
                        If Not groupValues.Exists(previousGroup) Then
                            groupValues.Add previousGroup, count
                        Else
                            groupValues(previousGroup) = groupValues(previousGroup) + count ' �l���X�V����
                        End If

                        If groupCount > 0 Then '2��ڈȍ~�̃O���[�v�̏ꍇ
                            insertRowOffset = insertRowOffset + insertRows
                        End If
                    End If

                    count = 1
                    groupCount = groupCount + 1 ' �V�����O���[�v���J�n���ꂽ���߁A�O���[�v�̌��𑝂₷

                End If

                ' �O�̃O���[�v���X�V
                previousGroup = currentGroup
            Next i

            ' �Ō�̃O���[�v�̒l�ƃJ�E���g��Dictionary�Ɋi�[
            If Not groupValues.Exists(previousGroup) Then
                groupValues.Add previousGroup, count
            Else
                groupValues(previousGroup) = groupValues(previousGroup) + count
            End If

'            ' �O���[�v�̐��Ɗe�O���[�v�̃J�E���g���o��
'            Debug.Print "�V�[�g: " & ws.Name
'            Debug.Print "�O���[�v�̐�: " & groupCount
            For Each key In groupValues.Keys
'                Debug.Print "�O���[�v " & key & ": " & groupValues(key) & " ��"
'                Debug.Print "�}���ʒu: : " & insertRowOffset

                ' �O���[�v�̃J�E���g�Ɋ�Â��đ}���ƃR�s�[���s��
                Select Case groupValues(key)
                    Case 3
                        insertRows = wsResult.Range("A3:A5").Rows.count
                        ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                        With wsResult
                            .Range(.Cells(3, "A"), .Cells(5, "G")).Copy
                        End With
                        ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                        ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & key

                    Case 2
                        insertRows = wsResult.Range("A7:A9").Rows.count
                        ws.Rows(2 + insertRowOffset).Resize(insertRows).Insert Shift:=xlDown
                        With wsResult
                            .Range(.Cells(7, "A"), .Cells(9, "G")).Copy
                        End With
                        ws.Range("A" & 2 + insertRowOffset).PasteSpecial xlPasteAll
                        ws.Range("I" & 2 + insertRowOffset).Resize(insertRows).value = "Insert" & key
                End Select
                
                ' �e�O���[�v�̏������ insertRowOffset ���X�V
                insertRowOffset = insertRowOffset + insertRows
            Next key

            ' ���[�v�̍Ō�Ɋe�J�E���^�����Z�b�g
            count = 0
            insertRows = 0
            insertRowOffset = 0
        End If
    Next ws
    Call SetCellDimensions
End Sub

Sub SetCellDimensions()
    ' ProcessImpactSheets�̃T�u���[�`���B�V�[�g����"Impact"���܂ރV�[�g�����[�v����
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' �V�[�g���A�N�e�B�u�ɂ���
            ws.Activate

            ' I���"Insert" + �����������Ă���s�����[�v����
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
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
    ' A2:G4�Z���͈͂��擾 (startRow��endRow�ɍ��킹�Ē���)
    Dim targetRange As Range
    Set targetRange = ws.Range("A" & startRow & ":G" & endRow)

    ' A��̕���ݒ�
    targetRange.Columns(1).ColumnWidth = 2.8 ' �s�N�Z�����|�C���g�ɕϊ�

    ' B�񂩂�G��̕���ݒ�
    targetRange.Columns(2).Resize(1, 6).ColumnWidth = 11.5 ' �s�N�Z�����|�C���g�ɕϊ�

    ' 1�s�ڂ�3�s�ڂ̍�����ݒ� (startRow�ɍ��킹�Ē���)
    targetRange.Rows(1).RowHeight = 18
    targetRange.Rows(3).RowHeight = 18

    ' 2�s�ڂ̍�����ݒ� (startRow�ɍ��킹�Ē���)
    targetRange.Rows(2).RowHeight = 161
End Sub





