Attribute VB_Name = "AddHeaderToReportSheet"
Option Explicit

Public Sub InsertHeaderRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCell As Range
    Dim previousValue As String
    Dim currentValue As String
    Dim insertRow As Long
    Dim newValue As String
    
    ' "���|�[�g�O���t"�V�[�g��ݒ�
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    
    previousValue = ""
    insertRow = 0
    
    ' I����ŏI�s�����Ɍ������ĒT��
    For i = lastRow To 1 Step -1
        Set currentCell = ws.Cells(i, "I")
        
        ' �Z������łȂ��AInsert�Ŏn�܂�ꍇ������
        If Not IsEmpty(currentCell) Then
            If Left(currentCell.value, 6) = "Insert" Then
                ' ���l�������Ă��邩�m�F�i��FInsert1, Insert2�Ȃǁj
                If IsNumeric(Mid(currentCell.value, 7)) Then
                    currentValue = currentCell.value
                    
                    ' �O�̒l�ƈقȂ�ꍇ�i�V�����O���[�v�̊J�n�j
                    If currentValue <> previousValue And previousValue <> "" Then
                        ' ���������𒊏o���ĐV�����l���쐬
                        newValue = "NewColumn" & Mid(previousValue, 7)
                        
                        ' ���݂̍s�̏�ɐV�����s��}��
                        ws.Rows(insertRow).Insert Shift:=xlDown
                        ' �}�������s��NewColumn+Num��ݒ�
                        ws.Cells(insertRow, "I").value = newValue
                        Debug.Print "Inserted row at " & insertRow & " with value " & newValue
                    End If
                    
                    ' ���݂̒l���L�^
                    previousValue = currentValue
                    ' ���̑}���ʒu�����݂̍s�ɐݒ�
                    insertRow = i
                End If
            End If
        End If
    Next i
    
    ' �ŏ��̃O���[�v�̂��߂̍s�}��
    If insertRow > 0 Then
        ' �Ō�̃O���[�v�̐������g�p���ĐV�����l���쐬
        newValue = "NewColumn" & Mid(previousValue, 7)
        
        ws.Rows(insertRow).Insert Shift:=xlDown
        ws.Cells(insertRow, "I").value = newValue
        Debug.Print "Inserted row at first group " & insertRow & " with value " & newValue
    End If
    
    MsgBox "�w�b�_�[�s�̑}�����������܂����B", vbInformation
End Sub

Public Sub FormatNewColumnRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentCell As Range
    Dim aCell As Range
    Dim hasNewColumn As Boolean
    Dim excludeList As Variant
    
    ' ���O���镶����̃��X�g���`
    excludeList = Array("SampleText")
    
    ' "���|�[�g�O���t"�V�[�g��ݒ�
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    
    ' �ŏI�s���擾�iI���A��̑傫�������g�p�j
    lastRow = WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "I").End(xlUp).row, _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    
    ' I���"NewColumn"�����݂��邩�`�F�b�N
    hasNewColumn = False
    For i = 1 To lastRow
        If Not IsEmpty(ws.Cells(i, "I")) Then
            If Left(ws.Cells(i, "I").value, 9) = "NewColumn" Then
                hasNewColumn = True
                Exit For
            End If
        End If
    Next i
    
    ' "NewColumn"��������Ȃ��ꍇ�͏����𒆎~
    If Not hasNewColumn Then
        MsgBox "I���'NewColumn'���܂ޒl��������܂���B" & vbCrLf & _
               "�����𒆎~���܂��B", vbExclamation
        Exit Sub
    End If
    
    ' ���C������
    For i = 1 To lastRow
        Set currentCell = ws.Cells(i, "I")
        Set aCell = ws.Cells(i, "A")
        
        ' A��̒l���m�F�i3�����ȏォ���O���X�g�Ɋ܂܂�Ȃ��ꍇ�j
        If Not IsEmpty(aCell) Then
            If Len(aCell.value) >= 3 Then
                ' ���O���X�g�Ɋ܂܂�Ă��Ȃ����`�F�b�N
                Dim isExcluded As Boolean
                Dim excludeWord As Variant
                isExcluded = False
                
                For Each excludeWord In excludeList
                    If aCell.value = excludeWord Then
                        isExcluded = True
                        Exit For
                    End If
                Next excludeWord
                
                ' ���O���X�g�Ɋ܂܂�Ă��Ȃ��ꍇ�̂ݏ���
                If Not isExcluded Then
                    ' �O�̍s��B-G��ɒl��]�L
                    If i > 1 Then  ' 1�s�ڂ�艺�̏ꍇ�̂�
                        With ws.Range(ws.Cells(i - 1, "B"), ws.Cells(i - 1, "G"))
                            .Merge
                            .value = aCell.value
                            .HorizontalAlignment = xlLeft
                        End With
                    End If
                End If
            End If
        End If
        
        ' I���NewColumn�̏���
        If Not IsEmpty(currentCell) Then
            If Left(currentCell.value, 9) = "NewColumn" Then
                ' �s�̍�����ݒ�
                ws.Rows(i).RowHeight = 18
                
                ' B�񂩂�G�������
                With ws.Range(ws.Cells(i, "B"), ws.Cells(i, "G"))
                    .Merge
                    .HorizontalAlignment = xlLeft
                End With
                
                ' �w�i�F�ƃt�H���g�F��ݒ�
                With ws.Range(ws.Cells(i, "A"), ws.Cells(i, "G"))
                    .Interior.Color = RGB(48, 84, 150)
                    .Font.Color = RGB(242, 242, 242)
                End With
                
                Debug.Print "Formatted NewColumn row at " & i
            End If
        End If
    Next i
    
    MsgBox "�t�H�[�}�b�g���������܂����B", vbInformation
End Sub

Public Sub SetReportHeader()
    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim headerRange As Range
    Dim headerExists As Boolean
    
    ' �V�[�g�̑��݊m�F
    If Not WorksheetExists("���|�[�g�O���t") Or Not WorksheetExists("���|�[�g�{��") Then
        MsgBox "�K�v�ȃV�[�g��������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �V�[�g�̐ݒ�
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    Set wsSource = ThisWorkbook.Worksheets("���|�[�g�{��")
    
    ' HeaderColumn�̑��݊m�F
    headerExists = False
    If Not IsEmpty(ws.Range("I1")) Then
        If ws.Range("I1").value = "HeaderColumn" Or ws.Range("I2").value = "HeaderColumn" Then
            headerExists = True
        End If
    End If
    
    ' �w�b�_�[�����ɑ��݂���ꍇ�͏������I��
    If headerExists Then
        Debug.Print "�w�b�_�[�͊��ɑ��݂��܂��B"
        Exit Sub
    End If
    
    ' �����̃w�b�_�[�s��}��
    ws.Rows("1:2").Insert Shift:=xlDown
    
    Application.ScreenUpdating = False
    
    ' A1:B2 �̌����ƃR�s�[
    With ws.Range("A1:B2")
        .Merge
        .value = wsSource.Range("A1").value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    ' C1:D2 �̌����ƃR�s�[
    With ws.Range("C1:E2")
        .Merge
        .value = wsSource.Range("C1").value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' F��̐ݒ�
    With ws.Range("F1:F2")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    ws.Range("F1").value = wsSource.Range("G1").value
    ws.Range("F2").value = wsSource.Range("G2").value
    
    ' G��̐ݒ�
    With ws.Range("G1:G2")
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("G1").value = wsSource.Range("H1").value
    ws.Range("G2").value = wsSource.Range("H2").value
    
    ' HeaderColumn�̐ݒ�
    ws.Range("I1:I2").value = "HeaderColumn"
    
    ' �S�̂̏����ݒ�
    With ws.Range("A1:G2")
        .Font.Name = "���S�V�b�N"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' �s�̍�����ݒ�
    ws.Rows(1).RowHeight = 20
    ws.Rows(2).RowHeight = 20
    
    Application.ScreenUpdating = True
    
    Debug.Print "���|�[�g�w�b�_�[��ݒ肵�܂����B"
End Sub
Public Sub ClearColoredCellsInColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim clearedCount As Long
    
    ' �V�[�g�̑��݊m�F
    If Not WorksheetExists("���|�[�g�O���t") Then
        MsgBox "���|�[�g�O���t�V�[�g��������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �V�[�g�̐ݒ�
    Set ws = ThisWorkbook.Worksheets("���|�[�g�O���t")
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Application.ScreenUpdating = False
    
    ' �J�E���^�[�̏�����
    clearedCount = 0
    
    ' A��̊e�Z�����`�F�b�N
    For Each cell In ws.Range("A1:A" & lastRow)
        ' �Z���ɔw�i�F������ꍇ
        If cell.Interior.colorIndex <> xlNone Then
            ' �Z���̓��e���N���A
            cell.ClearContents
            clearedCount = clearedCount + 1
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    ' ���ʂ�\��
    If clearedCount > 0 Then
        Debug.Print clearedCount & "�̃Z���̓��e���������܂����B"
    Else
        Debug.Print "�w�i�F�̂��Ă���Z���͌�����܂���ł����B"
        MsgBox "�w�i�F�̂��Ă���Z���͌�����܂���ł����B", vbInformation
    End If
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
