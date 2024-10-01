Attribute VB_Name = "Utlities"
' Impact�V�[�g�ƃ��|�[�g�{�����̕\���폜����B
Sub CleanUpSheetsByName()
    Call DeleteImpactSheets
    Call DeleteInsertedRows
End Sub

' Impact�V�[�g�̃O���t�t���̃e�[�u�����폜����B
Sub DeleteInsertRows()

  ' �V�[�g����"Impact"���܂ރV�[�g�����[�v����
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
    If InStr(ws.Name, "Impact") > 0 Then
      ' �V�[�g���A�N�e�B�u�ɂ���
      ws.Activate

      ' I���"Insert" + �����������Ă���s���t���Ƀ��[�v����
      Dim lastRow As Long
      lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
      Dim i As Long
      For i = lastRow To 1 Step -1
        ' I��̒l�� "Insert" �Ŏn�܂�A���̌�ɐ����������ꍇ�ɍ폜
        If ws.Cells(i, "I").value Like "Insert[0-9]*" Then
          ' �s�S�̂��폜
          ws.Rows(i).Delete
        End If
      Next i
    End If
  Next ws

End Sub

Sub DeleteRowsAfterGroup_ImpactSheets()

    ' "Impact"���܂ރV�[�g�������[�v����
    Dim ws As Worksheet
    Dim groupRowNumber As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' �S���[�N�V�[�g�����[�v����
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g���� "Impact" ���܂܂��V�[�g��Ώۂɂ���
        If InStr(ws.Name, "Impact") > 0 Then
            ' A��� "Group" �Ə����Ă���s��������
            lastRow = ws.UsedRange.Rows.count
            groupRowNumber = 0
            For i = 1 To lastRow
                If ws.Cells(i, "A").value = "Group" Then
                    groupRowNumber = i
                    Exit For
                End If
            Next i
            
            ' "Group"�����������ꍇ
            If groupRowNumber > 0 Then
                ' �폜����͈͂�1�s�ȏ゠�邱�Ƃ��m�F
                If groupRowNumber + 1 <= lastRow Then
                    ws.Rows(groupRowNumber + 1 & ":" & lastRow).EntireRow.Delete
                End If
            Else
                ' "Group"��������Ȃ������ꍇ�̏��� (��: ���b�Z�[�W�{�b�N�X��\��)
                MsgBox "�V�[�g '" & ws.Name & "' �� 'Group' ��������܂���ł����B", vbExclamation
            End If
        End If
    Next ws

End Sub


Sub PrintedReportSheets()
    Call PrintImpactSheet
    Call PrintSideImpactSheet
End Sub

Sub PrintImpactSheet()
    Dim ws As Worksheet
    Dim sheetNames1 As Variant
    Dim sheetFound As Boolean
    Dim i As Long
    Dim sheetName As String

    ' ����1: ����̃V�[�g����� ("Impact_Top", "Impact_Front", "Impact_Back", "���|�[�g�{��"���܂�)
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back", "���|�[�g�{��")

    ' �e�V�[�g���������Ĉ��
    For i = LBound(sheetNames1) To UBound(sheetNames1)
        sheetName = sheetNames1(i)
        sheetFound = False
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = sheetName Then
                ws.PrintOut From:=1, To:=1
                sheetFound = True
                Exit For
            End If
        Next ws
        ' �V�[�g��������Ȃ��ꍇ�̓��b�Z�[�W��\��
        If Not sheetFound Then
            MsgBox "�V�[�g '" & sheetName & "' ��������܂���B", vbExclamation
        End If
    Next i
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    Dim sheetFound As Boolean

    ' ����2: "Impact_Side"�𖼑O�Ɋ܂ރV�[�g�����
    sheetFound = False
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "Impact_Side") > 0 Then
            ws.PrintOut From:=1, To:=1
            sheetFound = True
        End If
    Next ws
    
    ' "Impact_Side"�V�[�g��������Ȃ������ꍇ�̃��b�Z�[�W
    If Not sheetFound Then
        MsgBox "�V�[�g���� 'Impact_Side' ���܂ރV�[�g��������܂���B", vbExclamation
    End If

    ' ���������
    Set ws = Nothing
End Sub


' Impact���܂ރV�[�g���̒���
Sub DeleteRowsBelowHeader()
    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim sheetName As String

    ' ���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g����"Impact"���܂܂�Ă��邩�`�F�b�N
        If InStr(ws.Name, "Impact") > 0 Then
            ' �w�b�_�[�̉��̍s����ŏI�s�܂ł��폜
            ws.Rows("15:" & ws.Rows.count).Delete
        End If
    Next ws
End Sub


Sub PrintChartIDs()
    Dim ws As Worksheet
    Dim chtObj As ChartObject

    ' �e���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �e���[�N�V�[�g���̃`���[�g�I�u�W�F�N�g�����[�v
        For Each chtObj In ws.ChartObjects
            ' CreateChartID�֐����g�p���� "Chart ID" �𐶐�
            Dim chartID As String
            chartID = CreateChartID(chtObj.chart.ChartArea.TopLeftCell)

            ' �C�~�f�B�G�C�g�E�B���h�E�ɏo��
            Debug.Print "Chart Name: " & chtObj.Name & ", Chart ID: " & chartID
        Next chtObj
    Next ws
End Sub

' �A�N�e�B�u�V�[�g��I��� "Insert" �ƈ󂪂��Ă���s���폜����
Private Sub DeleteInsertedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' "���|�[�g�{��"�V�[�g���擾
    Set ws = ThisWorkbook.Sheets("���|�[�g�{��")
    
    ' I��̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    
    ' �Ō�̍s����1�s����Ɍ������č폜���m�F
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "I").value, 6) = "Insert" Then
            ws.Rows(currentRow).Delete
        End If
    Next currentRow
End Sub

' �V�[�g����"Impact"�Ƃ��Ă���V�[�g���폜����B
Sub DeleteImpactSheets()
    Dim ws As Worksheet
    Dim sheetNamesToDelete As Collection
    Dim sheetName As String
    Dim i As Long
    
    ' �폜�Ώۂ̃V�[�g�����ꎞ�I�ɕێ�����R���N�V�������쐬
    Set sheetNamesToDelete = New Collection
    
    ' ���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g����"Impact"���܂܂�Ă��邩�`�F�b�N
        If InStr(ws.Name, "Impact") > 0 Then
            ' �폜�Ώۂ̃V�[�g�����R���N�V�����ɒǉ�
            sheetNamesToDelete.Add ws.Name
        End If
    Next ws
    
    ' �R���N�V�������̃V�[�g���폜
    For i = sheetNamesToDelete.count To 1 Step -1
        ThisWorkbook.Sheets(sheetNamesToDelete(i)).Delete
    Next i
End Sub

