Attribute VB_Name = "Utlities"


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




Sub PrintImpactSheet()
    Dim ws As Worksheet
    
    ' ����1: ����̃V�[�g�����
    Dim sheetNames1 As Variant
    sheetNames1 = Array("Impact_Top", "Impact_Front", "Impact_Back")
    
    For Each ws In ThisWorkbook.Sheets
        If foundSheetName(ws.Name, sheetNames1) Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Sub PrintSideImpactSheet()
    Dim ws As Worksheet
    
    ' ����2: "Impact_Side"�𖼑O�Ɋ܂ރV�[�g�����
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "Impact_Side") > 0 Then
            ws.PrintOut From:=1, To:=1
        End If
    Next ws
End Sub

Function foundSheetName(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            foundSheetName = True
            Exit Function
        End If
    Next i
    foundSheetName = False
End Function
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


Private Sub DeleteInsertedRows()
    ' �A�N�e�B�u�V�[�g��I��� "Insert" �ƈ󂪂��Ă���s���폜����
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    
    ' �A�N�e�B�u�V�[�g���擾
    Set ws = ActiveSheet
    
    ' I��̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    
    ' �Ō�̍s����1�s����Ɍ������č폜���m�F
    For currentRow = lastRow To 1 Step -1
        If Left(ws.Cells(currentRow, "I").value, 6) = "Insert" Then
            ws.Rows(currentRow).Delete
        End If
    Next currentRow
End Sub

