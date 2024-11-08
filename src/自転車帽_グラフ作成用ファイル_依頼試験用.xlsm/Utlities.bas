Attribute VB_Name = "Utlities"
' ���|�[�g�O���t�̈���͈͂�ݒ肷��
Sub SetPrintAreaForGroups()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim groupCount As Integer
    Dim groupRows() As Long
    Dim i As Long
    Dim printStart As Long
    Dim printEnd As Long
    Dim j As Integer
    Dim fileName As String
    Dim pdfFolder As String
    Dim currentGroup As String
    Dim previousGroup As String

    ' �ۑ�����PDF�̃t�H���_�p�X��ݒ�
    pdfFolder = ThisWorkbook.Path & "\PDFs\"
    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder ' �t�H���_�����݂��Ȃ��ꍇ�A�쐬
    End If

    ' �V�[�g�����[�v����"���|�[�g�O���t"���܂ރV�[�g������
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then
            ' �V�[�g�̍ŏI�s���擾
            lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row

            ' �O���[�v�̍s��ێ����邽�߂̔z��𓮓I�Ɋm��
            ReDim groupRows(0)

            ' "Insert + ���l"���܂܂��s�������Ĕz��ɕۑ�
            groupCount = 0 ' ������
            previousGroup = ""
            For i = 1 To lastRow
                currentGroup = ws.Cells(i, "I").value
                If InStr(currentGroup, "Insert") > 0 And currentGroup <> previousGroup Then
                    groupCount = groupCount + 1
                    ReDim Preserve groupRows(groupCount)
                    groupRows(groupCount - 1) = i
                    previousGroup = currentGroup ' �O���[�v���ς�������̂ݍX�V
                End If
            Next i

            ' �O���[�v���ɉ����Ĉ���͈͂�ݒ�
            If groupCount = 0 Then
                MsgBox "����͈͂ɊY������O���[�v��������܂���ł����B"
                Exit For ' �V�[�g���Ȃ��ꍇ�͎��ɐi��
            End If

            ' �O���[�v��2������͈͂ɐݒ�
            For j = 0 To groupCount - 1 Step 2
                printStart = groupRows(j) ' �O���[�v�̊J�n�s��ݒ�

                If j + 2 < groupCount Then
                    ' ���̎��̃O���[�v�̊J�n�s��1�s�O���I���s�Ƃ���
                    printEnd = groupRows(j + 2) - 1
                Else
                    ' �Ō�̃O���[�v���P�Ƃ̏ꍇ�A�Ō�̍s�܂Ŋ܂߂�
                    printEnd = lastRow
                End If

                ' ����͈͂�ݒ�
                ws.PageSetup.PrintArea = ws.Range("A" & printStart & ":G" & printEnd).Address

                ' �O���[�v�ƈ���͈͂��f�o�b�O�E�C���h�E�ɕ\��
                If j + 1 < groupCount Then
                    Debug.Print "�V�[�g: " & ws.Name & ", �O���[�v: " & ws.Cells(groupRows(j), "I").value & " - " & ws.Cells(groupRows(j + 1), "I").value & ", ����͈�: " & ws.PageSetup.PrintArea
                Else
                    Debug.Print "�V�[�g: " & ws.Name & ", �O���[�v: " & ws.Cells(groupRows(j), "I").value & ", ����͈�: " & ws.PageSetup.PrintArea
                End If

                ' PDF�o�́i�R�����g�A�E�g���j
                 fileName = pdfFolder & ws.Name & "_Group_" & j + 1 & ".pdf"
                 ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard
            Next j
        End If
    Next ws
End Sub






' ���|�[�g�O���t�V�[�g�̓��e���폜����B
Sub DeleteContentFromReportGraphSheets()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' �V�[�g�����[�v���āA���O�Ɂu���|�[�g�O���t�v���܂܂�Ă���V�[�g������
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "���|�[�g�O���t") > 0 Then
            ' A��̍ŏI�s���擾
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

            ' �ŏI�s��1�ȏ�ł���΁A�s���폜����
            If lastRow > 0 Then
                ws.Rows("1:" & lastRow).Delete
            End If
        End If
    Next ws
End Sub


