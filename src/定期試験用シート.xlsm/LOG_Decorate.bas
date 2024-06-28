Attribute VB_Name = "LOG_Decorate"

Sub SaveChartsAsPNG()
    ' �O���t��PNG�ɕϊ����f�X�N�g�b�v�̃t�H���_�ɕۑ�����B
    ' ���[�N�V�[�g�̖��O��錾
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    ' Windows�̃f�X�N�g�b�v�̃p�X���擾
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    ' �����̓��t���擾���A�w��̃t�H���_�����쐬
    Dim folderName As String
    folderName = "Graph_" & Format(Date, "yyyy-mm-dd")

    ' �t�H���_�̃p�X���쐬
    Dim folderPath As String
    folderPath = desktopPath & "" & folderNamee

    ' �t�H���_�����݂��Ȃ��ꍇ�A�V���ɍ쐬
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    Dim ws As Worksheet
    Dim i As Integer
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
        
            ' �`���[�g�I�u�W�F�N�g��錾
            Dim chartObj As ChartObject

            ' �t�@�C������錾
            Dim FileName As String

            ' �`���[�g�I�u�W�F�N�g���ƂɃ��[�v
            For Each chartObj In ws.ChartObjects

                ' �O���t�̃^�C�g�����ꎞ�I�ɕۑ����A�O���t����͍폜
                FileName = chartObj.chart.ChartTitle.text
                chartObj.chart.HasTitle = False

                ' �t�@�C������ ".png" ��ǉ�
                FileName = FileName & ".png"

                ' �t�@�C���p�X��ݒ�i�t�H���_�̃p�X + �t�@�C�����j
                Dim filePath As String
                filePath = folderPath & "" & FileNamee

                ' �`���[�g�̌��݂̕��ƍ�����ۑ�
                Dim originalWidth As Double
                Dim originalHeight As Double
                originalWidth = chartObj.Width
                originalHeight = chartObj.Height

                ' �`���[�g�̕���ݒ肵�A�����̓A�X�y�N�g���ێ�
                Dim aspectRatio As Double
                aspectRatio = chartObj.Height / chartObj.Width
                chartObj.Width = 1000
                chartObj.Height = 1000 * aspectRatio

                ' �`���[�g��PNG�t�@�C���Ƃ��ăG�N�X�|�[�g
                chartObj.chart.Export FileName:=filePath, FilterName:="PNG"

                ' �`���[�g�̕��ƍ��������̑傫���ɖ߂�
                chartObj.Width = originalWidth
                chartObj.Height = originalHeight

                ' �O���t�̃^�C�g�������ɖ߂�
                chartObj.chart.HasTitle = True
                chartObj.chart.ChartTitle.text = FileName
            Next chartObj
        End If
        
        Set ws = Nothing
    Next i
End Sub

Sub ApplyColorToAllSheets()
    '�e���O�V�[�g�ɐF�������肷��
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    Dim ws As Worksheet
    Dim i As Integer
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Call ColorCells(ws)
            Set ws = Nothing
        End If
    Next i
End Sub

Sub ColorCells(ws As Worksheet)
    'ApplyColorToALlSHeets�̊֐�
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim cellColor As Long

    ' A��̍ŏI�s���擾���܂�
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' A���2�s�ڂ���ŏI�s�܂ł͈̔͂��`���܂�
    Set rng = ws.Range("A2:A" & lastRow)

    ' �͈͓��̊e�Z���ɂ��ă��[�v���܂�
    For Each cell In rng
        If InStr(cell.value, "HEL") > 0 Then
            ' "HEL"���Z���̒l�̈ꕔ�ł���ꍇ�AG���H��̃Z���̐F���I�����W�ɂ��܂�
            cellColor = RGB(255, 111, 56)
            ColorAndFont ws.Range("H" & cell.row & ":I" & cell.row), cellColor
        ElseIf InStr(cell.value, "BICYCLE") > 0 Then
            ' "BICYCLE"���Z���̒l�̈ꕔ�ł���ꍇ�AI��̃Z���̐F���u���[�ɂ��܂�
            cellColor = RGB(8, 92, 255)
            ColorAndFont ws.Range("I" & cell.row), cellColor
        ElseIf InStr(cell.value, "BASEBALL") > 0 Then
            ' "BASEBALL"���Z���̒l�̈ꕔ�ł���ꍇ�AK��̃Z���̐F���O���[�ɂ��܂�
            cellColor = RGB(218, 218, 218)
            ColorAndFont ws.Range("K" & cell.row), cellColor
        ElseIf InStr(cell.value, "FALLARR") > 0 Then
            ' "FALLARR"���Z���̒l�̈ꕔ�ł���ꍇ�AL�񂩂�N��̃Z���̐F��΂ɂ��܂�
            cellColor = RGB(22, 187, 98)
            ColorAndFont ws.Range("L" & cell.row & ":N" & cell.row), cellColor
        End If

        ' F��̃Z���̐F�����l�ɕύX���܂�
        ColorAndFont ws.Range("F" & cell.row), cellColor
    Next cell
End Sub

Sub ColorAndFont(rng As Range, color As Long)
    'ColorCells�̊֐�
    rng.Interior.color = color
    rng.Font.color = RGB(255, 255, 255)
    rng.Font.Bold = True
End Sub

Sub DataMidrationAndCSVSheetDelete()
    Call FillColumnsQtoS
    Call CustomSort_Helmet
End Sub
Sub FillColumnsQtoS()
    ' �����\�̍��ڂɕ֋X��̍��i��������B
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' "LOG_Helmet"�V�[�g���w��
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' �ォ��ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' S,T��� "���i" �����
        ws.Cells(i, "S").value = "���i"
        ws.Cells(i, "T").value = "���i"

        ' Q��� "�X�V" �����
        'ws.Cells(i, "S").Value = "�X�V"
    Next i

    ' �������̊J��
    Set ws = Nothing

End Sub

Sub CustomSort_Helmet()
    'B���V���A�O�����A�㓪���A���̑��̏��Ƀ\�[�g����B
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' �f�[�^�͈̔͂��w�肵�܂��B1�s�ڂ͖�������̂�2����n�܂�܂��B
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim rng As Range
    Set rng = ws.Range("B2:B" & lastRow)
    
    ' �V�������ǉ����āA�\�[�g�L�[��ݒ肵�܂��B
    ws.Columns("C").Insert
    Dim cell As Range
    For Each cell In rng
        If InStr(cell.value, "HEL_TOP") > 0 Then
            cell.Offset(0, 1).value = 10000 + CInt(Mid(cell.value, 1, 4)) ' HEL_TOP�̏ꍇ
        ElseIf InStr(cell.value, "HEL_FRONT") > 0 Then
            cell.Offset(0, 1).value = 20000 + CInt(Mid(cell.value, 1, 4)) ' HEL_FRONT�̏ꍇ
        ElseIf InStr(cell.value, "HEL_BACK") > 0 Then
            cell.Offset(0, 1).value = 30000 + CInt(Mid(cell.value, 1, 4)) ' HEL_BACK�̏ꍇ
        ElseIf InStr(cell.value, "HEL_ZENGO") > 0 Then
            cell.Offset(0, 1).value = 40000 + CInt(Mid(cell.value, 1, 4)) ' HEL_ZENGO�̏ꍇ
        End If
    Next cell
    
    ' �S�Ă̗�iA�񂩂�Ō�̗�܂Łj�Ń\�[�g���܂��B
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Sort Key1:=ws.Range("C2"), Order1:=xlAscending, Header:=xlNo
    
    ' �\�[�g�Ɏg�p��������폜���܂��B
    ws.Columns("C").Delete
End Sub



Sub GenerateSampleID()
    ' �����p��ID�𐶐�����B
    Dim ws As Worksheet
    Dim rng As Range
    Dim dic As Object
    Dim i As Long
    Dim key As String
    Dim prefix As String
    Dim idNum As Long
    Dim randChars As String

    ' "LOG_Helmet"���[�N�V�[�g���w�肷��
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' Scripting.Dictionary���쐬����
    Set dic = CreateObject("Scripting.Dictionary")

    ' �f�[�^�͈͂��w�肷��
    Set rng = ws.Range("D2:P" & ws.Cells(ws.Rows.Count, "D").End(xlUp).row)

    ' �ړ�����ݒ肷��
    prefix = "_Hel"

    For i = 1 To rng.Rows.Count
        ' D��AM��AN��AO��AL��(�O����)�̒l���������ăL�[���쐬����
        key = ws.Cells(i + 1, "D").value & "_" & ws.Cells(i + 1, "M").value & "_" & ws.Cells(i + 1, "N").value & "_" & ws.Cells(i + 1, "O").value & "_" & ws.Cells(i + 1, "L").value

        ' �L�[������dic�ɑ��݂���ꍇ�A������ID���g�p����B���݂��Ȃ��ꍇ�A�V����ID�𐶐�����
        If dic.Exists(key) Then
            ws.Cells(i + 1, "C").value = dic(key)
        Else
            idNum = idNum + 1
            ' �����_���ȃA���t�@�x�b�g2�����𐶐�����
            randChars = Chr(Int((90 - 65 + 1) * Rnd + 65)) & Chr(Int((90 - 65 + 1) * Rnd + 65))
            ' �����_���ȕ�����ǉ�����ID�𐶐�����
            dic.Add key, Format(idNum, "00000") & randChars & prefix & ws.Cells(i + 1, "D").value
            ws.Cells(i + 1, "C").value = dic(key)
        End If
    Next i
End Sub





