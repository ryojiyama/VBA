Attribute VB_Name = "BicycleCreate"
Public Sub Create_BicycleGraph()
    Call CreateGraphBicycle
    Call Bicycle_150G_DurationTime
    ' �J���Ă���u�b�N�̈�ԍ��̃V�[�g��I��
    ThisWorkbook.Sheets(1).Select

    ' A1�Z���ɃJ�[�\�����ړ�
    Range("A1").Select
End Sub


Function ColNumToLetter(colNum As Integer) As String
    ' CreateGraphBicycle�ɂĎg�p����֐�
    Dim d As Integer
    Dim m As Integer
    Dim name As String
    d = colNum
    name = ""
    While (d > 0)
        m = (d - 1) Mod 26
        name = Chr(65 + m) & name
        d = Int((d - m) / 26)
    Wend
    ColNumToLetter = name
End Function

Sub CreateGraphBicycle()

    ' ���[�N�V�[�g��錾
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Bicycle")
    ' �ŏI�s�ƍŏI�������
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

    Dim i As Long
    Dim maxVal As Double
    Dim colStart As String
    Dim colEnd As String

    ' P��(16�Ԗ�) + 100�񂩂�n�܂��
    colStart = ColNumToLetter(16 + 100)

    ' P��(16�Ԗ�) + 800�񂩂�n�܂��
    colEnd = ColNumToLetter(16 + 1200)
    
    ' �����̃`���[�g�̈ʒu
    Dim chartLeft As Long
    Dim chartTop As Long
    chartLeft = 250
    chartTop = 100
    

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To lastRow
        ' �ő�l�����߂�
        maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))

        ' K��ɍő�l��\��
        ws.Cells(i, "G").Value = maxVal

        ' �`���[�g���쐬
        Dim ChartObj As ChartObject
        Set ChartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=375, Top:=chartTop, Height:=225)
        Dim chart As chart
        Set chart = ChartObj.chart

        ' �܂���O���t��ݒ�
        chart.ChartType = xlLine

        ' �O���t�̃f�[�^�͈͂�ݒ�
        chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
        
        ' X���̃f�[�^�͈͂�ݒ�
        chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))

        ' �O���t�̃^�C�g����ݒ�
        chart.HasTitle = True
        chart.ChartTitle.Text = ws.Cells(i, "B").Value

        ' �n��̕\�����I�t
        chart.SetElement msoElementLegendNone

        ' ���̑�����ݒ�
        chart.SeriesCollection(1).Format.Line.Weight = 0.75

        ' Y���̐ݒ�
        Dim yAxis As Axis
        Set yAxis = chart.Axes(xlValue, xlPrimary)

        ' Y���̍ő�l��ݒ�
        If maxVal <= 295 Then
            yAxis.MaximumScale = 300
        Else
            yAxis.MaximumScale = Int(maxVal) + 1
        End If

        yAxis.MinimumScale = -100 ' Y���̍Œ�l��-10�ɐݒ肵�܂��B

        ' Y���� TickLabels ��ݒ�
        With yAxis.TickLabels
            .NumberFormatLocal = "0""G"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With


        ' X���̐ݒ�
        Dim xAxis As Axis
        Set xAxis = chart.Axes(xlCategory, xlPrimary)
        xAxis.TickLabelSpacing = 100
        xAxis.TickMarkSpacing = 50


        ' X���� TickLabels ��ݒ�
        With xAxis.TickLabels
            .NumberFormatLocal = "0""ms"""
            .Font.color = RGB(89, 89, 89)
            .Font.Size = 8
        End With
        
        ' �`���[�g�̈ʒu�����ɍX�V
        chartLeft = chartLeft + 10

    Next i

End Sub



Sub Bicycle_150G_DurationTime()
    '���]�ԖX�����̃f�[�^���������郁�C���̃T�u���[�`��
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("LOG_BICYCLE")
    lastRow = GetLastRow(ws, "V")
    
    '�e�s�̍ő�l��F�t�����A�ő�l�̎��Ԃ��L�^���܂�
    ColorAndRecordMaxVal ws, lastRow, 150
    
    '150G�ȏ���L�^�����͈͂�F�t�����A���͈̔͂̎��ԍ����L�^���܂�
    ColorAndRecordTimeDifference ws, lastRow, 150
    
    '��̃Z����"-"�Ŗ��߂܂�
    FillEmptyCells ws, GetLastRow(ws, "B")
End Sub

'Bicycle150GDuration_����̃J�����̍ŏI�s���擾����֐��ł��B
Function GetLastRow(ws As Worksheet, column As String) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).row
End Function

'Bicycle150GDuration_�e�s�̍ő�l�̃Z����F�t�����A�ő�l�̎��Ԃ��L�^����T�u���[�`���ł��B
Sub ColorAndRecordMaxVal(ws As Worksheet, lastRow As Long, threshold As Double)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range
    
    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "V"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        
        Dim MaxValue As Double
        MaxValue = Application.WorksheetFunction.Max(rng)
        ws.Cells(i, "H").Value = MaxValue
        
        For Each cell In rng
            If cell.Value = MaxValue Then
                cell.Interior.color = RGB(255, 111, 56)
                ws.Cells(i, "I").Value = ws.Cells(1, cell.column).Value ' �Ή����鎞�Ԃ�I��ɋL�^
                Exit For ' �ŏ��̍ő�l�����������烋�[�v�𔲂���
            End If
        Next cell
    Next i
End Sub

'Bicycle150GDuration_150G�ȏ���L�^�����͈͂�F�t�����A���͈̔͂̎��ԍ����L�^����T�u���[�`���ł��B
Sub ColorAndRecordTimeDifference(ws As Worksheet, lastRow As Long, threshold As Double)
    Dim rng As Range
    Dim i As Long
    Dim cell As Range
    Dim startRange150 As Long
    Dim endRange150 As Long
    Dim maxRange150 As Long
    Dim maxStart150 As Long
    Dim maxEnd150 As Long
    Dim rangeCollection150 As Collection
    
    For i = 2 To lastRow
        Set rng = ws.Range(ws.Cells(i, "V"), ws.Cells(i, ws.Columns.Count).End(xlToLeft))
        startRange150 = 0
        endRange150 = 0
        maxRange150 = 0
        Set rangeCollection150 = New Collection
        
        For Each cell In rng
            If cell.Value >= threshold Then
                If startRange150 = 0 Then startRange150 = cell.column
                endRange150 = cell.column
                cell.Interior.color = RGB(0, 138, 211)
            Else
                If startRange150 > 0 And endRange150 > 0 Then
                    rangeCollection150.Add Array(startRange150, endRange150)
                    If (endRange150 - startRange150 + 1) > maxRange150 Then
                        maxRange150 = endRange150 - startRange150 + 1
                        maxStart150 = startRange150
                        maxEnd150 = endRange150
                    End If
                    startRange150 = 0
                    endRange150 = 0
                End If
            End If
        Next cell
        
        If startRange150 > 0 And endRange150 > 0 Then
            rangeCollection150.Add Array(startRange150, endRange150)
            If (endRange150 - startRange150 + 1) > maxRange150 Then
                maxRange150 = endRange150 - startRange150 + 1
                maxStart150 = startRange150
                maxEnd150 = endRange150
            End If
        End If
        
        If maxRange150 > 0 Then
            Dim timeDifference150 As Double
            timeDifference150 = ws.Cells(1, maxEnd150).Value - ws.Cells(1, maxStart150).Value
            ws.Cells(i, "K").Value = timeDifference150
        Else
            ws.Cells(i, "K").Value = "-"
        End If
    Next i
End Sub

'Bicycle150GDuration_��̃Z����"-"�Ŗ��߂�T�u���[�`���ł��B
Sub FillEmptyCells(ws As Worksheet, lastRow As Long)
    Dim cellRng As Range
    
    For Each cellRng In ws.Range("F2:P" & lastRow)
        If IsEmpty(cellRng) Then
            cellRng.Value = "-"
        End If
    Next cellRng
End Sub


