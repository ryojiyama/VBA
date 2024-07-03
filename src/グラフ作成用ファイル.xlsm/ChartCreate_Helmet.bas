Attribute VB_Name = "ChartCreate_Helmet"
Sub HelmetTestResultChartBuilder()
    '�O���t�쐬�ƃw�����b�g�������Ԃ̕\���A�F�t���Ȃ�
    Call CreateGraphHelmet
    Call InspectHelmetDurationTime
    Call Utlities.AdjustingDuplicateValues
End Sub

' ��̏I�������肷��֐�
Function GetColumnEnd(ByRef ws As Worksheet, ByVal rowNumber As Long) As String
    Dim lastCol As Long
    Dim col As Long
    Dim found As Boolean
    found = False

    ' ��̍Ōォ��J�n���Ēl��1.0�𒴂���Ō�̗�ԍ���T��
    For col = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If ws.Cells(rowNumber, col).value > 1# Then
            lastCol = col
            found = True
            Exit For
        End If
    Next col

    ' �l��1.0�𒴂���񂩂�100�����v�Z
    If found Then
        lastCol = lastCol + 100
        If lastCol > ws.Columns.Count Then lastCol = ws.Columns.Count ' �񐔂̍ő�l�𒴂��Ȃ��悤�ɒ���
    Else
        ' 1.0�𒴂���l��������Ȃ��ꍇ�́A�K���ȃf�t�H���g�l��ݒ肷�邩�A�G���[�������s��
        lastCol = 150
    End If

    ' ��ԍ������̃A�h���X���擾���A�s�ԍ����폜
    Dim fullAddress As String
    fullAddress = ws.Cells(1, lastCol).Address(False, False)  ' ��ΎQ�Ƃ������
    GetColumnEnd = Replace(fullAddress, "1", "")  ' �s�ԍ����폜
End Function



' �O���t�̃T�C�Y�����肷��֐�
Function GetChartSize(ByVal graphType As String) As Variant
    Dim size(1) As Long
    
    Select Case graphType
        Case "��������p"
            size(0) = 400  ' Width
            size(1) = 200  ' Height
        Case "�^���\�������p"
            size(0) = 300  ' Width
            size(1) = 350  ' Height
        Case Else
            size(0) = 400  ' Default Width
            size(1) = 250  ' Default Height
    End Select
    
    GetChartSize = size
End Function

' �O���t���쐬���郁�C���̃T�u�v���V�[�W��
Sub CreateGraphHelmet(userInput As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim chartLeft As Long
    Dim chartTop As Long
    Dim colStart As String
    Dim colEnd As String
    Dim chartSize As Variant

    colStart = "GY"  ' �J�n��������ݒ�
    chartTop = ws.Rows(lastRow + 1).Top + 10
    chartLeft = 250

    For i = 2 To lastRow
        colEnd = GetColumnEnd(ws, i)
        chartSize = GetChartSize(userInput)
        CreateIndividualChart ws, i, chartLeft, chartTop, colStart, colEnd, chartSize
        chartLeft = chartLeft + 10 ' ���̃O���t�̍��ʒu�𒲐�
    Next i

End Sub


' CreateGraphHelmet_�ʂ̃O���t��ݒ�E�ǉ�����T�u�v���V�[�W��
Sub CreateIndividualChart(ByRef ws As Worksheet, ByVal i As Long, ByRef chartLeft As Long, ByVal chartTop As Long, ByVal colStart As String, ByVal colEnd As String, ByVal chartSize As Variant)
    Dim maxVal As Double
    maxVal = Application.WorksheetFunction.Max(ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd)))
    ws.Cells(i, "H").value = maxVal

    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartSize(0), Top:=chartTop, Height:=chartSize(1))
    Dim chart As chart
    Set chart = chartObj.chart

    With chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))  ' X���͈̔͂�1�s�ڂ���ݒ�
        .SeriesCollection(1).Values = ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))  ' Y���̃f�[�^�͈͂�ݒ�
        .SeriesCollection(1).Name = "Data Series " & i
    End With

    ' ID���쐬���ăO���t�^�C�g���ɐݒ�
    Dim recordID As String
    recordID = CreateChartID(ws.Cells(i, "B"))
    Debug.Print "recordID:" & recordID
    chartObj.Name = recordID
    ConfigureChart chart, ws, i, colStart, colEnd, maxVal
End Sub

Function CreateChartID(cell As Range) As String
    Dim parts() As String
    Dim createID As String

    ' B��̒l����̏ꍇ��"00000"��Ԃ�
    If IsEmpty(cell) Or cell.value = "" Then
        createID = "00000"
    Else
        ' B��̒l��Split�֐��ŕ������APart(0) & Part(1)�̌`����ID���쐬
        parts = Split(cell.value, "-")
'        Debug.Print "Cell value: " & cell.value  ' �f�o�b�O: �Z���̒l���o��
'        Debug.Print "Parts count: " & UBound(parts) + 1  ' �f�o�b�O: �������ꂽ�����̐����o��
        If UBound(parts) >= 1 Then
            createID = parts(0) & "-" & parts(1) & "-" & parts(2)
        Else
            createID = cell.value
        End If
    End If
    CreateChartID = createID
End Function

' CreateGraphHelmet_�O���t�̏����ݒ������T�u�v���V�[�W��
Sub ConfigureChart(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal colStart As String, ByVal colEnd As String, ByVal maxVal As Double)
    '���̃v���V�[�W����X����Y���̖ڐ�����ǉ�����B�������Ȃ��Ƃ��܂������Ȃ��B
    chart.ChartType = xlLine
    chart.SetSourceData Source:=ws.Range(ws.Cells(i, colStart), ws.Cells(i, colEnd))
    chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, colStart), ws.Cells(1, colEnd))
    chart.HasTitle = True
    chart.chartTitle.text = ws.Cells(i, "B").value
    chart.SetElement msoElementLegendNone
    chart.SeriesCollection(1).Format.Line.Weight = 1#

    SetYAxis chart, ws, i, maxVal
    SetXAxis chart

    ' Y���ڐ�����ǉ�
    With chart.Axes(xlValue, xlPrimary)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.Weight = 0.25
        .MajorGridlines.Format.Line.DashStyle = msoLineDashDot
    End With

    ' X���ڐ�����ǉ�
    With chart.Axes(xlCategory, xlPrimary)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.Weight = 0.25
        .MajorGridlines.Format.Line.DashStyle = msoLineDashDot
    End With
End Sub

' CreateGraphHelmet_Y���̏����ݒ�
Sub SetYAxis(ByRef chart As chart, ByRef ws As Worksheet, ByVal i As Long, ByVal maxVal As Double)
    Dim yAxis As Axis
    Set yAxis = chart.Axes(xlValue, xlPrimary)

    If maxVal >= 5# Then
        yAxis.MaximumScale = 10
        yAxis.MajorUnit = 2# '2.0����
    Else
        yAxis.MaximumScale = 5
        yAxis.MajorUnit = 1# '1.0����
    End If

    yAxis.MinimumScale = 0

    With yAxis.TickLabels
        .NumberFormatLocal = "0.0""kN"""
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With

End Sub


'CreateGraphHelmet_X���̏����ݒ�
Sub SetXAxis(ByRef chart As chart)
    Dim xAxis As Axis
    Set xAxis = chart.Axes(xlCategory, xlPrimary)

    xAxis.TickLabelSpacing = 100
    xAxis.TickMarkSpacing = 100

    With xAxis.TickLabels
        .NumberFormatLocal = "0.0""ms"""
        .Font.Color = RGB(89, 89, 89)
        .Font.size = 8
    End With

End Sub


Sub InspectHelmetDurationTime()
    ' �w�����b�g�����ɂ����čő�l�̍X�V�A�ő�l�̎��Ԃ̍X�V�A�������e�̍X�V�A�p�����Ԃ̐F�������s��
    Dim ws As Worksheet
    Dim lastRow As Long

    ' "LOG_Helmet" �V�[�g���w�肷��B
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    ' �ŏI�s���擾����B
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).row

    ' �e�s����������B
    For i = 2 To lastRow
        UpdateMaxValueInRow ws, i             ' �s���̍ő�l���X�V����B
        UpdatePartOfHelmet ws, i              ' �w�����b�g�̕������X�V����B
        UpdateRangeForThresholds ws, i, 4.9, "J"  ' 臒l�͈̔͂��X�V����B
        UpdateRangeForThresholds ws, i, 7.35, "K"
    Next i

End Sub

'InspectHelmetDurationTime()���̃v���V�[�W��_�s���̍ő�l���X�V���A�ő�l���L�^���������̃Z���ɐF������B
Sub UpdateMaxValueInRow(ByRef ws As Worksheet, ByVal row As Long)
    
    Dim rng As Range
    Dim MaxValue As Double
    Dim maxValueColumn As Long

    ' �s���͈̔͂��Z�b�g����B
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))
    ' �ő�l���擾����B
    MaxValue = Application.WorksheetFunction.Max(rng)
    ws.Cells(row, "H").value = MaxValue

    ' �ő�l�̈ʒu��������B
    For j = 1 To rng.Columns.Count
        If rng(1, j).value = MaxValue Then
            maxValueColumn = j
            rng(1, j).Interior.Color = RGB(250, 150, 0)  ' �F��ݒ肷��B
            Exit For
        End If
    Next j

    If maxValueColumn <> 0 Then
        ws.Cells(row, "I").value = ws.Cells(1, "V").Offset(0, maxValueColumn - 1).value
    End If
End Sub
 'InspectHelmetDurationTime()���̃v���V�[�W��_�w�����b�g�̎����ӏ����X�V����
Sub UpdatePartOfHelmet(ByRef ws As Worksheet, ByVal row As Long)

    Dim cellValue As String
    cellValue = ws.Cells(row, "B").value
    
    ' �����̒l���擾����B
    Dim existingValue As String
    existingValue = ws.Cells(row, "E").value

    ' �w�����b�g�̕����Ɋ�Â��Ēl��ݒ肷��B�������A"�V��"��"����"�����łɊ܂܂�Ă���ꍇ�͕ύX���Ȃ��B
    ' �����߂ł͍ŏ���E��̒l���`�F�b�N����B
    If InStr(existingValue, "�V��") > 0 Or InStr(existingValue, "����") > 0 Then
    ElseIf InStr(cellValue, "HEL_TOP") > 0 Then
        ws.Cells(row, "E").value = "�V��"
    ElseIf InStr(cellValue, "HEL_ZENGO") > 0 Then
        ws.Cells(row, "E").value = "�O�㓪��"
    ElseIf InStr(cellValue, "HEL_SIDE") > 0 Then
        ws.Cells(row, "E").value = "������"
    End If
End Sub

'InspectHelmetDurationTime()����4.9�A7.35�͈̔͒l�̐F�t���ƏՌ����Ԃ��L������B
Sub UpdateRangeForThresholds(ByRef ws As Worksheet, ByVal row As Long, ByVal threshold As Double, ByVal columnToWrite As String)

    Dim rng As Range, cell As Range
    Dim startRange As Long, endRange As Long, maxRange As Long
    Dim rangeCollection As New Collection
    Dim timeDifference As Double

    ' �s�͈̔͂��Z�b�g����B
    Set rng = ws.Range(ws.Cells(row, "V"), ws.Cells(row, ws.Columns.Count).End(xlToLeft))

    ' 臒l�𒴂���Z���͈̔͂�������B
    For Each cell In rng
        If cell.value >= threshold Then
            If startRange = 0 Then startRange = cell.Column
            endRange = cell.Column
            cell.Interior.Color = IIf(threshold = 4.9, RGB(255, 111, 56), RGB(234, 67, 53))
        Else
            If startRange > 0 And endRange > 0 Then
                rangeCollection.Add Array(startRange, endRange)
                startRange = 0
                endRange = 0
            End If
        End If
    Next cell

    If startRange > 0 And endRange > 0 Then rangeCollection.Add Array(startRange, endRange)

    For Each item In rangeCollection
        If (item(1) - item(0) + 1) > maxRange Then
            maxRange = item(1) - item(0) + 1
            startRange = item(0)
            endRange = item(1)
        End If
    Next item
    
    If startRange > 0 And endRange > 0 Then
        timeDifference = ws.Cells(1, endRange).value - ws.Cells(1, startRange).value
        If timeDifference < 0 Then
            timeDifference = 0
        End If
        ws.Cells(row, columnToWrite).value = timeDifference
    Else
        ws.Cells(row, columnToWrite).value = 0
    End If
End Sub


' TestCode---------------------------------------------------------------------------------------------
Sub GroupAndListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim part0 As String
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    ' �A�N�e�B�u�V�[�g�̃`���[�g�I�u�W�F�N�g�����[�v����
    For Each chartObj In ActiveSheet.ChartObjects
        ' �O���t�Ƀ^�C�g�������邩�ǂ������`�F�b�N
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"  ' �^�C�g�����Ȃ��ꍇ
        End If

        ' chartName��"-"�ŕ������Apart(0)���擾
        part0 = Split(chartObj.Name, "-")(0)

        ' �O���[�v���܂����݂��Ȃ��ꍇ�A�V�K�쐬
        If Not groups.Exists(part0) Then
            groups.Add part0, New Collection
        End If

        ' �O���[�v�Ƀ`���[�g���ƃ^�C�g����ǉ�
        groups(part0).Add "Chart Name: " & chartObj.Name & "; Title: " & chartTitle
    Next chartObj

    ' �e�O���[�v�̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo��
    Dim key As Variant
    For Each key In groups.Keys
        Debug.Print "Group: " & key
        Dim item As Variant
        For Each item In groups(key)
            Debug.Print item
        Next item
    Next key
End Sub



' �A�N�e�B�u�V�[�g�̃`���[�g�I�u�W�F�N�g��Debug.Print�ŕ\������B
Sub ListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String

    ' �A�N�e�B�u�V�[�g�̃`���[�g�I�u�W�F�N�g�����[�v����
    For Each chartObj In ActiveSheet.ChartObjects
        ' �O���t�Ƀ^�C�g�������邩�ǂ������`�F�b�N
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"  ' �^�C�g�����Ȃ��ꍇ
        End If

        ' �C�~�f�B�G�C�g�E�B���h�E�ɃO���t�̖��O�ƃ^�C�g����\��
        Debug.Print "Chart Name: " & chartObj.Name & "; Title: " & chartTitle
    Next chartObj
End Sub

' CreateChartID���@�\���Ă��邩�m�F����e�X�g�R�[�h
Sub TestCreateChartID()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")  ' �e�X�g���郏�[�N�V�[�g�����w��
    Dim testRange As Range
    Dim cell As Range
    Dim outputID As String

    ' �e�X�g�Ώۂ̃Z���͈͂��w��
    Set testRange = ws.Range("B2:B12")  ' B1����B5�܂ł̃Z�����e�X�g�ΏۂƂ���

    ' �e�Z���ɑ΂���CreateChartID�֐���K�p���A���ʂ��C�~�f�B�G�C�g�E�B���h�E�ɏo��
    For Each cell In testRange
        outputID = CreateChartID(cell)
        Debug.Print "Cell " & cell.Address & " Value: '" & cell.value & "' -> ID: '" & outputID & "'"
    Next cell
End Sub

