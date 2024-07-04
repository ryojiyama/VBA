Attribute VB_Name = "TransferData"
' "LOG_Helmet"�̃f�[�^���e�V�[�g�ɕ��z����
Sub TransferDataBasedOnID()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim nextRow As Long
    Dim data As Collection
    Dim dataItem As Variant
    
    ' �\�[�X�V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' �\�[�X�V�[�g�̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' �e�s�����[�v����
    For i = 1 To lastRow
        ' ID�𕪊�
        idParts = Split(wsSource.Cells(i, 3).value, "-")
        If UBound(idParts) >= 2 Then
            ' �O���[�v���i���ʁj���擾
            group = idParts(2)
            
            ' �O���[�v���Ɋ�Â��ăV�[�g����ݒ�
            Select Case group
                Case "�V"
                    targetSheetName = "Impact_Top"
                Case "�O"
                    targetSheetName = "Impact_Front"
                Case "��"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' �Ή�����O���[�v���Ȃ��ꍇ�̓X�L�b�v
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).value
                    GoTo NextIteration
            End Select
            
            ' �f�[�^���R���N�V�����ɒǉ�
            data.Add Array(i, targetSheetName)
        End If
NextIteration:
    Next i
    
    ' �R���N�V��������e�V�[�g�Ƀf�[�^��]�L
    For Each dataItem In data
        Dim rowIndex As Long
        rowIndex = dataItem(0)
        targetSheetName = dataItem(1)
        
        ' �ړI�̃V�[�g���쐬
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            wsDest.name = targetSheetName
        End If
        On Error GoTo 0
        
        ' �w�b�_�[�s��ݒ�iB15�Z���ɐݒ�j
        If wsDest.Range("B15").value = "" Then
            wsSource.Range("B1:Z1").Copy Destination:=wsDest.Range("B15")
        End If
        
        ' �ŏI�s�������A���̍s����f�[�^�̓]�L���J�n���܂�
        nextRow = wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).row + 1
        If nextRow < 16 Then
            nextRow = 16 ' �ŏ��̃f�[�^�]�L�J�n�ʒu��B16�ɐݒ�
        End If
        
        ' �f�[�^�͈͂�]�L
        wsSource.Range("B" & rowIndex & ":Z" & rowIndex).Copy Destination:=wsDest.Range("B" & nextRow)
    Next dataItem

    ' ���\�[�X�����
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub

Sub GroupAndListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim partEnd As String
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

        ' chartName��"-"�ŕ������A�Ō�̕������擾
        partEnd = Split(chartObj.name, "-")(UBound(Split(chartObj.name, "-")))

        ' �O���[�v���܂����݂��Ȃ��ꍇ�A�V�K�쐬
        If Not groups.Exists(partEnd) Then
            groups.Add partEnd, New Collection
        End If

        ' �O���[�v�Ƀ`���[�g���ƃ^�C�g����ǉ�
        groups(partEnd).Add "Chart Name: " & chartObj.name & "; Title: " & chartTitle
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


'Sub GroupAndListChartNamesAndTitles()
'    Dim chartObj As ChartObject
'    Dim chartTitle As String
'    Dim part0 As String
'    Dim groups As Object
'    Set groups = CreateObject("Scripting.Dictionary")
'
'    ' �A�N�e�B�u�V�[�g�̃`���[�g�I�u�W�F�N�g�����[�v����
'    For Each chartObj In ActiveSheet.ChartObjects
'        ' �O���t�Ƀ^�C�g�������邩�ǂ������`�F�b�N
'        If chartObj.chart.HasTitle Then
'            chartTitle = chartObj.chart.chartTitle.text
'        Else
'            chartTitle = "No Title"  ' �^�C�g�����Ȃ��ꍇ
'        End If
'
'        ' chartName��"-"�ŕ������Apart(0)���擾
'        part0 = Split(chartObj.name, "-")(0)
'
'        ' �O���[�v���܂����݂��Ȃ��ꍇ�A�V�K�쐬
'        If Not groups.Exists(part0) Then
'            groups.Add part0, New Collection
'        End If
'
'        ' �O���[�v�Ƀ`���[�g���ƃ^�C�g����ǉ�
'        groups(part0).Add "Chart Name: " & chartObj.name & "; Title: " & chartTitle
'    Next chartObj
'
'    ' �e�O���[�v�̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo��
'    Dim key As Variant
'    For Each key In groups.Keys
'        Debug.Print "Group: " & key
'        Dim item As Variant
'        For Each item In groups(key)
'            Debug.Print item
'        Next item
'    Next key
'End Sub

' �O���t�̕��ג���
Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim groupName As String
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"�V�[�g��Ώۂɂ���
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"�V�[�g�̃`���[�g�I�u�W�F�N�g���O���[�v����
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartName��"-"�ŕ������A�Ō�̕������擾
        groupName = Split(chartObj.name, "-")(UBound(Split(chartObj.name, "-")))
        If Not groups.Exists(groupName) Then
            groups.Add groupName, New Collection
        End If
        groups(groupName).Add chartObj
    Next chartObj
    
    ' �O���[�v���ƂɃ`���[�g��Ή�����V�[�g�Ɉړ�
    Dim key As Variant
    'Dim SheetName As String
    For Each key In groups.Keys
        ' �O���[�v���Ɋ�Â��ăV�[�g��������
        Select Case key
            Case "�V"
                SheetName = "Impact_Top"
            Case "�O"
                SheetName = "Impact_Front"
            Case "��"
                SheetName = "Impact_Back"
            Case Else
                SheetName = "" ' �Y�����Ȃ��ꍇ�͋�̃V�[�g��
        End Select
        
        If SheetName <> "" Then
            ' �V�[�g�̑��݂��m�F
            On Error Resume Next
            Set targetSheet = ThisWorkbook.Sheets(SheetName)
            On Error GoTo 0
            
            ' �V�[�g�����݂���ꍇ�̂݃`���[�g���ړ�
            If Not targetSheet Is Nothing Then
                ' �`���[�g�̈ړ�
                Dim chart As ChartObject
                For Each chart In groups(key)
                    chart.chart.Location Where:=xlLocationAsObject, name:=targetSheet.name
                Next chart
                Set targetSheet = Nothing
            Else
                Debug.Print "Sheet " & SheetName & " does not exist. Charts not moved."
            End If
        Else
            Debug.Print "Group " & key & " does not have a corresponding sheet. Charts not moved."
        End If
    Next key
End Sub

