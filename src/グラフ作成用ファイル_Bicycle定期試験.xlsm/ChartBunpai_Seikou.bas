Attribute VB_Name = "ChartBunpai_Seikou"
' 2024-11-08�쐬 ���܂��s�������A�ۑ�������B


' ===== �^��`���ŏ��ɔz�u =====
Private Type sampleInfo
    SampleNumber As String
    Condition As String
End Type

Private Type pointInfo
    Position As String
    Shape As String
End Type

' ===== �萔��` =====
Private Const CHART_SUFFIX As String = "-E"
Private Const SERIES_PREFIX As String = "500S"

' ===== �����I�u�W�F�N�g�i�[�p�ϐ� =====
Private locationDict As Object
Private conditionDict As Object
Private shapeDict As Object
Private searchPatternDict As Object


' ===== �����̏����� =====
Private Sub InitializeDictionaries()
    On Error GoTo ErrorHandler
    
    ' �����̎������N���A
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    ' �ʒu�ϊ��p����
    Set locationDict = CreateObject("Scripting.Dictionary")
    With locationDict
        .Add "�O����", "�O"
        .Add "�㓪��", "��"
        .Add "�E������", "�E"
        .Add "��������", "��"
    End With
    
    ' ��ԕϊ��p����
    Set conditionDict = CreateObject("Scripting.Dictionary")
    With conditionDict
        .Add "����", "Hot"
        .Add "�ቷ", "Cold"
        .Add "�Z����", "Wet"
    End With
    
    ' �`��ϊ��p����
    Set shapeDict = CreateObject("Scripting.Dictionary")
    With shapeDict
        .Add "����", "��"
        .Add "����", "��"
    End With
    
    ' �����p�^�[���p����
    Set searchPatternDict = CreateObject("Scripting.Dictionary")
    With searchPatternDict
        .Add "format", "{0}-" & SERIES_PREFIX & "-{1}-{2}-{3}" & CHART_SUFFIX
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "�����̏������G���[: " & Err.Description
    
    ' �����I�u�W�F�N�g�̃N���[���A�b�v
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    Err.Raise Err.Number, "InitializeDictionaries", "�����̏������Ɏ��s���܂���"
End Sub

' ===== ���C���̃`���[�g���z���� =====
Sub �`���[�g���z����()
    On Error GoTo ErrorHandler
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.screenUpdating = False
    
    ' �����̏�����
    InitializeDictionaries
    
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim chartErrorLogs As String
    
    ' �V�[�g�̑��݊m�F
    If Not SheetExists("LOG_Bicycle") Then
        MsgBox "LOG_Bicycle�V�[�g��������܂���B", vbCritical
        GoTo CleanUp
    End If
    
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    chartErrorLogs = ""
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' �e�V�[�g�̏���
    For Each sheetName In sheetNames
        If SheetExists(CStr(sheetName)) Then
            Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
            ProcessSheet productSheet, logSheet, chartErrorLogs
        Else
            chartErrorLogs = chartErrorLogs & "�V�[�g��������܂���: " & CStr(sheetName) & vbCrLf
        End If
    Next sheetName

CleanUp:
    ' �����I�u�W�F�N�g�̉��
    Set locationDict = Nothing
    Set conditionDict = Nothing
    Set shapeDict = Nothing
    Set searchPatternDict = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.screenUpdating = True
    
    If Len(chartErrorLogs) > 0 Then
        MsgBox "�`���[�g�����ňȉ��̖�肪�������܂����F" & vbCrLf & vbCrLf & _
               chartErrorLogs, vbInformation, "�`���[�g��������"
    End If
    Exit Sub

ErrorHandler:
    chartErrorLogs = chartErrorLogs & "�\�����ʃG���[: " & Err.Description & vbCrLf
    Resume CleanUp
End Sub

' ===== �V�[�g���� =====
Private Sub ProcessSheet(ByRef productSheet As Worksheet, _
                        ByRef logSheet As Worksheet, _
                        ByRef errorLogs As String)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSampleInfo As sampleInfo
    Dim hasSample As Boolean
    Dim cellValue As String
    
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' �������̎擾
        If InStr(1, cellValue, "����") > 0 Then
            currentSampleInfo = ExtractSampleInfo(cellValue)
            hasSample = True
        End If
        
        ' �Ռ��_�̏���
        If hasSample And InStr(1, cellValue, "�Ռ��_&�A���r��") > 0 Then
            ProcessMeasurementPoint productSheet, logSheet, i, currentSampleInfo, errorLogs
        End If
    Next i
End Sub

' ===== ������񒊏o =====
Private Function ExtractSampleInfo(ByVal cellValue As String) As sampleInfo
    Dim info As sampleInfo
    Dim parts As Variant
    
    On Error GoTo ErrorHandler
    
    parts = Split(cellValue)
    
    ' �z��̋��E�`�F�b�N
    If UBound(parts) >= 1 Then
        ' �����ԍ��̒��o�i���l�ȊO�̕����������j
        Dim numStr As String
        numStr = Replace(Replace(parts(0), "����", ""), " ", "")
        info.SampleNumber = Format(Val(numStr), "00")
        
        ' ��Ԃ̔���i�����ɑ��݂��邩�`�F�b�N�j
        If conditionDict.Exists(parts(1)) Then
            info.Condition = conditionDict(parts(1))
        Else
            info.Condition = "Unknown"
            Debug.Print "����`�̏��: " & parts(1)
        End If
    End If
    
    ExtractSampleInfo = info
    Exit Function
    
ErrorHandler:
    info.SampleNumber = "00"
    info.Condition = "Error"
    ExtractSampleInfo = info
End Function

' ===== ����_���� =====
Private Sub ProcessMeasurementPoint(ByRef productSheet As Worksheet, _
                                  ByRef logSheet As Worksheet, _
                                  ByVal currentRow As Long, _
                                  ByRef sampleInfo As sampleInfo, _
                                  ByRef errorLogs As String)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim pointInfo As pointInfo
    Dim chartId As String
    
    targetColumns = Array(2, 7)  ' B=2, G=7
    
    For Each colIndex In targetColumns
        pointInfo = ExtractPointInfo(productSheet, currentRow, CLng(colIndex))
        If Len(pointInfo.Position) > 0 Then
            ' �`���[�gID�̐���
            chartId = GenerateChartId(logSheet, sampleInfo, pointInfo)
            
            ' �`���[�g�̃R�s�[
            CopyChart logSheet, productSheet, chartId, _
                     productSheet.Cells(currentRow + 1, CLng(colIndex) + 2), errorLogs
        End If
    Next colIndex
End Sub

' ===== ����_��񒊏o =====
Private Function ExtractPointInfo(ByRef sheet As Worksheet, _
                                ByVal Row As Long, _
                                ByVal Col As Long) As pointInfo
    Dim info As pointInfo
    Dim targetCell As Range
    Dim cellValue As String
    Dim parts() As String
    
    On Error GoTo ErrorHandler
    
    Set targetCell = sheet.Cells(Row, Col + 1)
    
    If targetCell.MergeCells Then
        cellValue = Trim(targetCell.mergeArea.item(1).Offset(0, 1).value)
    Else
        cellValue = Trim(targetCell.Offset(0, 1).value)
    End If
    
    If Len(cellValue) > 0 Then
        parts = Split(cellValue, "�E")
        If UBound(parts) >= 1 Then
            ' �����̑��݃`�F�b�N
            If locationDict.Exists(parts(0)) And shapeDict.Exists(parts(1)) Then
                info.Position = locationDict(parts(0))
                info.Shape = shapeDict(parts(1))
            Else
                Debug.Print "����`�̈ʒu�܂��͌`��: " & cellValue
            End If
        End If
    End If
    
    ExtractPointInfo = info
    Exit Function
    
ErrorHandler:
    info.Position = ""
    info.Shape = ""
    ExtractPointInfo = info
End Function

' ===== �`���[�gID���� =====
Private Function GetChartPattern(ByRef logSheet As Worksheet) As String
    ' �����l�i�G���[���̃t�H�[���o�b�N�p�j
    Dim suffixPattern As String: suffixPattern = "-XX"
    Dim seriesPrefix As String: seriesPrefix = "Sample"
    
    On Error Resume Next
    
    ' �V�[�g����l���擾
    If Not logSheet Is Nothing Then
        ' �T�t�B�b�N�X�p�^�[���̎擾
        If Len(Trim(logSheet.Cells(2, "Q").value)) > 0 Then
            suffixPattern = Trim(logSheet.Cells(2, "Q").value)
            If Left(suffixPattern, 1) <> "-" Then
                suffixPattern = "-" & suffixPattern
            End If
        End If
        
        ' �V���[�Y�v���t�B�b�N�X�̎擾
        If Len(Trim(logSheet.Cells(2, "D").value)) > 0 Then
            seriesPrefix = Trim(logSheet.Cells(2, "D").value)
        End If
    End If
    
    ' �p�^�[��������𐶐�
    GetChartPattern = "{0}-" & seriesPrefix & "-{1}-{2}-{3}" & suffixPattern
End Function

' ===== �`���[�gID�����i�C���Łj =====
Private Function GenerateChartId(ByRef logSheet As Worksheet, _
                               ByRef sampleInfo As sampleInfo, _
                               ByRef pointInfo As pointInfo) As String
    On Error GoTo ErrorHandler
    
    If Len(sampleInfo.SampleNumber) = 0 Or Len(sampleInfo.Condition) = 0 _
       Or Len(pointInfo.Position) = 0 Or Len(pointInfo.Shape) = 0 Then
        GenerateChartId = ""
        Exit Function
    End If
    
    ' �p�^�[���𓮓I�Ɏ擾
    Dim pattern As String
    pattern = GetChartPattern(logSheet)
    
    ' ID�𐶐�
    GenerateChartId = Replace(pattern, "{0}", sampleInfo.SampleNumber)
    GenerateChartId = Replace(GenerateChartId, "{1}", pointInfo.Position)
    GenerateChartId = Replace(GenerateChartId, "{2}", sampleInfo.Condition)
    GenerateChartId = Replace(GenerateChartId, "{3}", pointInfo.Shape)
    Exit Function
    
ErrorHandler:
    GenerateChartId = ""
End Function

' ===== �`���[�g�R�s�[���� =====
Private Sub CopyChart(ByRef sourceSheet As Worksheet, _
                     ByRef targetSheet As Worksheet, _
                     ByVal chartId As String, _
                     ByRef targetCell As Range, _
                     ByRef errorLogs As String)
    On Error GoTo ErrorHandler
    
    Dim cht As ChartObject
    Dim foundCharts As Long
    Const WAIT_TIME As String = "0:00:02.20"  ' �ҋ@���Ԃ�2.20�b�ɉ���
    Dim retryCount As Integer
    Const MAX_RETRY As Integer = 2  ' ���g���C��
    
    foundCharts = 0
    Debug.Print "����ID: " & chartId
    
    For Each cht In sourceSheet.ChartObjects
        Debug.Print "�`�F�b�N���̃`���[�g: " & cht.Name
        
        If cht.Name = chartId Then
            foundCharts = foundCharts + 1
            If foundCharts = 1 Then
                ' �I���W�i���̃T�C�Y��ۑ�
                Dim originalWidth As Double
                Dim originalHeight As Double
                originalWidth = cht.Width
                originalHeight = cht.Height
                
                ' �R�s�[�����i���g���C�t���j
                For retryCount = 0 To MAX_RETRY
                    ' �N���b�v�{�[�h���N���A
                    Application.CutCopyMode = False
                    DoEvents
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' �`���[�g���R�s�[
                    cht.Copy
                    DoEvents
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' �y�[�X�g���s
                    targetSheet.Paste Destination:=targetCell.Offset(0, 3)
                    DoEvents
                    
                    ' �V�����`���[�g�̃T�C�Y�𒲐�
                    Dim newChart As ChartObject
                    Set newChart = targetSheet.ChartObjects(targetSheet.ChartObjects.count)
                    
                    ' �T�C�Y�����̃`���[�g�ɍ��킹��
                    newChart.Width = originalWidth
                    newChart.Height = originalHeight
                    
                    ' ������ɑҋ@
                    Application.Wait Now + TimeValue(WAIT_TIME)
                    
                    ' �N���b�v�{�[�h���N���A
                    Application.CutCopyMode = False
                    
                    ' �����m�F
                    If Not newChart Is Nothing Then Exit For
                    
                    ' ���g���C���̃��O
                    If retryCount < MAX_RETRY Then
                        Debug.Print "�R�s�[�Ď��s: " & (retryCount + 1) & " ���"
                    End If
                Next retryCount
            Else
                errorLogs = errorLogs & "�d���`���[�g - �V�[�g: " & targetSheet.Name & _
                           ", ID: " & chartId & " (" & foundCharts & "��)" & vbCrLf
            End If
        End If
    Next cht
    
    If foundCharts = 0 Then
        errorLogs = errorLogs & "�������`���[�g - �V�[�g: " & targetSheet.Name & _
                   ", ID: " & chartId & vbCrLf
    End If
    
    Exit Sub

ErrorHandler:
    errorLogs = errorLogs & "�G���[���� - �V�[�g: " & targetSheet.Name & _
               ", ID: " & chartId & ", �G���[: " & Err.Description & vbCrLf
    Resume Next
End Sub



' ===== ���[�e�B���e�B�֐� =====
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function

