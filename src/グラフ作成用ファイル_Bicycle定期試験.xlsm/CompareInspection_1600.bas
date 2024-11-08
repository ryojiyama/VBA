Attribute VB_Name = "CompareInspection_1600"
Option Explicit

' ���������i�[����^��`
Private Type TestSample
    Number As String
    Condition As String
    Row As Long
End Type

' ����_�����i�[����^��`
Private Type MeasurementPoint
    Position As String
    Shape As String
    Row As Long
    Column As Long
End Type

' ��Ԃ̕ϊ��p�萔
Private Const HIGH_TEMP As String = "Hot"
Private Const LOW_TEMP As String = "Cold"
Private Const WET_CONDITION As String = "Wet"

' ������ϊ��p�̎�����������
Private Function InitializeConversionDicts() As Object
    Dim locationDict As Object
    Set locationDict = CreateObject("Scripting.Dictionary")
    
    ' �ʒu�ƌ`��̕ϊ��}�b�s���O
    locationDict.Add "�O����", "�O"
    locationDict.Add "�㓪��", "��"
    locationDict.Add "�E������", "�E"
    locationDict.Add "��������", "��"
    locationDict.Add "����", "��"
    locationDict.Add "����", "��"
    
    Set InitializeConversionDicts = locationDict
End Function

' ���C������
Sub �]�L����()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim convDict As Object
    Dim sheetNames As Variant
    Dim sheetName As Variant
    
    ' ������
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    Set convDict = InitializeConversionDicts()
    
    ' �����ΏۃV�[�g���̔z��
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' �e���i�V�[�g������
    For Each sheetName In sheetNames
        Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
        ProcessProductSheet productSheet, logSheet, convDict
    Next sheetName
End Sub

' ���i�V�[�g����
Private Sub ProcessProductSheet(ByRef productSheet As Worksheet, _
                              ByRef logSheet As Worksheet, _
                              ByRef convDict As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSample As TestSample
    Dim hasSample As Boolean
    Dim cellValue As String
    
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    ' �V�[�g�̊e�s������
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' �����s�̌��o
        If InStr(1, cellValue, "����") > 0 Then
            currentSample = GetSampleInfo(cellValue)
            currentSample.Row = i
            hasSample = True
        End If
        
        ' �Ռ��_�̌��o�Ƒ���_����
        If hasSample And InStr(1, cellValue, "�Ռ��_&�A���r��") > 0 Then
            Debug.Print "�Ռ��_���o - �V�[�g:" & productSheet.Name & ", �s:" & i & ", �l:" & cellValue
            ProcessMeasurementPoints productSheet, logSheet, i, currentSample, convDict
        End If
    Next i
End Sub

' �������̎擾
Private Function GetSampleInfo(ByVal cellValue As String) As TestSample
    Dim sample As TestSample
    Dim parts As Variant
    
    ' "����1 ����" �̂悤�ȕ�����𕪉�
    parts = Split(cellValue)
    
    ' �����ԍ��i2���ɐ��`�j
    sample.Number = Format(Val(Mid(parts(0), 3)), "00")
    
    ' ��Ԃ̔���
    Select Case parts(1)
        Case "����"
            sample.Condition = HIGH_TEMP
        Case "�ቷ"
            sample.Condition = LOW_TEMP
        Case "�Z����"
            sample.Condition = WET_CONDITION
    End Select
    
    ' �f�o�b�O�o��
    Debug.Print cellValue & " �� " & sample.Number & ", " & sample.Condition
    
    GetSampleInfo = sample
End Function

' ����_�̏���
Private Sub ProcessMeasurementPoints(ByRef productSheet As Worksheet, _
                                   ByRef logSheet As Worksheet, _
                                   ByVal currentRow As Long, _
                                   ByRef sample As TestSample, _
                                   ByRef convDict As Object)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim point As MeasurementPoint
    Dim searchCode As String
    Dim logLastRow As Long
    Dim i As Long
    Dim valueCell As Range
    Dim valueCellBelow As Range
    Dim foundMatch As Boolean
    Dim skippedLogs As String

    targetColumns = Array(2, 7)  ' B=2, G=7
    logLastRow = logSheet.Cells(Rows.count, "B").End(xlUp).Row
    skippedLogs = ""

    For Each colIndex In targetColumns
        point = GetMeasurementPoint(productSheet, currentRow, CLng(colIndex), convDict)
        If Len(point.Position) > 0 Then
            searchCode = sample.Number & "-500S-" & point.Position & "-" & _
                        sample.Condition & "-" & point.Shape

            Set valueCell = productSheet.Cells(currentRow + 1, CLng(colIndex) + 2)
            Set valueCellBelow = productSheet.Cells(currentRow + 2, CLng(colIndex) + 2)
            foundMatch = False

            For i = 2 To logLastRow
                Dim logValue As String
                logValue = logSheet.Cells(i, "B").value

                If Replace(logValue, "-E", "") = searchCode Then
                    foundMatch = True

                    If Len(Trim(logSheet.Cells(i, "V").value)) = 0 Then
                        ' �ŏ��̒l�iJ��j�̓]�L
                        If valueCell.MergeCells Then
                            valueCell.mergeArea.item(1).value = logSheet.Cells(i, "J").value
                        Else
                            valueCell.value = logSheet.Cells(i, "J").value
                        End If

                        ' ��ڂ̒l�iL��j�̓]�L
                        If valueCellBelow.MergeCells Then
                            valueCellBelow.mergeArea.item(1).value = logSheet.Cells(i, "L").value
                        Else
                            valueCellBelow.value = logSheet.Cells(i, "L").value
                        End If

                        logSheet.Cells(i, "V").value = "��"
                    Else
                        ' �X�L�b�v�������O���L�^
                        skippedLogs = skippedLogs & "�V�[�g: " & productSheet.Name & _
                                    ", �R�[�h: " & logValue & _
                                    ", LOG�s: " & i & _
                                    ", �l1: " & logSheet.Cells(i, "J").value & _
                                    ", �l2: " & logSheet.Cells(i, "L").value & vbCrLf
                    End If
                    Exit For
                End If
            Next i
        End If
    Next colIndex

    ' �X�L�b�v�������O������ꍇ�A�Ō�ɂ܂Ƃ߂ĕ\��
    If Len(skippedLogs) > 0 Then
        MsgBox "�ȉ��̃f�[�^�͊��ɓ]�L�ς݂̂��߃X�L�b�v����܂����F" & vbCrLf & vbCrLf & _
               skippedLogs, vbInformation, "�]�L�X�L�b�v���O"
    End If
End Sub

' ����_���̎擾
Private Function GetMeasurementPoint(ByRef sheet As Worksheet, _
                                   ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   ByRef convDict As Object) As MeasurementPoint
    Dim point As MeasurementPoint
    Dim targetCell As Range
    Dim cellValue As String
    Dim parts() As String
    
    Set targetCell = sheet.Cells(Row, Col + 1)
    
    If targetCell.MergeCells Then
        cellValue = Trim(targetCell.mergeArea.item(1).Offset(0, 1).value)
    Else
        cellValue = Trim(targetCell.Offset(0, 1).value)
    End If
    
    If Len(cellValue) > 0 Then
        parts = Split(cellValue, "�E")
        If UBound(parts) >= 1 Then
            point.Position = convDict(parts(0))
            point.Shape = convDict(parts(1))
        End If
    End If
    
    GetMeasurementPoint = point
End Function


' �`���[�g���z�̃��C������
' �`���[�g���z�̃��C������
Sub �`���[�g���z����()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim chartErrorLogs As String
    
    ' ������
    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    chartErrorLogs = ""
    
    Debug.Print "========== �`���[�g���z�����J�n ==========" & vbCrLf
    
    ' �����ΏۃV�[�g���̔z��
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' �e���i�V�[�g������
    For Each sheetName In sheetNames
        Debug.Print "�V�[�g�����J�n: " & CStr(sheetName)
        Set productSheet = ThisWorkbook.Sheets(CStr(sheetName))
        ProcessChartDistribution productSheet, logSheet, chartErrorLogs
    Next sheetName
    
    ' �G���[���O�̕\���ƃf�o�b�O�o��
    If Len(chartErrorLogs) > 0 Then
        Dim logMessage As String
        logMessage = "�`���[�g�����ňȉ��̖�肪�������܂����F" & vbCrLf & vbCrLf & chartErrorLogs
        
        ' �C�~�f�B�G�C�g�E�B���h�E�ɏo��
        Debug.Print "---------- �G���[���O ----------"
        Debug.Print logMessage
        Debug.Print "--------------------------------"
        
        ' ���b�Z�[�W�{�b�N�X�ŕ\��
        MsgBox logMessage, vbInformation, "�`���[�g��������"
    End If
    
    Debug.Print "========== �`���[�g���z�����I�� ==========" & vbCrLf
End Sub

' �e�V�[�g�̃`���[�g���z����
Private Sub ProcessChartDistribution(ByRef productSheet As Worksheet, _
                                   ByRef logSheet As Worksheet, _
                                   ByRef errorLogs As String)
    Dim lastRow As Long
    Dim i As Long
    Dim currentSample As TestSample
    Dim hasSample As Boolean
    Dim cellValue As String
    Dim convDict As Object
    
    ' ������
    Set convDict = InitializeConversionDicts()
    lastRow = productSheet.Cells(Rows.count, "B").End(xlUp).Row
    hasSample = False
    
    Debug.Print "�V�[�g[" & productSheet.Name & "] �����J�n - �ŏI�s: " & lastRow
    
    ' �V�[�g�̊e�s������
    For i = 1 To lastRow
        cellValue = Trim(productSheet.Cells(i, "B").value)
        
        ' �����s�̌��o
        If InStr(1, cellValue, "����") > 0 Then
            currentSample = GetSampleInfo(cellValue)
            hasSample = True
            Debug.Print "  �������o: " & cellValue & " -> �T���v���ԍ�: " & currentSample.Number & ", ���: " & currentSample.Condition
        End If
        
        ' �Ռ��_�̌��o�Ə���
        If hasSample And InStr(1, cellValue, "�Ռ��_&�A���r��") > 0 Then
            Debug.Print "  �Ռ��_���o - �s: " & i & ", �l: " & cellValue
            ProcessChartPoints productSheet, logSheet, i, currentSample, convDict, errorLogs
        End If
    Next i
    
    Debug.Print "�V�[�g[" & productSheet.Name & "] �����I��" & vbCrLf
End Sub

' ����_�̃`���[�g����
Private Sub ProcessChartPoints(ByRef productSheet As Worksheet, _
                             ByRef logSheet As Worksheet, _
                             ByVal currentRow As Long, _
                             ByRef sample As TestSample, _
                             ByRef convDict As Object, _
                             ByRef errorLogs As String)
    Dim targetColumns As Variant
    Dim colIndex As Variant
    Dim point As MeasurementPoint
    Dim searchCode As String
    Dim valueCell As Range
    
    targetColumns = Array(2, 7)  ' B=2, G=7
    
    For Each colIndex In targetColumns
        point = GetMeasurementPoint(productSheet, currentRow, CLng(colIndex), convDict)
        If Len(point.Position) > 0 Then
            ' �����p�R�[�h�̐���
            searchCode = sample.Number & "-500S-" & point.Position & "-" & _
                        sample.Condition & "-" & point.Shape & "-E"
            
            Debug.Print "    �����R�[�h����: " & searchCode
            
            ' �l���L�������Z���̈ʒu���擾
            Set valueCell = productSheet.Cells(currentRow + 1, CLng(colIndex) + 2)
            Debug.Print "    �ΏۃZ��: " & valueCell.Address
            
            ' �`���[�g�̃R�s�[����
            CopyMatchingChart logSheet, productSheet, searchCode, valueCell, errorLogs
        End If
    Next colIndex
End Sub

' �`���[�g�̃R�s�[����
Private Sub CopyMatchingChart(ByRef sourceSheet As Worksheet, _
                            ByRef targetSheet As Worksheet, _
                            ByVal searchID As String, _
                            ByRef targetCell As Range, _
                            ByRef errorLogs As String)
    Dim cht As ChartObject
    Dim foundCharts As Long
    Dim retryCount As Integer
    Const MAX_RETRIES As Integer = 3
    
    On Error Resume Next
    
    foundCharts = 0
    Debug.Print "      �`���[�g�����J�n - ID: " & searchID
    
    ' �\�[�X�V�[�g�̑S�`���[�g�����[�v
    For Each cht In sourceSheet.ChartObjects
        Debug.Print "        �m�F���̃`���[�g - Name: " & cht.Name
        
        ' �`���[�gID�ƌ���ID����v����ꍇ
        If cht.Name = searchID Then
            foundCharts = foundCharts + 1
            
            ' �ŏ��Ɍ��������`���[�g�̏ꍇ
            If foundCharts = 1 Then
                retryCount = 0
                Do
                    ' �N���b�v�{�[�h���N���A
                    Application.CutCopyMode = False
                    Err.Clear
                    
                    ' �`���[�g���R�s�[
                    cht.Copy
                    
                    If Err.Number = 0 Then
                        ' �����ҋ@���Ă���y�[�X�g
                        Application.Wait Now + TimeValue("00:00:00.2")
                        targetSheet.Paste targetCell.Offset(0, 3)
                        
                        If Err.Number = 0 Then
                            Debug.Print "        �`���[�g�𕡐�: " & targetCell.Offset(0, 3).Address & " �ɔz�u����"
                            Exit Do
                        End If
                    End If
                    
                    ' �G���[�����������ꍇ
                    If Err.Number <> 0 Then
                        retryCount = retryCount + 1
                        If retryCount >= MAX_RETRIES Then
                            errorLogs = errorLogs & "�z�u�G���[ - �V�[�g: " & targetSheet.Name & _
                                      ", ID: " & searchID & _
                                      ", �G���[: " & Err.Description & vbCrLf
                            Exit Do
                        End If
                        ' �Ď��s�O�ɏ������߂ɑҋ@
                        Application.Wait Now + TimeValue("00:00:00.5")
                    End If
                Loop While retryCount < MAX_RETRIES
                
            Else
                errorLogs = errorLogs & "�d���`���[�g - �V�[�g: " & targetSheet.Name & _
                           ", ID: " & searchID & _
                           " (" & foundCharts & "��)" & vbCrLf
            End If
        End If
    Next cht
    
    If foundCharts = 0 Then
        errorLogs = errorLogs & "�������`���[�g - �V�[�g: " & targetSheet.Name & _
                    ", ID: " & searchID & vbCrLf
    End If
    
    ' ����������ɃN���b�v�{�[�h���N���A
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub

