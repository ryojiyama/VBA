Attribute VB_Name = "TransferInspectionData"
Option Explicit

' ���������i�[����^��`
Private Type TestSample
    Number As String
    condition As String
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

'*******************************************************************************
' �^��`�ƒ萔
' ��������ё���_�̏����i�[����\���̂Ə�Ԃ������萔�̒�`
'*******************************************************************************

'*******************************************************************************
' ������ϊ��p�̎�����������
' �@�\�F�ʒu�ƌ`��̓��{��\�L���ȗ��`�ɕϊ����邽�߂̎������쐬
' �ߒl�FDictionary�I�u�W�F�N�g�i�O�������O�A�㓪������A�Ȃǁj
'*******************************************************************************
Private Function InitializeConversionDicts()
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

'*******************************************************************************
' ���C���v���V�[�W��
' �T�v�FLOG_Bicycle�V�[�g����e���i�V�[�g�֎����f�[�^��]�L
' �ΏہF500S_1, 500S_2, 500S_3�V�[�g�̃f�[�^�]�L
' �ˑ��FInitializeConversionDicts, ProcessProductSheet
'*******************************************************************************
Sub TransferBicycleTestData()
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

'*******************************************************************************
' TransferBicycleTestData�̃T�u�v���V�[�W��
' �@�\�F�ʃV�[�g�̎����f�[�^�����o���A�]�L���������s
' �����FproductSheet - �]�L��̐��i�V�[�g
'       logSheet     - �f�[�^����LOG�V�[�g
'       convDict     - �ϊ��p�����I�u�W�F�N�g
' �����F�������̌��o�Ƒ���_�f�[�^�̓]�L�����s
'*******************************************************************************
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

'*******************************************************************************
' ProcessProductSheet�̃T�u�v���V�[�W��
' �@�\�F�Z���̕����񂩂玎���ԍ��Ǝ��������𒊏o
' �����FcellValue - "����1 ����" �`���̕�����
' �ߒl�FTestSample�^�iNumber = "01", Condition = "Hot" �Ȃǁj
'*******************************************************************************
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
            sample.condition = HIGH_TEMP
        Case "�ቷ"
            sample.condition = LOW_TEMP
        Case "�Z����"
            sample.condition = WET_CONDITION
    End Select
    
    ' �f�o�b�O�o��
    Debug.Print cellValue & " �� " & sample.Number & ", " & sample.condition
    
    GetSampleInfo = sample
End Function

'*******************************************************************************
' ProcessProductSheet�̃T�u�v���V�[�W��
' �@�\�F���o���ꂽ����_�̃f�[�^��LOG�V�[�g���琻�i�V�[�g�֓]�L
' �����FproductSheet - �]�L��V�[�g
'       logSheet     - LOG���V�[�g
'       currentRow   - �������̍s�ԍ�
'       sample       - �������
'       convDict     - �ϊ�����
' ���L�F���]�L�f�[�^�̃X�L�b�v�����ƃ��O�o�͂��܂�
'*******************************************************************************
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
                        sample.condition & "-" & point.Shape

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

'*******************************************************************************
' ProcessMeasurementPoints�̃T�u�v���V�[�W��
' �@�\�F�V�[�g�̃Z�����瑪��_�̈ʒu�ƌ`��𒊏o
' �����Fsheet    - �ΏۃV�[�g
'       Row      - �Ώۍs
'       Col      - �Ώۗ�
'       convDict - �ϊ�����
' �ߒl�FMeasurementPoint�^�iPosition = "�O", Shape = "��" �Ȃǁj
' ���L�F�}�[�W�Z���ɑΉ�
'*******************************************************************************
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


