Attribute VB_Name = "Test"
Sub TransferDataToDynamicSheets()

    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim lastRow As Long, i As Long
    Dim sourceData As String
    Dim parts() As String
    Dim destinationSheetName As String
    Dim productNum As String

    ' �\�[�X�V�[�g�̐ݒ�
    Set wsSource = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).Row

    ' Excel�̃p�t�H�[�}���X����̂��߂̐ݒ�
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    ' wsSource��C������[�v���ăf�[�^������
    For i = 2 To lastRow
        sourceData = wsSource.Cells(i, "B").value
        parts = Split(sourceData, "-")

        ' �V�[�g���𐶐����AwsDesitination�ɑ���B
        If UBound(parts) >= 2 Then
            destinationSheetName = parts(1) & "_" & 1

            ' �]�L��V�[�g�̑��݊m�F
            On Error Resume Next
            Set wsDestination = ThisWorkbook.Sheets(destinationSheetName)
            On Error GoTo 0

            ' �V�[�g�����݂���ꍇ�Ƀf�[�^��]�L
            If Not wsDestination Is Nothing Then
                productNum = wsSource.Cells(i, "D").value
                wsDestination.Range("D3").value = "No." & Left(productNum, Len(productNum) - 1) & "-" & Right(productNum, 1)
                wsDestination.Range("D4").value = wsSource.Cells(i, "O").value
                wsDestination.Range("D5").value = wsSource.Cells(i, "E").value
                wsDestination.Range("D6").value = wsSource.Cells(i, "Q").value
                wsDestination.Range("I3").value = wsSource.Cells(i, "F").value
                wsDestination.Range("I4").value = wsSource.Cells(i, "G").value
                ' �����f�[�^�̓]�L
                wsDestination.Range("D22").value = wsSource.Cells(i, "J").value
                wsDestination.Range("D23").value = wsSource.Cells(i, "L").value
            End If
        End If
    Next i

    ' Excel�̐ݒ�����ɖ߂�
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
' ������ϊ��p�̊֐�
Function ConvertCompareString(ByVal strValue As String) As String
    ' �����֘A�̕ϊ�
    strValue = Replace(strValue, "�O����", "�O")
    strValue = Replace(strValue, "�㓪��", "��")
    strValue = Replace(strValue, "�E������", "�E")
    strValue = Replace(strValue, "��������", "��")
    
    ' �`��֘A�̕ϊ�
    strValue = Replace(strValue, "����", "��")
    strValue = Replace(strValue, "����", "��")
    
    ConvertCompareString = strValue
End Function

Sub �]�L����()
    Dim logSheet As Worksheet
    Dim productSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long, sheetNum As Long, j As Long, k As Long
    Dim productCode As String, productName As String, productSheetName As String
    Dim parts() As String, inspectionSheetPartsB() As String, inspectionSheetPartsG() As String
    Dim impactCellB As String, impactCellG As String
    Dim impactRowsB As Variant, impactRowsG As Variant
    Dim foundB As Boolean, foundG As Boolean
    Dim mergeArea As Range

    Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
    lastRow = logSheet.Cells(Rows.count, "B").End(xlUp).Row

    For i = 2 To lastRow
        productCode = logSheet.Cells(i, "B").value
        parts = Split(productCode, "-")
        productName = parts(1)
        foundB = False
        foundG = False

        For sheetNum = 1 To 3
            productSheetName = productName & "_" & sheetNum

            On Error Resume Next
            Set productSheet = ThisWorkbook.Sheets(productSheetName)
            On Error GoTo 0

            If Not productSheet Is Nothing Then
                impactRowsB = FindAllRows(productSheet, "B", "�Ռ��_&�A���r��")
                impactRowsG = FindAllRows(productSheet, "G", "�Ռ��_&�A���r��")

                ' B��̏���
                If IsArray(impactRowsB) Then
                    For j = LBound(impactRowsB) To UBound(impactRowsB)
                        If productSheet.Cells(impactRowsB(j), "B").MergeCells Then
                            Set mergeArea = productSheet.Cells(impactRowsB(j), "B").mergeArea
                            Dim nextColB As Long
                            nextColB = mergeArea.Column + mergeArea.Columns.count
                            impactCellB = productSheet.Cells(impactRowsB(j), nextColB).value

                            If Len(Trim(impactCellB)) > 0 Then
                                inspectionSheetPartsB = Split(impactCellB, "�E")
                                If UBound(inspectionSheetPartsB) >= 1 Then
                                    Dim convertedFirst As String, convertedSecond As String
                                    convertedFirst = ConvertCompareString(inspectionSheetPartsB(0))
                                    convertedSecond = ConvertCompareString(inspectionSheetPartsB(1))

                                    If parts(2) = convertedFirst And parts(4) = convertedSecond Then
                                        productSheet.Cells(impactRowsB(j), nextColB).value = logSheet.Cells(i, "J").value
                                        foundB = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If

                ' G��̏���
                If IsArray(impactRowsG) Then
                    For k = LBound(impactRowsG) To UBound(impactRowsG)
                        If productSheet.Cells(impactRowsG(k), "G").MergeCells Then
                            Set mergeArea = productSheet.Cells(impactRowsG(k), "G").mergeArea
                            Dim nextColG As Long
                            nextColG = mergeArea.Column + mergeArea.Columns.count
                            impactCellG = productSheet.Cells(impactRowsG(k), nextColG).value

                            If Len(Trim(impactCellG)) > 0 Then
                                inspectionSheetPartsG = Split(impactCellG, "�E")
                                If UBound(inspectionSheetPartsG) >= 1 Then
                                    convertedFirst = ConvertCompareString(inspectionSheetPartsG(0))
                                    convertedSecond = ConvertCompareString(inspectionSheetPartsG(1))

                                    If parts(2) = convertedFirst And parts(4) = convertedSecond Then
                                        productSheet.Cells(impactRowsG(k), nextColG).value = logSheet.Cells(i, "J").value
                                        foundG = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If

                If foundB Or foundG Then Exit For
            End If
            Set productSheet = Nothing
        Next sheetNum
    Next i
End Sub

Function FindAllRows(sheet As Worksheet, Col As String, searchStr As String) As Variant
    Dim result(0 To 1000) As Long
    Dim resultCount As Long
    Dim lastRow As Long
    Dim i As Long
    Dim mergeArea As Range

    resultCount = -1
    lastRow = sheet.Cells(sheet.Rows.count, Col).End(xlUp).Row

    For i = 1 To lastRow
        If sheet.Cells(i, Col).MergeCells Then
            Set mergeArea = sheet.Cells(i, Col).mergeArea
            If sheet.Cells(i, Col).Address = mergeArea.Cells(1, 1).Address Then
                If InStr(1, mergeArea.Cells(1, 1).value, searchStr) > 0 Then
                    resultCount = resultCount + 1
                    result(resultCount) = i
                End If
            End If
        Else
            If InStr(1, sheet.Cells(i, Col).value, searchStr) > 0 Then
                resultCount = resultCount + 1
                result(resultCount) = i
            End If
        End If
    Next i

    If resultCount >= 0 Then
        Dim finalResult() As Long
        ReDim finalResult(0 To resultCount)
        For i = 0 To resultCount
            finalResult(i) = result(i)
        Next i
        FindAllRows = finalResult
    Else
        FindAllRows = Array()
    End If
End Function

Sub �Z���l�m�F()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim targetCells As Variant
    Dim i As Long, j As Long
    
    ' �m�F����V�[�g����z��Ɋi�[
    sheetNames = Array("500S_1", "500S_2", "500S_3")
    
    ' �m�F����Z���̈ʒu��z��Ɋi�[ (��, �s)
    targetCells = Array(Array("B", 21), Array("B", 25), Array("G", 21), Array("G", 25))
    
    Debug.Print "�Z���l�m�F�J�n"
    Debug.Print "-------------------"
    
    ' �e�V�[�g�����[�v
    For i = 0 To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            Debug.Print sheetNames(i) & " �̃Z���l:"
            
            ' �e�ΏۃZ�������[�v
            For j = 0 To UBound(targetCells)
                Dim Col As String
                Dim Row As Long
                Col = targetCells(j)(0)
                Row = targetCells(j)(1)
                
                ' �Z������������Ă��邩�m�F
                If ws.Cells(Row, Col).MergeCells Then
                    Dim mergeArea As Range
                    Set mergeArea = ws.Cells(Row, Col).mergeArea
                    Debug.Print "  " & Col & Row & ": " & mergeArea.Cells(1, 1).value & _
                              " (�����Z��: " & mergeArea.Address & ")"
                Else
                    Debug.Print "  " & Col & Row & ": " & ws.Cells(Row, Col).value
                End If
            Next j
            Debug.Print "-------------------"
        Else
            Debug.Print "�V�[�g " & sheetNames(i) & " ��������܂���"
            Debug.Print "-------------------"
        End If
    Next i
    
    Debug.Print "�m�F����"
End Sub



' -------------------------------------------------------------------------------------------------------------
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

Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
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
        
        ' chartName��"-"�ŕ������AsheetName���擾
        parts = Split(chartObj.Name, "-")
        If UBound(parts) >= 1 Then
            sheetName = parts(0) & "-" & parts(1)
        Else
            sheetName = parts(0)
        End If
        
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' �O���[�v���ƂɃ`���[�g��Ή�����V�[�g�Ɉړ�
    Dim key As Variant
    For Each key In groups.Keys
        ' �V�[�g�̑��݂��m�F
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' �V�[�g�����݂��Ȃ��ꍇ�A�`���[�g���ړ����Ȃ�
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' �`���[�g�̈ړ�
            Dim chart As ChartObject
            For Each chart In groups(key)
                chart.chart.Location Where:=xlLocationAsObject, Name:=targetSheet.Name
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not moved."
        End If
    Next key
End Sub





