Attribute VB_Name = "Module1"
Sub ExampleChartNameAndTitle()
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects(1)
    
    ' �O���t�̖��O��ݒ�
    chartObj.Name = "MyCustomChartName"
    
    ' �O���t�̃^�C�g����ݒ�
    If Not chartObj.chart.HasTitle Then
        chartObj.chart.SetElement msoElementChartTitleAboveChart
    End If
    chartObj.chart.chartTitle.text = "My Chart Title"
End Sub

Sub TransferValues_old()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    
    ' �V�[�g���Z�b�g
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' �w�b�_�[�̗�ԍ����擾
    colHinban = 0
    colBoutai = 0
    colTencho = 0
    
    For Each cell In wsHelSpec.Rows(1).Cells
        If cell.value = "�i��(D)" Then
            colHinban = cell.column
        ElseIf cell.value = "�V������" Then
            colTencho = cell.column
        End If
        If colHinban > 0 And colTencho > 0 Then Exit For
    Next cell
    
    For Each cell In wsSetting.Rows(1).Cells
        If cell.value = "�X��No." Then
            colBoutai = cell.column
            Exit For
        End If
    Next cell
    
    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "�i��(D)" ��̒l��T�����A�]�L
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
End Sub

Sub CopyAndSubtractValues_old()
    Dim wsHelSpec As Worksheet
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim i As Long
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim cell As Range
    
    ' �V�[�g���Z�b�g
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' �w�b�_�[�̗�ԍ����擾
    colTenchoSukima = 0
    colSokuteiSukima = 0
    colTenchoNikui = 0
    
    For Each cell In wsHelSpec.Rows(1).Cells
        If cell.value = "�V��������(N)" Then
            colTenchoSukima = cell.column
        ElseIf cell.value = "���肷����" Then
            colSokuteiSukima = cell.column
        ElseIf cell.value = "�V������" Then
            colTenchoNikui = cell.column
        End If
    Next cell
    
    ' �K�v�ȗ񂪌������������m�F
    If colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "�K�v�ȗ񂪌�����܂���B�w�b�_�[���m�F���Ă��������B", vbCritical
        Exit Sub
    End If
    
    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colTenchoSukima).End(xlUp).row
    
    ' "�V��������(N)" �̒l�� "���肷����" �ɃR�s�[���A�l���v�Z
    For i = 2 To lastRowHelSpec
        ' �e�Z���̒l���擾
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "�V��������(N)"�̒l��"���肷����"�ɃR�s�[
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "�V��������(N)"�̒l����"�V������"�̒l������
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q���R���"���i"�̒l����
        wsHelSpec.Cells(i, 17).value = "���i" ' Q���17�Ԗڂ̗�
        wsHelSpec.Cells(i, 18).value = "���i" ' R���18�Ԗڂ̗�
    Next i
End Sub



Sub TransferValues()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    
    ' �V�[�g���Z�b�g
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' �w�b�_�[�̗�ԍ����擾
    colHinban = GetColumnIndex(wsHelSpec, "�i��(D)")
    colTencho = GetColumnIndex(wsHelSpec, "�V������")
    colBoutai = GetColumnIndex(wsSetting, "�X��No.")
    
    ' �K�v�ȗ񂪌������������m�F
    If colHinban = 0 Or colTencho = 0 Or colBoutai = 0 Then
        MsgBox "�K�v�ȗ񂪌�����܂���B�w�b�_�[���m�F���Ă��������B", vbCritical
        Exit Sub
    End If
    
    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "�i��(D)" ��̒l��T�����A�]�L
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
End Sub

Sub CopyAndSubtractValues()
    Dim wsHelSpec As Worksheet
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim i As Long
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim cell As Range
    
    ' �V�[�g���Z�b�g
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' �w�b�_�[�̗�ԍ����擾
    colTenchoSukima = GetColumnIndex(wsHelSpec, "�V��������(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "���肷����")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "�V������")
    
    ' �K�v�ȗ񂪌������������m�F
    If colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "�K�v�ȗ񂪌�����܂���B�w�b�_�[���m�F���Ă��������B", vbCritical
        Exit Sub
    End If
    
    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colTenchoSukima).End(xlUp).row
    
    ' "�V��������(N)" �̒l�� "���肷����" �ɃR�s�[���A�l���v�Z
    For i = 2 To lastRowHelSpec
        ' �e�Z���̒l���擾
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "�V��������(N)"�̒l��"���肷����"�ɃR�s�[
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "�V��������(N)"�̒l����"�V������"�̒l������
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q���R���"���i"�̒l����
        wsHelSpec.Cells(i, 17).value = "���i" ' Q���17�Ԗڂ̗�
        wsHelSpec.Cells(i, 18).value = "���i" ' R���18�Ԗڂ̗�
    Next i
End Sub




Sub UpdateCrownClearance()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim i As Long
    
    ' �V�[�g���Z�b�g
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' �w�b�_�[�̗�ԍ����擾
    colHinban = GetColumnIndex(wsHelSpec, "�i��(D)")
    colBoutai = GetColumnIndex(wsSetting, "�X��No.")
    colTencho = GetColumnIndex(wsHelSpec, "�V������")
    colTenchoSukima = GetColumnIndex(wsHelSpec, "�V��������(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "���肷����")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "�V������")
    
    ' �K�v�ȗ񂪌������������m�F
    If colHinban = 0 Or colBoutai = 0 Or colTencho = 0 Or colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "�K�v�ȗ񂪌�����܂���B�w�b�_�[���m�F���Ă��������B", vbCritical
        Exit Sub
    End If
    
    ' �ŏI�s���擾
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "�i��(D)" ��̒l��T�����A�]�L
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
    
    ' "�V��������(N)" �̒l�� "���肷����" �ɃR�s�[���A�l���v�Z
    For i = 2 To lastRowHelSpec
        ' �e�Z���̒l���擾
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "�V��������(N)"�̒l��"���肷����"�ɃR�s�[
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "�V��������(N)"�̒l����"�V������"�̒l������
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q���R���"���i"�̒l����
        wsHelSpec.Cells(i, 17).value = "���i" ' Q���17�Ԗڂ̗�
        wsHelSpec.Cells(i, 18).value = "���i" ' R���18�Ԗڂ̗�
    Next i
End Sub

Function GetColumnIndex(sheet As Worksheet, headerName As String) As Integer
    Dim cell As Range
    For Each cell In sheet.Rows(1).Cells
        If cell.value = headerName Then
            GetColumnIndex = cell.column
            Exit Function
        End If
    Next cell
    GetColumnIndex = 0 ' ������Ȃ��ꍇ��0��Ԃ�
End Function


