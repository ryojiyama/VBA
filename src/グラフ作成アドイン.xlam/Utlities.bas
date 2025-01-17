Attribute VB_Name = "Utlities"

Option Explicit

'*******************************************************************************
' �萔��`
'*******************************************************************************
Private Const FIRST_DATA_ROW As Long = 2
Private Const FIRST_DATA_COL As String = "B"

'�w�b�_�[���̒�`
Private Type SheetHeaders
    sheetName As String
    headers() As String
End Type

'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�F�w�肳�ꂽLOG�V�[�g�̃f�[�^���N���A���w�b�_�[��ݒ�
' �����F�Ȃ�
'*******************************************************************************
Public Sub ResetLogSheets()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim targetSheets As Variant
    targetSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    
    Dim i As Long
    Dim sheetName As String
    Dim missingSheets As String
    Dim processedSheets As String
    Dim processedCount As Long
    Dim ws As Worksheet
    
    processedCount = 0
    
    ' �e�V�[�g������
    For i = LBound(targetSheets) To UBound(targetSheets)
        sheetName = CStr(targetSheets(i))
        
        ' �V�[�g�̑��݊m�F�Ǝ擾����x�ɍs��
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo ErrorHandler
        
        If Not ws Is Nothing Then
            ' �V�[�g�����݂���ꍇ�͏��������s
            ClearSheetData ws
            SetSheetHeaders ws
            
            ' �����ς݃V�[�g���L�^
            If processedSheets = "" Then
                processedSheets = sheetName
            Else
                processedSheets = processedSheets & ", " & sheetName
            End If
            processedCount = processedCount + 1
        Else
            ' ���݂��Ȃ��V�[�g���L�^
            If missingSheets = "" Then
                missingSheets = sheetName
            Else
                missingSheets = missingSheets & ", " & sheetName
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True

    ' ���ʕ�
    If processedCount = 0 Then
        Debug.Print "�ȉ��̃V�[�g��������܂���: " & missingSheets
    ElseIf missingSheets <> "" Then
        Debug.Print "��������: " & processedSheets
        Debug.Print "�������i�V�[�g�����j: " & missingSheets
    Else
        Debug.Print "�S�V�[�g��������"
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "�G���[����: " & Err.Description & " (�G���[�ԍ�: " & Err.Number & ")"
End Sub

' �V�[�g�̑��݂��`�F�b�N����֐�
Private Function sheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    sheetExists = Not ws Is Nothing
End Function

'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�w��V�[�g�̃f�[�^���N���A�iB��ȍ~�̃f�[�^�̂݁j
' �����Fws - �N���A�Ώۂ̃V�[�g��
' �O��FFIRST_DATA_ROW �͒萔�Ƃ��Ē�`�ς�
'*******************************************************************************
Private Sub ClearSheetData(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Debug.Print "�����J�n - �V�[�g��: " & ws.Name
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim clearRange As Range
    
    With ws
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        lastCol = .Cells(1, .columns.Count).End(xlToLeft).column
        
        If lastRow >= FIRST_DATA_ROW And lastCol >= 2 Then
            Set clearRange = .Range(.Cells(FIRST_DATA_ROW, "B"), .Cells(lastRow, lastCol))
            With clearRange
                .ClearContents
                .Interior.colorIndex = xlNone
                .Borders.LineStyle = xlNone
            End With
            Debug.Print "�f�[�^�N���A���� - �s��: " & (lastRow - FIRST_DATA_ROW + 1)
        Else
            Debug.Print "�N���A�Ώۃf�[�^�Ȃ�"
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "�G���[���� - �V�[�g '" & ws.Name & "': " & Err.Description
End Sub

'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�V�[�g�̃w�b�_�[��ݒ�
' �����Fws - �w�b�_�[��ݒ肷��V�[�g��
' �ˑ��FGetSheetHeaders�֐��ɂ��w�b�_�[���̎擾
'*******************************************************************************
Private Sub SetSheetHeaders(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' �w�b�_�[���̎擾
    Dim headers As Variant
    headers = GetSheetHeaders(ws.Name)
    
    If Not IsEmpty(headers) Then
        Debug.Print "�w�b�_�[�ݒ�J�n - �V�[�g��: " & ws.Name
        
        ' A�񂩂�Z��܂ł̃w�b�_�[���N���A
        ws.Range("A1:Z1").ClearContents
        
        ' �e��Ƀw�b�_�[��ݒ�
        Dim headerRange As Range
        Set headerRange = ws.Range("A1:Z1")
        
        ' A��͌Œ�� "����"
        ws.Range("A1").value = ""
        
        ' B��ȍ~�ɔz��̓��e��ݒ�
        Dim i As Long
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 2).value = headers(i)
        Next i
        
        ' �w�b�_�[�s�̏����ݒ�
        With headerRange
            .Font.Name = "���S�V�b�N"
            .Font.size = 12
            .Font.Bold = True
            .Font.Color = RGB(217, 217, 217)
            .Interior.Color = RGB(48, 84, 150)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(190, 190, 190)
            End With
            
        ' �s�̍�����ݒ�
        headerRange.EntireRow.RowHeight = 20
        ' �񕝂�ݒ�i�V�����T�u���[�`�����Ăяo���j
        SetColumnWidths ws

        End With
    Else
        Debug.Print "�x�� - �w�b�_�[��񂪋�ł�: " & ws.Name
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "�G���[���� - �V�[�g '" & ws.Name & "': " & Err.Description
    Debug.Print "�G���[�ԍ�: " & Err.Number
End Sub

'*******************************************************************************
' �֐�
' �@�\�F�V�[�g�ʂ̃w�b�_�[�����擾
' �����Fws - �w�b�_�[���擾����V�[�g��
' �ߒl�F�w�b�_�[�z��A����`�̏ꍇ��Empty
'*******************************************************************************
Private Function GetSheetHeaders(ByVal sheetName As String) As Variant
    
    Select Case Trim(sheetName)
        Case "LOG_Helmet"
            GetSheetHeaders = Array("ID", "����ID", "�i��", "�������e", _
                "������", "���x", "�ő�l(kN)", "�ő�l�̎���", "4.9(ms)", "7.3(ms)", _
                "�O����", "�d��", "�V��������", "�X�̐F", "���b�gNo.", _
                "�X�̃��b�g", "�������b�g", "�\��_��������", "�ϊђ�_��������", "�����敪")
        
        Case "LOG_FallArrest"
            GetSheetHeaders = Empty
            
        Case "LOG_Bicycle"
            GetSheetHeaders = Array("ID", "����ID", "�i��", "���b�g�ԍ�", _
                "������", "���x", "���x", "�d��", "�ő�l(G)", "�ő�l�̎���", "G�̌p������", _
                "�O����", "�����ӏ�", "�X�̐F", "�X�̂̍ގ�", "�A���r��", "�l���͌^", "������̏��", _
                "�O�ό���", "�����Ђ�����", "�ޗ��E�t���i����")
            
        Case "LOG_BaseBall"
            GetSheetHeaders = Empty
            
        Case Else
            GetSheetHeaders = Empty
    End Select
End Function
'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�V�[�g�ʂ̗񕝂�ݒ�
' �����Fws - �񕝂�ݒ肷�郏�[�N�V�[�g
'*******************************************************************************
Private Sub SetColumnWidths(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Debug.Print "�񕝐ݒ�J�n - �V�[�g��: " & ws.Name
    
    With ws
        ' A��i���ԁj�͋���
        .columns("A").ColumnWidth = 6
        
        Select Case ws.Name
            Case "LOG_Helmet"
                ' ID��Ƃ��̑��̗�ŕ���ς���
                .columns("B").ColumnWidth = 20  ' ID
                .columns("C").ColumnWidth = 12  ' ����ID
                .Range("D:U").ColumnWidth = 15  ' ���̑��̗�
                
            Case "LOG_Bicycle"
                .columns("B").ColumnWidth = 20  ' ID
                .columns("C").ColumnWidth = 12  ' ����ID
                .Range("D:V").ColumnWidth = 10  ' ���̑��̗�
                
                ' ����̗�͕����L����
                .columns("S").ColumnWidth = 20  ' ������̏��
                .columns("T").ColumnWidth = 20  ' �T�ό���
                .columns("U").ColumnWidth = 20  ' �����Ђ�����
                .columns("V").ColumnWidth = 20  ' �ޗ��E�t���i����
        End Select
    End With
    
    Debug.Print "�񕝐ݒ芮�� - �V�[�g��: " & ws.Name
    Exit Sub
    
ErrorHandler:
    Debug.Print "�G���[���� - �񕝐ݒ蒆: " & ws.Name
    Debug.Print "  - �G���[���e: " & Err.Description
End Sub
'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�V�����w�b�_�[��ǉ����邽�߂̕⏕�v���V�[�W��
' �����FsheetName - �w�b�_�[��ǉ�����V�[�g��
'       headers - �ǉ�����w�b�_�[�z��
'*******************************************************************************
Public Sub AddNewHeaders(ByVal sheetName As String, ByRef headers As Variant)
    ' ���̊֐��͏����I�Ƀw�b�_�[��ǉ�����ۂɎg�p
    ' ������F
    ' Dim newHeaders As Variant
    ' newHeaders = Array("�V�����w�b�_�[1", "�V�����w�b�_�[2", ...)
    ' AddNewHeaders "LOG_FallArrest", newHeaders
End Sub



' DeleteAllChartsAndSheets_�V�[�g���̃O���t���폜����
Sub DeleteAllChartsAndSheets()
    Dim sheet As Worksheet
    Dim chart As ChartObject
    Dim sheetName As String

    ' �V�[�g�̃��X�g
    Dim sheetList() As Variant
    sheetList = Array( _
        "Setting", _
        "LOG_Helmet", _
        "LOG_BaseBall", _
        "LOG_Bicycle", _
        "LOG_FallArrest", _
        "���|�[�g�{��", _
        "���|�[�g�O���t", _
        "��������" _
    )
    Application.DisplayAlerts = False

    ' �e�V�[�g�ɑ΂��ď��������s
    For Each sheet In ActiveWorkbook.Sheets
        sheetName = sheet.Name
        ' �O���t�̍폜
        If IsInArray(sheetName, sheetList) Then
            For Each chart In sheet.ChartObjects
                chart.Delete
            Next chart
        End If
    Next sheet

    Application.DisplayAlerts = True
End Sub

' DeleteAllChartsAndSheets_�z����ɓ���̒l�����݂��邩�`�F�b�N����֐�
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function





' �e��ɏ����ݒ������
Public Sub CustomizeSheetFormats()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range

    ' Apply to the following sheets
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' Loop through each sheet
    For Each sheet In sheetNames
        Set ws = Worksheets(sheet)

        ' Loop through each cell in the first row
        For Each cell In ws.Rows(1).Cells
            If InStr(1, cell.value, "ID") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "����ID") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�i��") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�������e") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "������") > 0 Then ' Date
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToDate(rng)
            ElseIf InStr(1, cell.value, "���x") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "�ő�l(kN)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericFourDecimals(rng)
            ElseIf InStr(1, cell.value, "�ő�l�̎���(ms)") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "4.9kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "7.3kN") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumericTwoDecimals(rng)
            ElseIf InStr(1, cell.value, "�O����") > 0 Then ' String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�d��") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "�V��������") > 0 Then ' Numeric
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToNumeric(rng)
            ElseIf InStr(1, cell.value, "���i���b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�X�̃��b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�������b�g") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�\������") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�ϊђʌ���") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            ElseIf InStr(1, cell.value, "�����敪") > 0 Then 'String
                Set rng = ws.Range(cell.Offset(1, 0), ws.Cells(Rows.Count, cell.column).End(xlUp))
                Call ConvertToString(rng)
            End If
        Next cell
    Next sheet
End Sub

Sub ConvertToNumeric(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericTwoDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.00"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToNumericFourDecimals(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "0.0000"
    For Each cell In rng
        If IsNumeric(cell.value) Then
            cell.value = CDbl(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

Sub ConvertToString(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "@"
    For Each cell In rng
        cell.value = CStr(cell.value)
    Next cell
End Sub

Sub ConvertToDate(rng As Range)
    Dim cell As Range
    rng.NumberFormat = "yyyy/mm/dd"  ' ���t�̕\���`����ݒ�
    For Each cell In rng
        If IsDate(cell.value) Then
            cell.value = CDate(cell.value)
        Else
            cell.ClearContents
        End If
    Next cell
End Sub
' �󔒃Z����"-"��}��
Public Sub FillBlanksWithHyphenInMultipleSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim sheetName As Variant

    ' �ΏۃV�[�g�̖��O��z��ɐݒ�
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")

    ' �e�V�[�g�ɂ��ď������s��
    For Each sheetName In sheetNames
        On Error Resume Next
        ' �ΏۃV�[�g��ݒ�
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = Nothing ' ws�ϐ����N���A
            GoTo NextSheet ' ���̃V�[�g�ɐi��
        End If

        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        lastCol = ws.Cells(1, "Z").column ' Z��̗�ԍ���ݒ�

        ' 2�s�ڂ���ŏI�s�܂Ń��[�v�i1�s�ڂ̓w�b�_�[�Ɖ���j
        For i = 2 To lastRow
            For j = ws.Cells(i, "B").column To lastCol
                If IsEmpty(ws.Cells(i, j).value) Then
                    'Debug.Print "EmptyCell:" & "Cells&("; i; "," & j; ")"
                    ws.Cells(i, j).value = "-"
                End If
            Next j
        Next i

        ' �V�[�g�����̏I�����x��
NextSheet:
        ' ���̃V�[�g�̏����Ɉڂ�O�ɕϐ����N���A
        Set ws = Nothing
    Next sheetName
End Sub

