Attribute VB_Name = "Utliteis"
'CopiedSheetNames�ŋL����Ă���V�[�g���폜����B
Sub DeleteCopiedSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "CopiedSheetNames�V�[�g��������܂���B"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim i As Long
    Application.DisplayAlerts = False
    For i = 1 To lastRow
        Dim sheetName As String
        sheetName = ws.Cells(i, 1).value
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    ClearCopiedSheetNames
End Sub
'CopiedSheetNames���N���A����B
Sub ClearCopiedSheetNames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CopiedSheetNames")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Cells.ClearContents
    End If
End Sub
' "LOG_Helmet��̃O���t���폜����
Public Sub DeleteAllChartsOnSheetsContainingName()

    Dim ws As Worksheet
    Dim chartObj As ChartObject

    ' ���[�N�u�b�N���̂��ׂẴV�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g����"500S_"���܂ޏꍇ
        If InStr(ws.Name, "500S_") > 0 Then
            ' �V�[�g��̂��ׂẴO���t�I�u�W�F�N�g�����[�v
            For Each chartObj In ws.ChartObjects
                chartObj.Delete
            Next chartObj
        End If
    Next ws

End Sub

Sub PrintMatchingSheetsFirstPage_SUb()
    Dim ws As Worksheet
    Dim copiedSheetNames As Worksheet
    Dim sheetName As String
    Dim cell As Range
    Dim foundSheet As Worksheet
    
    ' CopiedSheetNames�V�[�g��ݒ�
    Set copiedSheetNames = ThisWorkbook.Sheets("CopiedSheetNames")
    
    ' A��̒l�����[�v
    For Each cell In copiedSheetNames.Range("A1:A" & copiedSheetNames.Cells(copiedSheetNames.Rows.count, "A").End(xlUp).Row)
        sheetName = cell.value
        
        ' ��v����V�[�g������
        On Error Resume Next
        Set foundSheet = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' �V�[�g�����݂���ꍇ�A1�y�[�W�ڂ����
        If Not foundSheet Is Nothing Then
            With foundSheet
                ' ����̈��ݒ�
                .PageSetup.PrintArea = ""
                ' �V�[�g��1�y�[�W�ڂ݈̂��
                .PrintOut Preview:=False
            End With
            ' foundSheet���N���A
            Set foundSheet = Nothing
        End If
    Next cell
End Sub
Sub PrintFirstPageOfUniqueListedSheets()
    ' �w�肳�ꂽ�����[��1�y�[�W�ڂ��A�d���Ȃ�1�񂸂������v���V�[�W��
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim printedSheets As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String

    ' CopiedSheetNames �V�[�g��ݒ�
    Set wsSource = ThisWorkbook.Sheets("CopiedSheetNames")
    Set printedSheets = New Collection ' ������ꂽ�V�[�g����ǐՂ���R���N�V����

    ' A��̍ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    ' A��̒l�����[�v
    For i = 1 To lastRow
        sheetName = wsSource.Cells(i, 1).value

        On Error Resume Next
        ' �R���N�V�����ɓ������O�����ɑ��݂��邩�`�F�b�N
        printedSheets.Add sheetName, sheetName
        If Err.Number = 0 Then ' �ǉ������������ꍇ�A�V�[�g�͂܂��������Ă��Ȃ�
            Set wsTarget = ThisWorkbook.Sheets(sheetName)
            If Not wsTarget Is Nothing Then
                wsTarget.PrintOut From:=1, To:=1 ' �V�[�g��1�y�[�W�ڂ݂̂����
            End If
        End If
        On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g

        Set wsTarget = Nothing
    Next i
End Sub

' �E�N���b�N�J�X�^�����j���[�F�O���t��Y���̒l����
Sub UniformizeLineGraphAxes()
    On Error GoTo ErrorHandler
    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Check if there are any charts in the current sheet
        If ws.ChartObjects.count > 0 Then
            ' Loop through all the charts in the current sheet
            Dim chartObj As ChartObject
            For Each chartObj In ws.ChartObjects
                ' Split the chart name using "-"
                Dim parts() As String
                parts = Split(chartObj.Name, "-")
                
                ' Check the third part of the name
                If UBound(parts) >= 2 Then
                    With chartObj.chart.Axes(xlValue)
                        If parts(2) = "�V" Then
                            .MaximumScale = 5
                            .MajorUnit = 1# ' 1.0����
                        ElseIf parts(2) = "�O" Or parts(2) = "��" Or parts(2) = "����" Then
                            .MaximumScale = 10
                            .MajorUnit = 2# ' 2.0����
                        End If
                    End With
                End If
            Next chartObj
        End If
    Next ws

    MsgBox "���ׂẴV�[�g�̃O���t��Y���̍ő�l�Ɩڐ���P�ʂ�ݒ肵�܂����B", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical

End Sub

' LOG_Helmet�V�[�g�̃A�C�R���������B
Sub DeleteIconsKeepCharts()
    Dim ws As Worksheet
    Dim shp As Shape

    ' LOG_Helmet�V�[�g���w��
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' �V�[�g���̂��ׂẴV�F�C�v�����[�v����
    For Each shp In ws.Shapes
        ' �V�F�C�v���O���t�I�u�W�F�N�g�łȂ��ꍇ�A�폜
        If shp.Type <> msoChart Then
            shp.Delete
        End If
    Next shp
End Sub

' Setting��"B2"�Z���Ƀt�H�[�J�X
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    If SheetExists("Setting") Then
        Application.GoTo ActiveWorkbook.Sheets("Setting").Range("B2")
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
