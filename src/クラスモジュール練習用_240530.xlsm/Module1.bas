Attribute VB_Name = "Module1"

Sub CopyAndPopulateSheet( _
    sourceSheetName As String, _
    prefix As String, _
    index As Integer, _
    customPropertyName As String, _
    customPropertyValue As String, _
    dataCollection As Collection, _
    writeMethod As String)

    Dim sheetName As String
    Dim ws As Worksheet
    Dim DataSetManager As New DataSetManager

    ' �V�[�g���𐶐�
    sheetName = GenerateSheetName(prefix, index)
    Debug.Print "Generated sheet name: " & sheetName

    ' �V�[�g�̑��݊m�F�ƍ쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' �R�s�[����\�[�X�V�[�g�����݂��邩�m�F
        If Not SheetExists(sourceSheetName) Then
            Debug.Print "Source sheet not found: " & sourceSheetName
            Exit Sub
        End If
        
        On Error Resume Next
        Sheets(sourceSheetName).Copy After:=Sheets(Sheets.Count)
        Set ws = ActiveSheet
        ws.Name = sheetName
        On Error GoTo 0
        
        ' �V�[�g���̕ύX�������������m�F
        If ws.Name <> sheetName Then
            Debug.Print "Failed to rename the sheet correctly."
            Exit Sub
        End If
    End If

    If Not ws Is Nothing Then
        ' �J�X�^���v���p�e�B�̐ݒ�
        ws.CustomProperties.Add Name:=customPropertyName, Value:=customPropertyValue

        ' �f�[�^�̓]�L
        Select Case writeMethod
            Case "WriteSelectedValuesToOutputSheet"
                DataSetManager.WriteSelectedValuesToOutputSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToRstlSheet"
                DataSetManager.WriteSelectedValuesToRstlSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToResultTempSheet"
                DataSetManager.WriteSelectedValuesToResultTempSheet ws.Name, dataCollection
            Case Else
                Debug.Print "Unknown write method: " & writeMethod
        End Select
    Else
        Debug.Print "Failed to create or find the sheet: " & sheetName
    End If

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim tmpSheet As Worksheet
    On Error Resume Next
    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not tmpSheet Is Nothing
    On Error GoTo 0
End Function

Function GenerateSheetName(prefix As String, index As Integer) As String
    GenerateSheetName = prefix & Format(index, "00")
End Function





' ���s����v���V�[�W���F2024-05-30�F�V�[�g�̃R�s�[�͂ł��邪�]�@�ł��ĂȂ��B
Sub TestSheetCreationAndDataWriting()
    Dim sheetIndex As Integer
    Dim i As Integer '�����Œǉ�
    Dim resultTempIndex As Integer
    Dim testValues As Collection
    Dim Record As Record
    Dim resultTempValues As Collection
    Dim outputValues As Collection
    Dim rstlValues As Collection
    
    ' DataSetManager�̏�����
    Dim DataSetManager As DataSetManager
    Set DataSetManager = New DataSetManager
    DataSetManager.Init
    Debug.Print "records initialized"
    
    ' �e�X�g�f�[�^�̏���
    Set testValues = New Collection
    Set resultTempValues = New Collection
    Set outputValues = New Collection
    Set rstlValues = New Collection
    
    ' �T���v�����R�[�h�̒ǉ��ƃt�B���^�����O
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-�V", "110F", "�V��", DateValue("2024/5/17"), 29, 3.07
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
        outputValues.Add Record
        rstlValues.Add Record
    End If
    
    Set Record = New Record
    Record.Initialize "02-110-Cold-�V", "110", "�V��", DateValue("2024/5/17"), 26, 4.91
    Record.Initialize "03-F110F-Wet-�V", "110F", "�V��", DateValue("2024/5/17"), 26, 2.89
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-�O", "110F", "�O����", DateValue("2024/5/17"), 26, 5.25
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-�O", "110F", "�O����", DateValue("2024/5/17"), 29, 5.64
    Else
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Else
    Record.Initialize "01-F110F-Hot-��", "110F", "�㓪��", DateValue("2024/5/17"), 26, 5.12
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
        rstlValues.Add Record
    End If
    
    ' �C���f�b�N�X������
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-��", "110F", "�㓪��", DateValue("2024/5/17"), 29, 5.19
    testValues.Add Record
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-��", "110F", "�㓪��", DateValue("2024/5/17"), 26, 5.12
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-��", "110F", "�㓪��", DateValue("2024/5/17"), 29, 5.19
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
    Else
        outputValues.Add Record
    sheetIndex = 1
    resultTempIndex = 1
    
    ' OutputSingle/OutputSheet �V�[�g�̍쐬�ƃf�[�^�̏�������
    For i = 1 To 5
        CopyAndPopulateSheet "�\��_��", "�\��_��_", sheetIndex, "Temp_Shinsei", "�\��_��", outputValues, "WriteSelectedValuesToOutputSheet"
        CopyAndPopulateSheet "�\��_�ė�", "�\��_�ė�_", sheetIndex, "Temp_Shinsei", "�\��_�ė�", outputValues, "WriteSelectedValuesToOutputSheet"
        sheetIndex = sheetIndex + 1
    Next i
    
    ' Rstl_Single/Rstl_Triple �V�[�g�̍쐬�ƃf�[�^�̏�������
    For i = 1 To 5
        CopyAndPopulateSheet "���_��", "���_��_", sheetIndex, "Temp_Teiki", "���_��", rstlValues, "WriteSelectedValuesToRstlSheet"
        CopyAndPopulateSheet "���_�ė�", "���_�ė�_", sheetIndex, "Temp_Teiki", "���_�ė�", rstlValues, "WriteSelectedValuesToRstlSheet"
        sheetIndex = sheetIndex + 1
    Next i
    
    ' Result_Temp�V�[�g�̍쐬�ƃf�[�^�̏�������
    CopyAndPopulateSheet "�˗�����", "�˗�����_", resultTempIndex, "Temp_Irai", "�˗�����", resultTempValues, "WriteSelectedValuesToResultTempSheet"
End Sub


