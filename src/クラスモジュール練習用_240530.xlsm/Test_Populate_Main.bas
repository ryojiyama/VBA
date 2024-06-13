Attribute VB_Name = "Test_Populate_Main"

    
    'Main
Sub TestSheetCreationAndDataWriting()
    Call ResetSheetTypeIndex   ' �C���f�b�N�X�����Z�b�g
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim i As Integer

    Dim testValues As New Collection
    Dim recordIDs As Object
    Set recordIDs = CreateObject("Scripting.Dictionary")

    Dim groupedRecords As Object
    Set groupedRecords = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim Record As New Record
        Record.LoadData ws, i

        Dim recordID As String
        recordID = Record.sampleID

        If Not recordIDs.Exists(recordID) Then
            recordIDs.Add recordID, Nothing
            testValues.Add Record
            
            'Debug.Print "SheetType:"; Record.sheetType
            If Not groupedRecords.Exists(Record.sheetType) Then
                groupedRecords.Add Record.sheetType, New Collection
            End If
            'Debug.Print "Calling ClassifyKeys for sheetType: " & Record.sheetType
            Call ClassifyKeys(Record.sheetType, Record.groupID)
            groupedRecords(Record.sheetType).Add Record
                    groupedRecords(Record.sheetType).Add Record
'            Debug.Print "Record loaded: ID=" & Record.sampleID & _
'                        " SheetType=" & Record.sheetType & _
'                        " GroupID=" & Record.GroupID
        Else
            Debug.Print "Duplicate record skipped: ID=" & Record.sampleID
        End If
    Next i

    Debug.Print "Total unique records: " & testValues.Count
End Sub

Sub ClassifyKeys(sheetType As String, groupID As String)
    ' ���R�[�h���ƂɃV�[�g�l�[�����쐬����
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")
    
    ' �O���[�vID���쐬
    groupID = Left(groupID, 2)
    
    Dim baseTemplateName As String
    Dim additionalTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "�\��_��"
            additionalTemplateName = "���_��"
        Case "Multi"
            baseTemplateName = "�\��_�ė�"
            additionalTemplateName = "���_�ė�"
        Case Else
            baseTemplateName = "���̑�"
            additionalTemplateName = ""
    End Select
    ' ��{�e���v���[�g�ƒǉ��e���v���[�g�̃V�[�g����
    Debug.Print "baseTemp:"; baseTemplateName
    Debug.Print "addiTemp:"; additionalTemplateName
    Call ProcessTemplateSheet(baseTemplateName, sheetType, groupID, sheetTypeIndex)
    If additionalTemplateName <> "" Then
        Call ProcessTemplateSheet(additionalTemplateName, sheetType, groupID, sheetTypeIndex)
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object)
    Debug.Print "Processing template:" & templateName
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID
    ' �V�[�g���̌���
    If Not sheetTypeIndex.Exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = combinedKey
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName
    
    ' �V�[�g�̑��݊m�F�Ǝ擾/�쐬
    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        If templateName <> "" Then
            Worksheets(templateName).Copy After:=Worksheets(Worksheets.Count)
            Set newSheet = Worksheets(Worksheets.Count)
            newSheet.name = newSheetName
            ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).name = "Temp_" & newSheetName
            Debug.Print "Copied sheet from template: " & templateName & "to new sheet: "; newSheet.name
        Else
            Debug.Print "No template found for templateName: " & templateName
        End If
    Else
        Set newSheet = Worksheets(newSheetName)
    End If
    Debug.Print "Record added to sheet:" & newSheet.name & "for groupID:"
'    If Not SheetExists(newSheetName) Then
'        Set newSheet = Worksheets.Add
'        newSheet.name = newSheetName
'        Debug.Print "Created new sheet: " & newSheetName
'    Else
'        Set newSheet = Worksheets(newSheetName)
'    End If

    ' �f�o�b�O�o�́F�����O���[�vID�������R�[�h�����������ނ���Ă��邩�m�F
    'Debug.Print "Record added to sheet: " & newSheet.name & " for groupID: " & groupID

    ' ���R�[�h���V�[�g�ɒǉ����鏈���i�K�v�ɉ����Ēǉ��j
    ' ��FnewSheet.Cells(�s, ��).Value = �f�[�^
End Sub





Function SheetExists(sheetName As String) As Boolean
    ' �V�[�g�̑��݃`�F�b�N
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function
Sub ResetSheetTypeIndex()
    ' �C���f�b�N�X�̃��Z�b�g
    Static sheetTypeIndex As Object
    Set sheetTypeIndex = Nothing ' Dictionary �I�u�W�F�N�g�̉��
    Set sheetTypeIndex = CreateObject("Scripting.Dictionary") ' �V���� Dictionary �I�u�W�F�N�g�̏�����
    Static groupSheetIndexes As Object
    Set groupSheetIndexes = Nothing ' Dictionary �I�u�W�F�N�g�̉��
    Set groupSheetIndexes = CreateObject("Scripting.Dictionary") ' �V���� Dictionary �I�u�W�F�N�g�̏�����
End Sub

Sub ClassifyKeys1600_�V�[�g�̐����܂ł��܂��s�����(sheetType As String, groupID As String)
    ' ���R�[�h���ƂɃV�[�g�l�[�����쐬����
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")
    
    ' �O���[�vID���쐬
    groupID = Left(groupID, 2)
    
    Dim baseTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "�\��_��"
        Case "Multi"
            baseTemplateName = "�\��_�ė�"
        Case Else
            baseTemplateName = "���̑�"
    End Select
    
    Dim combinedKey As String
    combinedKey = sheetType & "_" & groupID
    
    ' �V�[�g���̌���
    If Not sheetTypeIndex.Exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = baseTemplateName & "_" & groupID
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName
    
    ' �V�[�g�̑��݊m�F�Ǝ擾/�쐬
    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        ' �w�肳�ꂽ�����̃V�[�g���R�s�[���ĐV�����V�[�g���쐬����
        Select Case baseTemplateName
            Case "�\��_��", "�\��_�ė�", "���_��", "���_�ė�", "���ʎ���", "�˗�����", "LOG_Helmet", "DataSheet"
                Worksheets(baseTemplateName).Copy After:=Worksheets(Worksheets.Count)
                Set newSheet = Worksheets(Worksheets.Count)
                newSheet.name = newSheetName
                ' �I�u�W�F�N�g����"Temp_" & newSheetName�ɐݒ�
                 ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).name = "Temp_" & newSheetName
                Debug.Print "Copied sheet from template: " & baseTemplateName & " to new sheet: " & newSheet.name
            Case Else
                Debug.Print "No template found for baseTemplateName: " & baseTemplateName
        End Select
    Else
        Set newSheet = Worksheets(newSheetName)
    End If
'    If Not SheetExists(newSheetName) Then
'        Set newSheet = Worksheets.Add
'        newSheet.name = newSheetName
'        Debug.Print "Created new sheet: " & newSheetName
'    Else
'        Set newSheet = Worksheets(newSheetName)
'    End If

    ' �f�o�b�O�o�́F�����O���[�vID�������R�[�h�����������ނ���Ă��邩�m�F
    Debug.Print "Record added to sheet: " & newSheetName & " for groupID: " & groupID

    ' ���R�[�h���V�[�g�ɒǉ����鏈���i�K�v�ɉ����Ēǉ��j
    ' ��FnewSheet.Cells(�s, ��).Value = �f�[�^
End Sub

Sub GenerateSheets()
    Dim sheetNames As Collection
    Set sheetNames = New Collection
    
    ' ���O�ɒ�`���ꂽ�V�[�g���̃��X�g
    sheetNames.Add "LOG_Helmet"
    sheetNames.Add "DataSheet"
    sheetNames.Add "�\��_��"
    sheetNames.Add "�\��_�ė�"
    sheetNames.Add "���_��"
    sheetNames.Add "���_�ė�"
    sheetNames.Add "���ʎ���"
    sheetNames.Add "�˗�����"
    
    Dim i As Integer
    Dim sheetName As String
    For i = 1 To sheetNames.Count
        sheetName = sheetNames(i)
        Debug.Print "Checking sheet: " & sheetName
        If Not SheetExists(sheetName) Then
            Dim newSheet As Worksheet
            Set newSheet = Worksheets.Add
            newSheet.name = sheetName
            Debug.Print "Created new sheet: " & sheetName
        Else
            Debug.Print "Sheet already exists: " & sheetName
        End If
    Next i
End Sub

'Function SheetExists_1200(sheetName As String) As Boolean
'    Debug.Print "Checking if sheet exists: " & sheetName
'    Dim sheet As Worksheet
'    On Error Resume Next
'    Set sheet = Worksheets(sheetName)
'    On Error GoTo 0
'    SheetExists = Not sheet Is Nothing
'End Function




Sub ClassifyKeys_20240613(testValues As Collection, ByRef singleGroups As Scripting.Dictionary, ByRef multiGroups As Scripting.Dictionary)
    Dim Record As Variant
    For Each Record In testValues
        ' �K�v�ȕϐ����擾
        Dim position As String
        Dim number As String
        Dim condition As String
        Dim recordType As String

        position = Record.testPart  ' Location�v���p�e�B���g�p
        number = Record.ID  ' ID�v���p�e�B���g�p
        condition = Record.Temperature  ' Temperature�v���p�e�B���g�p
        recordType = "Single"  ' �Œ�l�i�K�؂ȃv���p�e�B���Ȃ����߁j

        ' Record�I�u�W�F�N�g�̊e�v���p�e�B�����݂��邩�`�F�b�N
        On Error Resume Next
        Debug.Print "Checking properties for Record:"
        Debug.Print "  ID: " & Record.ID
        Debug.Print "  Location: " & Record.testPart
        Debug.Print "  Temperature: " & Record.Temperature
        Debug.Print "  DateValue: " & Record.DateValue
        Debug.Print "  TemperatureValue: " & Record.TemperatureValue
        Debug.Print "  Force: " & Record.Force
        On Error GoTo 0

        ' �G���[����������v���p�e�B�����
        If Err.number <> 0 Then
            Debug.Print "Error accessing property: " & Err.Description
            Exit Sub
        End If

        ' �O���[�v�L�[�̐����ƕ��ޏ���
        Dim groupKey As String
        If position = "��" Then
            groupKey = number & "-" & condition & "-��"
        Else
            groupKey = number & "-" & condition
        End If

        Dim tempDict As Scripting.Dictionary
        If recordType = "Single" Then
            If Not singleGroups.Exists(groupKey) Then
                singleGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = singleGroups(groupKey)
            AddToGroup tempDict, position, Record
        ElseIf recordType = "Multi" Then
            If Not multiGroups.Exists(groupKey) Then
                multiGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = multiGroups(groupKey)
            AddToGroup tempDict, position, Record
        End If
    Next Record
End Sub

Sub AddToGroup(ByVal group As Scripting.Dictionary, ByVal position As String, ByVal Record As Record)
    If Not group.Exists(position) Then
        group.Add position, New Collection
    End If
    group(position).Add Record
End Sub

Sub PrintGroups(ByVal groups As Scripting.Dictionary)
    Dim groupKey As Variant
    For Each groupKey In groups.keys
        Debug.Print "Group " & groupKey & ":"
        Dim position As Variant
        For Each position In groups(groupKey).keys
            Debug.Print "  " & position & ":"
            Dim Record As Variant
            For Each Record In groups(groupKey)(position)
                Debug.Print "    ID=" & Record.ID & ", Location=" & Record.testPart  ' �eRecord��ID��Location���o��
            Next Record
        Next position
    Next groupKey
End Sub



Sub PopulateGroupedSheets(singleGroups As Scripting.Dictionary, multiGroups As Scripting.Dictionary)
    Dim sheetIndex As Integer
    Dim groupKey As Variant
    Dim newSheetName As String
    Dim templateNames As Collection
    Dim templateName As Variant

    ' �V�[�g�̍쐬���W�b�N
    sheetIndex = 1
    For Each groupKey In singleGroups.keys
        Set templateNames = New Collection

        If InStr(groupKey, "SingleValue") > 0 Then
            templateNames.Add "�\��_��"
            templateNames.Add "���_��"
        ElseIf InStr(groupKey, "SideValue") > 0 Then
            templateNames.Add "���ʎ���"
        Else
            templateNames.Add "�\��_�ė�"
            templateNames.Add "���_�ė�"
        End If

        For Each templateName In templateNames
            newSheetName = templateName & "_" & sheetIndex
            If Not SheetExists(newSheetName) Then
                Debug.Print "newSheetName:"; newSheetName
                'Call CopyAndPopulateSheet(templateName, newSheetName, singleGroups(groupKey))
                sheetIndex = sheetIndex + 1
            End If
        Next templateName
    Next groupKey

    For Each groupKey In multiGroups.keys
        Set templateNames = New Collection

        If InStr(groupKey, "SingleValue") > 0 Then
            templateNames.Add "�\��_��"
            templateNames.Add "���_��"
        ElseIf InStr(groupKey, "SideValue") > 0 Then
            templateNames.Add "���ʎ���"
        Else
            templateNames.Add "�\��_�ė�"
            templateNames.Add "���_�ė�"
        End If

        For Each templateName In templateNames
            newSheetName = templateName & "_" & sheetIndex
            If Not SheetExists(newSheetName) Then
                Debug.Print "newSheetName:"; newSheetName
                'Call CopyAndPopulateSheet(templateName, newSheetName, multiGroups(groupKey))
                sheetIndex = sheetIndex + 1
            End If
        Next templateName
    Next groupKey
End Sub


'Sub TestSheetCreationAndDataWriting1350()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("DataSheet")
'    Dim lastRow As Long
'    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
'    Dim i As Integer
'
'    Dim testValues As New Collection
'    Dim groupedRecords As Object
'    Set groupedRecords = CreateObject("Scripting.Dictionary")
'
'    For i = 2 To lastRow
'        Dim Record As New Record
'        Record.LoadData ws, i
'
'        ' Debug.Print��Record�̋�̓I�ȃv���p�e�B���w��
'        Debug.Print "Record loaded: ID=" & Record.ID & " Group=" & Record.testPart  ' Record�̋�̓I�ȃv���p�e�B���o��
'
'        testValues.Add Record
'
'        ' ���ނ��ꂽ�O���[�v�Ƀ��R�[�h��ǉ�
'        If Not groupedRecords.Exists(Record.testPart) Then
'            groupedRecords.Add Record.testPart, New Collection
'        End If
'        groupedRecords(Record.testPart).Add Record
'    Next i
'
'    ' �O���[�v�̓��e���m�F�i�f�o�b�O�p�j
'    Dim key As Variant
'    For Each key In groupedRecords
'        'Debug.Print "Main()_Group: " & key & ", Count: " & groupedRecords(key).Count
'    Next key
'
'    ' singleGroups��multiGroups��K�؂ɏ���������
'    Dim singleGroups As Scripting.Dictionary
'    Set singleGroups = CreateObject("Scripting.Dictionary")
'    Dim multiGroups As Scripting.Dictionary
'    Set multiGroups = CreateObject("Scripting.Dictionary")
'
'    ' �L�[�̕���
'    Debug.Print "testValues count: " & testValues.Count  ' testValues�̌������m�F
'
'    ' �^�m�F�̂��߂̃f�o�b�O�o��
'    Debug.Print "singleGroups type: " & TypeName(singleGroups)  ' singleGroups�̌^���m�F
'    Debug.Print "multiGroups type: " & TypeName(multiGroups)    ' multiGroups�̌^���m�F
'
'    Call ClassifyKeys(testValues, singleGroups, multiGroups)
'
'    ' ���ʂ̕\���i�K�v�ɉ����ăR�����g�A�E�g�j
'    Debug.Print "SingleValue Groups:"
'    Call PrintGroups(singleGroups)
'
'    Debug.Print "MultiValue Groups:"
'    Call PrintGroups(multiGroups)
'
'    ' �f�[�^�̃O���[�v���ƃV�[�g�������݂��s��
'    Call PopulateGroupedSheets(singleGroups, multiGroups)
'End Sub

'Function SheetExists_20240613(sheetName As String) As Boolean
'    ' PopulateGroupedSheets�̃T�u�v���V�[�W��
'    Dim tmpSheet As Worksheet
'    On Error Resume Next
'    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
'    SheetExists = Not tmpSheet Is Nothing
'    On Error GoTo 0
'End Function
