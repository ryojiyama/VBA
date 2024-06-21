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

    Dim groupedRecords As Object
    Set groupedRecords = CreateObject("Scripting.Dictionary")
    
    Dim sheetNames As Object
    Set sheetNames = CreateObject("Scripting.Dictionary")
    
    Dim sheetRecords As Object
    Set sheetRecords = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim record As New record
        record.LoadData ws, i

        testValues.Add record

        Call ClassifyKeys(record.sheetType, record.groupID, sheetNames, sheetRecords, record)

        If Not groupedRecords.exists(record.sheetType) Then
            groupedRecords.Add record.sheetType, New Collection
        End If
        
        Call AddRecordToGroup(groupedRecords(record.sheetType), record)
    Next i
    
'        If Not sheetRecords(record.sheetType).exists(record.sheetType) Then
'            sheetRecords(record.sheetType).Add record.sheetType, New Collection
'        End If
'        sheetRecords(record.sheetType)(record.sheetType).Add record

'        Call AddRecordToGroup(groupedRecords(record.sheetType), record)
'        Dim j As Integer
'        For j = 1 To groupedRecords(record.sheetType).Count
'            Dim addedRecord As record
'            Set addedRecord = groupedRecords(record.sheetType)(j)
'            'Debug.Print "Record in group: ID=" & addedRecord.sampleID & " SheetType=" & addedRecord.sheetType & " GroupID=" & addedRecord.groupID & " SampleColor=" & addedRecord.sampleColor
'        Next j
'    Next i
    If Not groupedRecords Is Nothing Then
        For Each key In groupedRecords.keys
            Debug.Print "key: " & key & ", count:"; groupedRecords(key).Count
        Next key
    Else
        Debug.Print "groupedRecords is not initalized or empty."
    End If
       
    Call PrintGroupedRecords(groupedRecords, sheetNames, sheetRecords)
    Debug.Print "Total unique records: " & testValues.Count
End Sub
Sub ClassifyKeys(sheetType As String, groupID As String, ByRef sheetNames As Object, ByRef sheetRecords As Object, ByRef records As record)

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
    
    '�V�[�g����ێ�����
    If sheetNames Is Nothing Then Set sheetNames = CreateObject("Scripting.Dictionary")
    If Not sheetNames.exists(sheetType) Then
        sheetNames.Add sheetType, CreateObject("Scripting.Dictionary")
    End If
    
    If sheetRecords Is Nothing Then Set sheetRecords = CreateObject("Scripting.Dictionary")
    If Not sheetRecords.exists(sheetType) Then
        sheetRecords.Add sheetType, CreateObject("Scripting.Dictionary")
    End If
    
    ' �V�[�g����ǉ�
    If Not sheetNames(sheetType).exists(baseTemplateName) Then
        sheetNames(sheetType).Add baseTemplateName, True
    End If
    If additionalTemplateName <> "" And Not sheetNames(sheetType).exists(additionalTemplateName) Then
        sheetNames(sheetType).Add additionalTemplateName, True
    End If
    
'    ' �V�[�g����ǉ�
'    If Not sheetNames(sheetType).Contains(baseTemplateName) Then
'        sheetNames(sheetType).Add baseTemplateName
'    End If
'    If additionalTemplateName <> "" And Not sheetNames(sheetType).Contains(additionalTemplateName) Then
'        sheetNames(sheetType).Add additionalTemplateName
'    End If

    ' ��{�e���v���[�g�ƒǉ��e���v���[�g�̃V�[�g����
    Call ProcessTemplateSheet(baseTemplateName, sheetType, groupID, sheetTypeIndex, sheetRecords, sheetNames, records)
    If additionalTemplateName <> "" Then
        Call ProcessTemplateSheet(additionalTemplateName, sheetType, groupID, sheetTypeIndex, sheetRecords, sheetNames, records)
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object, ByRef sheetRecords As Object, ByRef sheetNames As Object, ByRef record As record)
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID
    ' �V�[�g���̌���
    If Not sheetTypeIndex.exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = combinedKey
        'Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    'Debug.Print "newSheetName: " & newSheetName
    
    ' �V�[�g�̑��݊m�F�Ǝ擾/�쐬
    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        If templateName <> "" Then
'            On Error GoTo ErrorHandler
            Worksheets(templateName).Copy After:=Worksheets(Worksheets.Count)
            Set newSheet = Worksheets(Worksheets.Count)
            newSheet.name = newSheetName
            ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).name = "Temp_" & newSheetName
            'Debug.Print "Copied sheet from template: " & templateName & "to new sheet: "; newSheet.name
'            GoTo ExitSub
        Else
            Debug.Print "No template found for templateName: " & templateName
        End If
    Else
        Set newSheet = Worksheets(newSheetName)
    End If
    'Debug.Print "Record added to sheet:" & newSheet.name & "for groupID:"; groupID
    
    ' ���R�[�h���V�[�g�ɒǉ����鏈���i�K�v�ɉ����Ēǉ��j
    ' ��FnewSheet.Cells(�s, ��).Value = �f�[�^
    
    ' �V�[�g���ƃ��R�[�hID���֘A�t����
    If Not sheetRecords(sheetType).exists(newSheetName) Then
        sheetNames(sheetType).Add newSheetName, True
    End If
    ' sheetNames�����ɐV�����V�[�g����ۑ�
    If Not sheetRecords(sheetType).exists(newSheetName) Then
        sheetRecords(sheetType).Add newSheetName, New Collection
    End If
    sheetRecords(sheetType)(newSheetName).Add record
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
Private Sub AddRecordToGroup(groupCollection As Collection, ByVal record As record)
    ' �V�����C���X�^���X���쐬���Ă���ǉ�
    Dim newRecord As New record
    newRecord.ID = record.ID
    newRecord.sampleID = record.sampleID
    newRecord.itemNum = record.itemNum
    newRecord.testPart = record.testPart
    newRecord.testDate = record.testDate
    newRecord.testTemp = record.testTemp
    newRecord.maxValue = record.maxValue
    newRecord.timeOfMax = record.timeOfMax
    newRecord.duration49kN = record.duration49kN
    newRecord.duration73kN = record.duration73kN
    newRecord.preProcess = record.preProcess
    newRecord.sampleWeight = record.sampleWeight
    newRecord.sampleTop = record.sampleTop
    newRecord.sampleColor = record.sampleColor
    newRecord.sampleLotNum = record.sampleLotNum
    newRecord.sampleHelLot = record.sampleHelLot
    newRecord.sampleBandLot = record.sampleBandLot
    newRecord.structureResult = record.structureResult
    newRecord.penetrationResult = record.penetrationResult
    newRecord.testSection = record.testSection
    newRecord.groupID = record.groupID
    newRecord.sheetType = record.sheetType

    ' �O���[�v�Ƀ��R�[�h��ǉ�
    groupCollection.Add newRecord

    ' �f�o�b�O�o��
    'Debug.Print "Adding record: ID=" & newRecord.sampleID & ", GroupID=" & newRecord.groupID & ", SheetType=" & newRecord.sheetType
End Sub


' -----------------------------------------------------------------------------------------------------
Sub PrintGroupedRecords(ByRef groupedRecords As Object, ByRef sheetNames As Object, ByRef sheetRecords As Object)
    Dim dictkey As Variant
    Dim sheetNameKey As Variant
    Dim record As record
    
    ' groupedRecords�̊esheetType�����[�v
    For Each dictkey In groupedRecords.keys
        Debug.Print "Sheet Type: " & dictkey & ", Number of Records: " & groupedRecords(dictkey).Count
        
        ' groupedRecords(dictkey)�������������I�u�W�F�N�g��Ԃ����Ƃ��m�F
        Dim sheetsDict As Object
        Set sheetsDict = groupedRecords(dictkey)
        
        ' �e�V�[�g�����o��
        If sheetNames.exists(dictkey) Then
            For Each sheetNameKey In sheetNames(dictkey).keys
                Debug.Print " sheet Name: " & sheetNameKey
                Debug.Print "TypeName:"; TypeName(groupedRecords(dictkey))
                ' sheetNameKey�ł̃��R�[�h�R���N�V�������擾
                Dim recordsCollection As Collection
                If groupedRecords(dictkey).exists(sheetNameKey) Then
                    Set recordsCollection = groupedRecords(dictkey)(sheetNameKey)
                    For Each record In recordsCollection
                        Debug.Print " Record ID :" & record.sampleID
                    Next record
                Else
                    Debug.Print " No records from for sheet:" & sheetNameKey
                End If
            Next sheetNameKey
        End If
    Next dictkey
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
    If Not sheetTypeIndex.exists(combinedKey) Then
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
    Dim record As Variant
    For Each record In testValues
        ' �K�v�ȕϐ����擾
        Dim position As String
        Dim number As String
        Dim condition As String
        Dim recordType As String

        position = record.testPart  ' Location�v���p�e�B���g�p
        number = record.ID  ' ID�v���p�e�B���g�p
        condition = record.Temperature  ' Temperature�v���p�e�B���g�p
        recordType = "Single"  ' �Œ�l�i�K�؂ȃv���p�e�B���Ȃ����߁j

        ' Record�I�u�W�F�N�g�̊e�v���p�e�B�����݂��邩�`�F�b�N
        On Error Resume Next
        Debug.Print "Checking properties for Record:"
        Debug.Print "  ID: " & record.ID
        Debug.Print "  Location: " & record.testPart
        Debug.Print "  Temperature: " & record.Temperature
        Debug.Print "  DateValue: " & record.DateValue
        Debug.Print "  TemperatureValue: " & record.TemperatureValue
        Debug.Print "  Force: " & record.Force
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
            If Not singleGroups.exists(groupKey) Then
                singleGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = singleGroups(groupKey)
            AddToGroup tempDict, position, record
        ElseIf recordType = "Multi" Then
            If Not multiGroups.exists(groupKey) Then
                multiGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = multiGroups(groupKey)
            AddToGroup tempDict, position, record
        End If
    Next record
End Sub

Sub AddToGroup(ByVal group As Scripting.Dictionary, ByVal position As String, ByVal record As record)
    If Not group.exists(position) Then
        group.Add position, New Collection
    End If
    group(position).Add record
End Sub

Sub PrintGroups(ByVal groups As Scripting.Dictionary)
    Dim groupKey As Variant
    For Each groupKey In groups.keys
        Debug.Print "Group " & groupKey & ":"
        Dim position As Variant
        For Each position In groups(groupKey).keys
            Debug.Print "  " & position & ":"
            Dim record As Variant
            For Each record In groups(groupKey)(position)
                Debug.Print "    ID=" & record.ID & ", Location=" & record.testPart  ' �eRecord��ID��Location���o��
            Next record
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

Sub ProcessTemplateSheet202406171100(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object)
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID
    Debug.Print "combinedKey:" & combinedKey
    ' �V�[�g���̌���
    If Not sheetTypeIndex.exists(combinedKey) Then
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
'            On Error GoTo ErrorHandler
            Worksheets(templateName).Copy After:=Worksheets(Worksheets.Count)
            Set newSheet = Worksheets(Worksheets.Count)
            newSheet.name = newSheetName
            ThisWorkbook.VBProject.VBComponents(newSheet.CodeName).name = "Temp_" & newSheetName
            Debug.Print "Copied sheet from template: " & templateName & "to new sheet: "; newSheet.name
'            GoTo ExitSub
        Else
            Debug.Print "No template found for templateName: " & templateName
        End If
    Else
        Set newSheet = Worksheets(newSheetName)
    End If
    Debug.Print "Record added to sheet:" & newSheet.name & "for groupID:"
    
'ErrorHandler:
'    Debug.Print "Error " & Err.number & ": " & Err.Description
'    Resume Next
'ExitSub:
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
