' Record �N���X���W���[��
Option Explicit

' Public�ϐ��̒�`
Public ID As String 'ID
Public sampleID As String '����ID
Public itemNum As String '�i��
Public testPart As String '�����ʒu
Public testDate As Date '������
Public testTemp As Double '���x
Public maxValue As Double '�ő�l
Public timeOfMax As Double '�ő�l���L�^��������
Public duration49kN As Double '4.9kN�̌p������
Public duration73kN As Double '7.3kN�̌p������
Public preProcess As String '�O����
Public sampleWeight As Double '�d��
Public sampleTop As Double '�V������
Public sampleColor As String '�X�̐F
Public sampleLotNum As String '���i���b�g
Public sampleHelLot As String '�X�̃��b�g
Public sampleBandLot As String '�������b�g
Public structureResult As String '�\������
Public penetrationResult As String '�ђʌ���
Public testSection As String '�����敪


Public Values As Collection
Public sheetType As String
Public groupID As String

' ���������\�b�h
Public Sub InitSingle()
    ' ����������
End Sub

Public Sub InitMultiple()
    Set Values = New Collection
End Sub

Public Function GetSpecificValues() As Collection
    Set GetSpecificValues = Values
End Function

Public Sub Initialize( _
    ByVal ID As String, _
    ByVal sampleID As String, _
    ByVal itemNum As String, _
    ByVal testPart As String, _
    ByVal testDate As Date, _
    ByVal testTemp As Double, _
    ByVal maxValue As Double, _
    ByVal timeOfMax As Double, _
    ByVal duration49kN As Double, _
    ByVal duration73kN As Double, _
    ByVal preProcess As String, _
    ByVal sampleWeight As Double, _
    ByVal sampleTop As Double, _
    ByVal sampleColor As String, _
    ByVal sampleLotNum As String, _
    ByVal sampleHelLot As String, _
    ByVal sampleBandLot As String, _
    ByVal structureResult As String, _
    ByVal penetrationResult As String, _
    ByVal testSection As String)

    Me.ID = ID
    Me.sampleID = sampleID
    Me.itemNum = itemNum
    Me.testPart = testPart
    Me.testDate = testDate
    Me.testTemp = testTemp
    Me.maxValue = maxValue
    Me.timeOfMax = timeOfMax
    Me.duration49kN = duration49kN
    Me.duration73kN = duration73kN
    Me.preProcess = preProcess
    Me.sampleWeight = sampleWeight
    Me.sampleTop = sampleTop
    Me.sampleColor = sampleColor
    Me.sampleLotNum = sampleLotNum
    Me.sampleHelLot = sampleHelLot
    Me.sampleBandLot = sampleBandLot
    Me.structureResult = structureResult
    Me.penetrationResult = penetrationResult
    Me.testSection = testSection
End Sub


' �f�[�^�����[�h���A���ނ���уO���[�v�����s�����\�b�h
Public Sub LoadData(ByVal ws As Worksheet, ByVal row As Integer)
    ID = ws.Cells(row, 2).Value
    sampleID = ws.Cells(row, 3).Value
    itemNum = ws.Cells(row, 4).Value
    testPart = ws.Cells(row, 5).Value
    testDate = ws.Cells(row, 6).Value
    testTemp = ws.Cells(row, 7).Value
    maxValue = ws.Cells(row, 8).Value
    timeOfMax = ws.Cells(row, 9).Value
    duration49kN = ws.Cells(row, 10).Value
    duration73kN = ws.Cells(row, 11).Value
    preProcess = ws.Cells(row, 12).Value
    sampleWeight = ws.Cells(row, 13).Value
    sampleTop = ws.Cells(row, 14).Value
    sampleColor = ws.Cells(row, 15).Value
    sampleLotNum = ws.Cells(row, 16).Value
    sampleHelLot = ws.Cells(row, 17).Value
    sampleBandLot = ws.Cells(row, 18).Value
    structureResult = ws.Cells(row, 19).Value
    penetrationResult = ws.Cells(row, 20).Value
    testSection = ws.Cells(row, 21).Value

    ' ID�𕪐͂��ăJ�e�S��������
    Dim parts() As String
    parts = Split(sampleID, "-")

    ' ���ԕ����ł̃J�e�S������
    If InStr(parts(1), "F") > 0 And InStr(parts(2), "��") > 0 Then
        sheetType = "Side"
    ElseIf InStr(parts(1), "F") > 0 Then
        sheetType = "Multi"
    Else
        sheetType = "Single"
    End If

    ' ���������ł̃O���[�v����
    Select Case parts(3)
        Case "��"
            groupID = parts(0) & "." & "SideValue." & parts(1) & "." & parts(3) & "." & parts(2)
        Case Else
            groupID = parts(0) & "." & parts(1) & "." & parts(3) & "." & parts(2)
    End Select
End Sub




Public Sub LoadData_ForDebug(ByVal ws As Worksheet, ByVal row As Integer)
    ID = ws.Cells(row, 2).Value
    sampleID = ws.Cells(row, 3).Value
    itemNum = ws.Cells(row, 4).Value
    testPart = ws.Cells(row, 5).Value
    testDate = ws.Cells(row, 6).Value
    testTemp = ws.Cells(row, 7).Value
    maxValue = ws.Cells(row, 8).Value
    timeOfMax = ws.Cells(row, 9).Value
    duration49kN = ws.Cells(row, 10).Value
    duration73kN = ws.Cells(row, 11).Value
    preProcess = ws.Cells(row, 12).Value
    sampleWeight = ws.Cells(row, 13).Value
    sampleTop = ws.Cells(row, 14).Value
    sampleColor = ws.Cells(row, 15).Value
    sampleLotNum = ws.Cells(row, 16).Value
    sampleHelLot = ws.Cells(row, 17).Value
    sampleBandLot = ws.Cells(row, 18).Value
    structureResult = ws.Cells(row, 19).Value
    penetrationResult = ws.Cells(row, 20).Value
    testSection = ws.Cells(row, 21).Value

    ' �f�o�b�O�o�͂Ŋe�t�B�[���h���m�F
    Debug.Print "Loaded Record - ID: " & ID & ", sampleID: " & sampleID & ", itemNum: " & itemNum & ", testPart: " & testPart
    Debug.Print "testDate: " & testDate & ", testTemp: " & testTemp & ", maxValue: " & maxValue & ", timeOfMax: " & timeOfMax
    Debug.Print "duration49kN: " & duration49kN & ", duration73kN: " & duration73kN & ", preProcess: " & preProcess
    Debug.Print "sampleWeight: " & sampleWeight & ", sampleTop: " & sampleTop & ", sampleColor: " & sampleColor
    Debug.Print "sampleLotNum: " & sampleLotNum & ", sampleHelLot: " & sampleHelLot & ", sampleBandLot: " & sampleBandLot
    Debug.Print "structureResult: "
End Sub


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

    For i = 2 To lastRow
        Dim record As New record
        record.LoadData ws, i

        testValues.Add record

        Call ClassifyKeys(record.sheetType, record.groupID, sheetNames)

        If Not groupedRecords.exists(record.sheetType) Then
            groupedRecords.Add record.sheetType, New Collection
        End If

        Call AddRecordToGroup(groupedRecords(record.sheetType), record)
        Dim j As Integer
        For j = 1 To groupedRecords(record.sheetType).Count
            Dim addedRecord As record
            Set addedRecord = groupedRecords(record.sheetType)(j)
            'Debug.Print "Record in group: ID=" & addedRecord.sampleID & " SheetType=" & addedRecord.sheetType & " GroupID=" & addedRecord.groupID & " SampleColor=" & addedRecord.sampleColor
        Next j
    Next i
    If Not groupedRecords Is Nothing Then
        For Each key In groupedRecords.keys
            Debug.Print "key: " & key & ", count:"; groupedRecords(key).Count
        Next key
    Else
        Debug.Print "groupedRecords is not initalized or empty."
    End If

    Call PrintGroupedRecords(groupedRecords, sheetNames)
    Debug.Print "Total unique records: " & testValues.Count
End Sub

Sub ClassifyKeys(sheetType As String, groupID As String, ByRef sheetNames As Object)
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
        sheetNames.Add sheetType, New Collection
    End If

    sheetNames(sheetType).Add baseTemplateName
    If additionalTemplateName <> "" Then
        sheetNames(sheetType).Add additionalTemplateName
    End If


    ' ��{�e���v���[�g�ƒǉ��e���v���[�g�̃V�[�g����
    Call ProcessTemplateSheet(baseTemplateName, sheetType, groupID, sheetTypeIndex)
    If additionalTemplateName <> "" Then
        Call ProcessTemplateSheet(additionalTemplateName, sheetType, groupID, sheetTypeIndex)
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object)
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



