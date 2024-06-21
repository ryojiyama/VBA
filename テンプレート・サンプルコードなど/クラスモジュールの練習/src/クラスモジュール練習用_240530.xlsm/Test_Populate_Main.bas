Attribute VB_Name = "Test_Populate_Main"

    
    'Main
Sub TestSheetCreationAndDataWriting()
    Call ResetSheetTypeIndex   ' インデックスをリセット
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

    ' レコードごとにシートネームを作成する
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")
    
    ' グループIDを作成
    groupID = Left(groupID, 2)
    
    Dim baseTemplateName As String
    Dim additionalTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "申請_飛来"
            additionalTemplateName = "定期_飛来"
        Case "Multi"
            baseTemplateName = "申請_墜落"
            additionalTemplateName = "定期_墜落"
        Case Else
            baseTemplateName = "その他"
            additionalTemplateName = ""
    End Select
    
    'シート名を保持する
    If sheetNames Is Nothing Then Set sheetNames = CreateObject("Scripting.Dictionary")
    If Not sheetNames.exists(sheetType) Then
        sheetNames.Add sheetType, CreateObject("Scripting.Dictionary")
    End If
    
    If sheetRecords Is Nothing Then Set sheetRecords = CreateObject("Scripting.Dictionary")
    If Not sheetRecords.exists(sheetType) Then
        sheetRecords.Add sheetType, CreateObject("Scripting.Dictionary")
    End If
    
    ' シート名を追加
    If Not sheetNames(sheetType).exists(baseTemplateName) Then
        sheetNames(sheetType).Add baseTemplateName, True
    End If
    If additionalTemplateName <> "" And Not sheetNames(sheetType).exists(additionalTemplateName) Then
        sheetNames(sheetType).Add additionalTemplateName, True
    End If
    
'    ' シート名を追加
'    If Not sheetNames(sheetType).Contains(baseTemplateName) Then
'        sheetNames(sheetType).Add baseTemplateName
'    End If
'    If additionalTemplateName <> "" And Not sheetNames(sheetType).Contains(additionalTemplateName) Then
'        sheetNames(sheetType).Add additionalTemplateName
'    End If

    ' 基本テンプレートと追加テンプレートのシート処理
    Call ProcessTemplateSheet(baseTemplateName, sheetType, groupID, sheetTypeIndex, sheetRecords, sheetNames, records)
    If additionalTemplateName <> "" Then
        Call ProcessTemplateSheet(additionalTemplateName, sheetType, groupID, sheetTypeIndex, sheetRecords, sheetNames, records)
    End If
End Sub

Sub ProcessTemplateSheet(templateName As String, sheetType As String, groupID As String, ByRef sheetTypeIndex As Object, ByRef sheetRecords As Object, ByRef sheetNames As Object, ByRef record As record)
    Dim combinedKey As String
    combinedKey = templateName & "_" & groupID
    ' シート名の決定
    If Not sheetTypeIndex.exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = combinedKey
        'Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    'Debug.Print "newSheetName: " & newSheetName
    
    ' シートの存在確認と取得/作成
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
    
    ' レコードをシートに追加する処理（必要に応じて追加）
    ' 例：newSheet.Cells(行, 列).Value = データ
    
    ' シート名とレコードIDを関連付ける
    If Not sheetRecords(sheetType).exists(newSheetName) Then
        sheetNames(sheetType).Add newSheetName, True
    End If
    ' sheetNames辞書に新しいシート名を保存
    If Not sheetRecords(sheetType).exists(newSheetName) Then
        sheetRecords(sheetType).Add newSheetName, New Collection
    End If
    sheetRecords(sheetType)(newSheetName).Add record
End Sub

Function SheetExists(sheetName As String) As Boolean
    ' シートの存在チェック
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing
End Function
Sub ResetSheetTypeIndex()
    ' インデックスのリセット
    Static sheetTypeIndex As Object
    Set sheetTypeIndex = Nothing ' Dictionary オブジェクトの解放
    Set sheetTypeIndex = CreateObject("Scripting.Dictionary") ' 新しい Dictionary オブジェクトの初期化
    Static groupSheetIndexes As Object
    Set groupSheetIndexes = Nothing ' Dictionary オブジェクトの解放
    Set groupSheetIndexes = CreateObject("Scripting.Dictionary") ' 新しい Dictionary オブジェクトの初期化
End Sub
Private Sub AddRecordToGroup(groupCollection As Collection, ByVal record As record)
    ' 新しいインスタンスを作成してから追加
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

    ' グループにレコードを追加
    groupCollection.Add newRecord

    ' デバッグ出力
    'Debug.Print "Adding record: ID=" & newRecord.sampleID & ", GroupID=" & newRecord.groupID & ", SheetType=" & newRecord.sheetType
End Sub


' -----------------------------------------------------------------------------------------------------
Sub PrintGroupedRecords(ByRef groupedRecords As Object, ByRef sheetNames As Object, ByRef sheetRecords As Object)
    Dim dictkey As Variant
    Dim sheetNameKey As Variant
    Dim record As record
    
    ' groupedRecordsの各sheetTypeをループ
    For Each dictkey In groupedRecords.keys
        Debug.Print "Sheet Type: " & dictkey & ", Number of Records: " & groupedRecords(dictkey).Count
        
        ' groupedRecords(dictkey)が正しく辞書オブジェクトを返すことを確認
        Dim sheetsDict As Object
        Set sheetsDict = groupedRecords(dictkey)
        
        ' 各シート名を出力
        If sheetNames.exists(dictkey) Then
            For Each sheetNameKey In sheetNames(dictkey).keys
                Debug.Print " sheet Name: " & sheetNameKey
                Debug.Print "TypeName:"; TypeName(groupedRecords(dictkey))
                ' sheetNameKeyでのレコードコレクションを取得
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








Sub ClassifyKeys1600_シートの生成までうまく行った｡(sheetType As String, groupID As String)
    ' レコードごとにシートネームを作成する
    Static sheetTypeIndex As Object
    If sheetTypeIndex Is Nothing Then Set sheetTypeIndex = CreateObject("Scripting.Dictionary")
    
    ' グループIDを作成
    groupID = Left(groupID, 2)
    
    Dim baseTemplateName As String
    Select Case sheetType
        Case "Single"
            baseTemplateName = "申請_飛来"
        Case "Multi"
            baseTemplateName = "申請_墜落"
        Case Else
            baseTemplateName = "その他"
    End Select
    
    Dim combinedKey As String
    combinedKey = sheetType & "_" & groupID
    
    ' シート名の決定
    If Not sheetTypeIndex.exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = baseTemplateName & "_" & groupID
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName
    
    ' シートの存在確認と取得/作成
    Dim newSheet As Worksheet
    If Not SheetExists(newSheetName) Then
        ' 指定された既存のシートをコピーして新しいシートを作成する
        Select Case baseTemplateName
            Case "申請_飛来", "申請_墜落", "定期_飛来", "定期_墜落", "側面試験", "依頼試験", "LOG_Helmet", "DataSheet"
                Worksheets(baseTemplateName).Copy After:=Worksheets(Worksheets.Count)
                Set newSheet = Worksheets(Worksheets.Count)
                newSheet.name = newSheetName
                ' オブジェクト名を"Temp_" & newSheetNameに設定
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

    ' デバッグ出力：同じグループIDを持つレコードが正しく分類されているか確認
    Debug.Print "Record added to sheet: " & newSheetName & " for groupID: " & groupID

    ' レコードをシートに追加する処理（必要に応じて追加）
    ' 例：newSheet.Cells(行, 列).Value = データ
End Sub

Sub GenerateSheets()
    Dim sheetNames As Collection
    Set sheetNames = New Collection
    
    ' 事前に定義されたシート名のリスト
    sheetNames.Add "LOG_Helmet"
    sheetNames.Add "DataSheet"
    sheetNames.Add "申請_飛来"
    sheetNames.Add "申請_墜落"
    sheetNames.Add "定期_飛来"
    sheetNames.Add "定期_墜落"
    sheetNames.Add "側面試験"
    sheetNames.Add "依頼試験"
    
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
        ' 必要な変数を取得
        Dim position As String
        Dim number As String
        Dim condition As String
        Dim recordType As String

        position = record.testPart  ' Locationプロパティを使用
        number = record.ID  ' IDプロパティを使用
        condition = record.Temperature  ' Temperatureプロパティを使用
        recordType = "Single"  ' 固定値（適切なプロパティがないため）

        ' Recordオブジェクトの各プロパティが存在するかチェック
        On Error Resume Next
        Debug.Print "Checking properties for Record:"
        Debug.Print "  ID: " & record.ID
        Debug.Print "  Location: " & record.testPart
        Debug.Print "  Temperature: " & record.Temperature
        Debug.Print "  DateValue: " & record.DateValue
        Debug.Print "  TemperatureValue: " & record.TemperatureValue
        Debug.Print "  Force: " & record.Force
        On Error GoTo 0

        ' エラーが発生するプロパティを特定
        If Err.number <> 0 Then
            Debug.Print "Error accessing property: " & Err.Description
            Exit Sub
        End If

        ' グループキーの生成と分類処理
        Dim groupKey As String
        If position = "側" Then
            groupKey = number & "-" & condition & "-側"
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
                Debug.Print "    ID=" & record.ID & ", Location=" & record.testPart  ' 各RecordのIDとLocationを出力
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

    ' シートの作成ロジック
    sheetIndex = 1
    For Each groupKey In singleGroups.keys
        Set templateNames = New Collection

        If InStr(groupKey, "SingleValue") > 0 Then
            templateNames.Add "申請_飛来"
            templateNames.Add "定期_飛来"
        ElseIf InStr(groupKey, "SideValue") > 0 Then
            templateNames.Add "側面試験"
        Else
            templateNames.Add "申請_墜落"
            templateNames.Add "定期_墜落"
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
            templateNames.Add "申請_飛来"
            templateNames.Add "定期_飛来"
        ElseIf InStr(groupKey, "SideValue") > 0 Then
            templateNames.Add "側面試験"
        Else
            templateNames.Add "申請_墜落"
            templateNames.Add "定期_墜落"
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
    ' シート名の決定
    If Not sheetTypeIndex.exists(combinedKey) Then
        sheetTypeIndex(combinedKey) = combinedKey
        Debug.Print "Added new entry to sheetTypeIndex: " & combinedKey & " = " & sheetTypeIndex(combinedKey)
    End If
    
    Dim newSheetName As String
    newSheetName = sheetTypeIndex(combinedKey)
    Debug.Print "newSheetName: " & newSheetName
    
    ' シートの存在確認と取得/作成
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

    ' デバッグ出力：同じグループIDを持つレコードが正しく分類されているか確認
    'Debug.Print "Record added to sheet: " & newSheet.name & " for groupID: " & groupID

    ' レコードをシートに追加する処理（必要に応じて追加）
    ' 例：newSheet.Cells(行, 列).Value = データ
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
'        ' Debug.PrintにRecordの具体的なプロパティを指定
'        Debug.Print "Record loaded: ID=" & Record.ID & " Group=" & Record.testPart  ' Recordの具体的なプロパティを出力
'
'        testValues.Add Record
'
'        ' 分類されたグループにレコードを追加
'        If Not groupedRecords.Exists(Record.testPart) Then
'            groupedRecords.Add Record.testPart, New Collection
'        End If
'        groupedRecords(Record.testPart).Add Record
'    Next i
'
'    ' グループの内容を確認（デバッグ用）
'    Dim key As Variant
'    For Each key In groupedRecords
'        'Debug.Print "Main()_Group: " & key & ", Count: " & groupedRecords(key).Count
'    Next key
'
'    ' singleGroupsとmultiGroupsを適切に初期化する
'    Dim singleGroups As Scripting.Dictionary
'    Set singleGroups = CreateObject("Scripting.Dictionary")
'    Dim multiGroups As Scripting.Dictionary
'    Set multiGroups = CreateObject("Scripting.Dictionary")
'
'    ' キーの分類
'    Debug.Print "testValues count: " & testValues.Count  ' testValuesの件数を確認
'
'    ' 型確認のためのデバッグ出力
'    Debug.Print "singleGroups type: " & TypeName(singleGroups)  ' singleGroupsの型を確認
'    Debug.Print "multiGroups type: " & TypeName(multiGroups)    ' multiGroupsの型を確認
'
'    Call ClassifyKeys(testValues, singleGroups, multiGroups)
'
'    ' 結果の表示（必要に応じてコメントアウト）
'    Debug.Print "SingleValue Groups:"
'    Call PrintGroups(singleGroups)
'
'    Debug.Print "MultiValue Groups:"
'    Call PrintGroups(multiGroups)
'
'    ' データのグループ化とシート書き込みを行う
'    Call PopulateGroupedSheets(singleGroups, multiGroups)
'End Sub

'Function SheetExists_20240613(sheetName As String) As Boolean
'    ' PopulateGroupedSheetsのサブプロシージャ
'    Dim tmpSheet As Worksheet
'    On Error Resume Next
'    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
'    SheetExists = Not tmpSheet Is Nothing
'    On Error GoTo 0
'End Function
