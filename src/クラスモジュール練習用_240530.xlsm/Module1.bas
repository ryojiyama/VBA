Attribute VB_Name = "Module1"



''新しいコードには含まれていない。
'Function GenerateSheetName(prefix As String, index As Integer) As String
'    GenerateSheetName = prefix & Format(index, "00")
'End Function
'
'
'
'
'' Mainプロシージャ
'Sub TestSheetCreationAndDataWriting()
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
'        testValues.Add Record
'
'        ' 分類されたグループにレコードを追加
'        If Not groupedRecords.Exists(Record.group) Then
'            groupedRecords.Add Record.group, New Collection
'        End If
'        groupedRecords(Record.group).Add Record
'    Next i
'
'    ' グループの内容を確認（デバッグ用）
'    Dim key As Variant
'    For Each key In groupedRecords
'        'Debug.Print "Main()_Group: " & key & ", Count: " & groupedRecords(key).Count
'    Next key
'
'    ' データのグループ化とシート書き込みを行う
'    Call PopulateGroupedSheets(groupedRecords)
'End Sub
'
'
'Sub PopulateGroupedSheets(groupedRecords As Object)
'    Dim ws As Worksheet
'    Dim sheetIndex As Integer
'    Dim key As Variant
'    Dim newSheetName As String
'    Dim templateNames As Collection
'    Dim templateName As Variant
'    Dim keyPrefix As String
'
'    sheetIndex = 1
'
'    For Each key In groupedRecords.keys
'        Set templateNames = New Collection
'
'        If InStr(key, "SingleValue") > 0 Then
'            templateNames.Add "申請_飛来"
'            templateNames.Add "定期_飛来"
'        ElseIf InStr(key, "SideValue") > 0 Then
'            templateNames.Add "側面試験"
'        Else
'            templateNames.Add "申請_墜落"
'            templateNames.Add "定期_墜落"
'        End If
'
'        For Each templateName In templateNames
'            newSheetName = templateName & "_" & sheetIndex
'            If Not SheetExists(newSheetName) Then
'                Debug.Print "key:"; key
'                'Debug.Print "New sheet would be created: " & newSheetName
'                sheetIndex = sheetIndex + 1
'            Else
'                Debug.Print "Sheet already exists: " & newSheetName
'            End If
'        Next templateName
'    Next key
'End Sub
'
'Sub CopyAndPopulateSheet(templateSheetName As String, newSheetName As String, dataCollection As Collection)
'    Dim sourceSheet As Worksheet, targetSheet As Worksheet
'    Dim lastRow As Long
'    Dim Record As Variant
'    Dim copyCount As Integer
'    Dim newCodeName As String
'
'    ' テンプレートシートが存在することを確認
'    Set sourceSheet = ThisWorkbook.Sheets(templateSheetName)
'
'    ' テンプレートシートをコピー
'    sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
'    Set targetSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
'    targetSheet.Name = newSheetName
'
'    ' コピー回数をインクリメント
'    If Not IsError(Application.Evaluate("'" & sourceSheet.Name & "'!Temp_0")) Then
'        copyCount = Application.Evaluate("'" & sourceSheet.Name & "'!Temp_0") + 1
'    Else
'        copyCount = 1  ' 名前定義が存在しない場合、初めてのコピーとする
'    End If
'    sourceSheet.Range("Z1").Value = copyCount
'    sourceSheet.Names.Add Name:="Temp_0", RefersToR1C1:=sourceSheet.Range("Z1")
'
'    ' 新しいオブジェクト名を設定
'    newCodeName = "Temp" & copyCount & "_" & Mid(sourceSheet.CodeName, InStr(sourceSheet.CodeName, "_") + 1)
'    ThisWorkbook.VBProject.VBComponents(targetSheet.CodeName).Name = newCodeName
'
'    ' 新しいシートにデータを書き込み
'    lastRow = 2  ' ヘッダーが最初の行にあると仮定
'    For Each Record In dataCollection
'        With targetSheet
'            .Cells(lastRow, "B").Value = Record.ID
'            .Cells(lastRow, "C").Value = Record.Temperature
'            .Cells(lastRow, "D").Value = Record.Location
'            .Cells(lastRow, "E").Value = Record.DateValue
'            .Cells(lastRow, "F").Value = Record.TemperatureValue
'            .Cells(lastRow, "G").Value = Record.Force
'            lastRow = lastRow + 1  ' 行カウンターをインクリメント
'        End With
'    Next Record
'End Sub
'
'
'
'
'Function SheetExists(sheetName As String) As Boolean
'    ' PopulateGroupedSheetsのサブプロシージャ
'    Dim tmpSheet As Worksheet
'    On Error Resume Next
'    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
'    SheetExists = Not tmpSheet Is Nothing
'    On Error GoTo 0
'End Function
'
'
'Sub InitializeTempValues()
'    Dim ws As Worksheet
'
'    For Each ws In ThisWorkbook.Sheets
'        ws.Range("Z1").Value = 0
'        ws.Names.Add Name:="Temp_0", RefersTo:=ws.Range("Z1")
'    Next ws
'
'    MsgBox "すべてのシートに名前定義 'Temp_0' が設定されました。", vbInformation
'End Sub
'
'
'
'Sub PopulateGroupedSheets_06101120(groupedRecords As Object)
'    Dim ws As Worksheet
'    Dim sheetIndex As Integer
'    Dim key As Variant
'    Dim newSheetName As String
'    Dim templateName As String
'
'    sheetIndex = 1
'
'    For Each key In groupedRecords.keys
'        ' Template sheet determination based on group key
'        If InStr(key, "SingleValue") > 0 Then
'            templateName = "申請_飛来"
'        ElseIf InStr(key, "SideValue") > 0 Then
'            templateName = "側面試験"
'        Else
'            templateName = "申請_墜落"
'        End If
'
'        ' Generate unique sheet name
'        Debug.Print "key:"; key
'        newSheetName = key & "_" & sheetIndex
'
'        ' Check if the sheet already exists
'        If Not SheetExists(newSheetName) Then
'            ' Copy and populate the sheet if it does not exist
'            Call CopyAndPopulateSheet(templateName, newSheetName, groupedRecords(key))
'            sheetIndex = sheetIndex + 1  ' Increment sheet index only if a new sheet was created
'        Else
'            ' Optionally, you can handle the case where the sheet already exists
'            Debug.Print "Sheet already exists: " & newSheetName
'        End If
'    Next key
'End Sub
