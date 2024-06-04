Attribute VB_Name = "Module1"
Sub CopyAndPopulateSheet(templateSheetName As String, newSheetName As String, dataCollection As Collection)
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    Dim record As Variant

    ' Ensure template exists
    Set sourceSheet = ThisWorkbook.Sheets(templateSheetName)
    sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set targetSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    targetSheet.Name = newSheetName

    ' Write data to the new sheet
    lastRow = 2 ' Assuming headers are in the first row
    For Each record In dataCollection
        With targetSheet
            .Cells(lastRow, "B").Value = record.ID
            .Cells(lastRow, "C").Value = record.Temperature
            .Cells(lastRow, "D").Value = record.Location
            .Cells(lastRow, "E").Value = record.DateValue
            .Cells(lastRow, "F").Value = record.TemperatureValue
            .Cells(lastRow, "G").Value = record.Force
        End With
        lastRow = lastRow + 1
    Next record
End Sub


'新しいコードには含まれていない。
Function GenerateSheetName(prefix As String, index As Integer) As String
    GenerateSheetName = prefix & Format(index, "00")
End Function




' Mainプロシージャ
Sub TestSheetCreationAndDataWriting()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DataSheet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Integer

    Dim testValues As New Collection
    Dim groupedRecords As Object
    Set groupedRecords = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim Record As New Record
        Record.LoadData ws, i
        testValues.Add Record

        ' 分類されたグループにレコードを追加
        If Not groupedRecords.Exists(Record.Group) Then
            groupedRecords.Add Record.Group, New Collection
        End If
        groupedRecords(Record.Group).Add Record
    Next i

    ' グループの内容を確認（デバッグ用）
    Dim key As Variant
    For Each key In groupedRecords
        Debug.Print "Group: " & key & ", Count: " & groupedRecords(key).Count
    Next key

    ' データのグループ化とシート書き込みを行う
    Call PopulateGroupedSheets
End Sub

Sub PopulateGroupedSheets(groupedRecords As Object)
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim key As Variant
    Dim newSheetName As String
    Dim templateName As String

    sheetIndex = 1

    For Each key In groupedRecords.Keys
        ' Template sheet determination based on group key
        If InStr(key, "SingleValue") > 0 Then
            templateName = "Temp_Shinsei"
        ElseIf InStr(key, "OtherValue") > 0 Then
            templateName = "Temp_Teiki"
        Else
            templateName = "Temp_Irai"
        End If

        ' Generate unique sheet name
        newSheetName = key & "_" & sheetIndex

        ' Check if the sheet already exists
        If Not SheetExists(newSheetName) Then
            ' Copy and populate the sheet if it does not exist
            Call CopyAndPopulateSheet(templateName, newSheetName, groupedRecords(key))
            sheetIndex = sheetIndex + 1  ' Increment sheet index only if a new sheet was created
        Else
            ' Optionally, you can handle the case where the sheet already exists
            Debug.Print "Sheet already exists: " & newSheetName
        End If
    Next key
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim tmpSheet As Worksheet
    On Error Resume Next
    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not tmpSheet Is Nothing
    On Error GoTo 0
End Function
