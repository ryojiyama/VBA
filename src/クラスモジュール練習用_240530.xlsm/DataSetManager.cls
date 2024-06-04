VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DataSetManager クラスモジュール
Option Explicit
Private records As Collection

Public Sub Init_001()
    Set records = New Collection
    Debug.Print "records initialized"
End Sub

Public Sub LoadData(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim rec As Record

    If records Is Nothing Then
        Debug.Print "records is nothing"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Debug.Print "Last Row: " & lastRow

    For i = 2 To lastRow
        Set rec = New Record
        Debug.Print "New record created for row: " & i
        rec.Initialize _
            ID:=ws.Cells(i, 1).Value, _
            Temperature:=ws.Cells(i, 2).Value, _
            Location:=ws.Cells(i, 3).Value, _
            DateValue:=ws.Cells(i, 4).Value, _
            TemperatureValue:=ws.Cells(i, 5).Value, _
            Force:=ws.Cells(i, 6).Value
        Debug.Print "Record initialized for row: " & i & " with ID: " & ws.Cells(i, 1).Value
        records.Add rec
        Debug.Print "Record added for row: " & i
    Next i
End Sub


Public Function GetRecords() As Collection
    Set GetRecords = records
End Function


' DataSetManager クラスモジュール
Public Sub Init()
    ' 初期化処理
    Debug.Print "DataSetManager initialized."
End Sub




Function SheetExists(sheetName As String) As Boolean
    Dim tmpSheet As Worksheet
    On Error Resume Next
    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not tmpSheet Is Nothing
End Function



Public Sub WriteSelectedValuesToOutputSheet(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim Record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Error: sheet not found - " & sheetName
        Exit Sub
    End If
    
    Debug.Print "Writing data to sheet: " & ws.Name
    
    ' ヘッダーを書き込み
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Temperature"
    ws.Range("C1").Value = "Location"
    ws.Range("D1").Value = "DateValue"
    ws.Range("E1").Value = "TemperatureValue"
    ws.Range("F1").Value = "Force"
    
    row = 2
    
    ' データを書き込み
    For Each Record In specificValues
        ws.Range("A" & row).Value = Record.ID
        ws.Range("B" & row).Value = Record.Temperature
        ws.Range("C" & row).Value = Record.Location
        ws.Range("D" & row).Value = Record.DateValue
        ws.Range("E" & row).Value = Record.TemperatureValue
        ws.Range("F" & row).Value = Record.Force
        row = row + 1
    Next Record
    Debug.Print "Class: Data written successfully to " & ws.Name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToOutputSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub

Public Sub WriteSelectedValuesToRstlSheet(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim Record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Error: sheet not found - " & sheetName
        Exit Sub
    End If
    
    ' ヘッダーを書き込み
    ws.Range("B1").Value = "ID"
    ws.Range("C1").Value = "Temperature"
    ws.Range("D1").Value = "Location"
    ws.Range("E1").Value = "DateValue"
    ws.Range("F1").Value = "TemperatureValue"
    ws.Range("G1").Value = "Force"
    
    row = 2
    
    ' データを書き込み前にコレクションの内容を確認
    For Each Record In specificValues
        Debug.Print "Record ID: " & Record.ID & ", Temp: " & Record.Temperature & ", Location: " & Record.Location & ", Date: " & Record.DateValue & ", TempValue: " & Record.TemperatureValue & ", Force: " & Record.Force
    Next Record

    ' データを書き込み
    For Each Record In specificValues
        ws.Range("B" & row).Value = Record.ID
        ws.Range("C" & row).Value = Record.Temperature
        ws.Range("D" & row).Value = Record.Location
        ws.Range("E" & row).Value = Record.DateValue
        ws.Range("F" & row).Value = Record.TemperatureValue
        ws.Range("G" & row).Value = Record.Force
        row = row + 1
    Next Record
    Debug.Print "Class: Data written successfully to " & ws.Name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToRstlSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub





Public Sub WriteSelectedValuesToResultTempSheet(sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim Record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Error: sheet not found - " & sheetName
        Exit Sub
    End If
    
    Debug.Print "Writing data to sheet: " & ws.Name
    
    ' ヘッダーを書き込み
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Temperature"
    ws.Range("C1").Value = "Location"
    ws.Range("D1").Value = "DateValue"
    ws.Range("E1").Value = "TemperatureValue"
    ws.Range("F1").Value = "Force"
    
    row = 2
    
    ' データを書き込み
    For Each Record In specificValues
        ws.Range("A" & row).Value = Record.ID
        ws.Range("B" & row).Value = Record.Temperature
        ws.Range("C" & row).Value = Record.Location
        ws.Range("D" & row).Value = Record.DateValue
        ws.Range("E" & row).Value = Record.TemperatureValue
        ws.Range("F" & row).Value = Record.Force
        row = row + 1
    Next Record
    
    Debug.Print "Class: Data written successfully to " & ws.Name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToResultTempSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub




