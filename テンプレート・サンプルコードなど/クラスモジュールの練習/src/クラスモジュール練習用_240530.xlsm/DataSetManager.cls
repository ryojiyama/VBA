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

Public Sub WriteSelectedValuesToOutputSheet(sourceSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(sheetName)

    ws.Range("B1:G1").Value = Array("ID", "Temperature", "Location", "DateValue", "TemperatureValue", "Force")
    row = 2

    For Each record In specificValues
        ws.Cells(row, "B").Value = record.ID
        ws.Cells(row, "C").Value = record.Temperature
        ws.Cells(row, "D").Value = record.Location
        ws.Cells(row, "E").Value = record.DateValue
        ws.Cells(row, "F").Value = record.TemperatureValue
        ws.Cells(row, "G").Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToOutputSheet: " & Err.Description
    Err.Clear
End Sub

Public Sub WriteSelectedValuesToRstlSheet(sourceSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(sheetName)

    ws.Range("B1:G1").Value = Array("ID", "Temperature", "Location", "DateValue", "TemperatureValue", "Force")
    row = 2

    For Each record In specificValues
        ws.Cells(row, "B").Value = record.ID
        ws.Cells(row, "C").Value = record.Temperature
        ws.Cells(row, "D").Value = record.Location
        ws.Cells(row, "E").Value = record.DateValue
        ws.Cells(row, "F").Value = record.TemperatureValue
        ws.Cells(row, "G").Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToRstlSheet: " & Err.Description
    Err.Clear
End Sub

Public Sub WriteSelectedValuesToSideSheet(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' ヘッダーを書き込み
    ws.Range("B1:G1").Value = Array("ID", "Temperature", "Location", "DateValue", "TemperatureValue", "Force")
    row = 2

    ' データを書き込み
    For Each record In specificValues
        ws.Cells(row, "B").Value = record.ID
        ws.Cells(row, "C").Value = record.Temperature
        ws.Cells(row, "D").Value = record.Location
        ws.Cells(row, "E").Value = record.DateValue
        ws.Cells(row, "F").Value = record.TemperatureValue
        ws.Cells(row, "G").Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToSideSheet: " & Err.Description
    Err.Clear
End Sub

Public Sub WriteSelectedValuesToResultTempSheet(sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(sheetName)

    ws.Range("B1:G1").Value = Array("ID", "Temperature", "Location", "DateValue", "TemperatureValue", "Force")
    row = 2

    For Each record In specificValues
        ws.Cells(row, "B").Value = record.ID
        ws.Cells(row, "C").Value = record.Temperature
        ws.Cells(row, "D").Value = record.Location
        ws.Cells(row, "E").Value = record.DateValue
        ws.Cells(row, "F").Value = record.TemperatureValue
        ws.Cells(row, "G").Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToResultTempSheet: " & Err.Description
    Err.Clear
End Sub





' ------------------------------------------------------------
Public Sub WriteSelectedValuesToOutputSheet_Change240604(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)

    Debug.Print "Writing data to sheet: " & ws.name

    ' ヘッダーを書き込み
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Temperature"
    ws.Range("C1").Value = "Location"
    ws.Range("D1").Value = "DateValue"
    ws.Range("E1").Value = "TemperatureValue"
    ws.Range("F1").Value = "Force"

    row = 2

    ' データを書き込み
    For Each record In specificValues
        ws.Range("A" & row).Value = record.ID
        ws.Range("B" & row).Value = record.Temperature
        ws.Range("C" & row).Value = record.Location
        ws.Range("D" & row).Value = record.DateValue
        ws.Range("E" & row).Value = record.TemperatureValue
        ws.Range("F" & row).Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Class: Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToOutputSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub

Public Sub WriteSelectedValuesToRstlSheet_Change240604(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' ヘッダーを書き込み
    ws.Range("B1").Value = "ID"
    ws.Range("C1").Value = "Temperature"
    ws.Range("D1").Value = "Location"
    ws.Range("E1").Value = "DateValue"
    ws.Range("F1").Value = "TemperatureValue"
    ws.Range("G1").Value = "Force"

    row = 2

    ' データを書き込み前にコレクションの内容を確認
    For Each record In specificValues
        Debug.Print "Record ID: " & record.ID & ", Temp: " & record.Temperature & ", Location: " & record.Location & ", Date: " & record.DateValue & ", TempValue: " & record.TemperatureValue & ", Force: " & record.Force
    Next record

    ' データを書き込み
    For Each record In specificValues
        ws.Range("B" & row).Value = record.ID
        ws.Range("C" & row).Value = record.Temperature
        ws.Range("D" & row).Value = record.Location
        ws.Range("E" & row).Value = record.DateValue
        ws.Range("F" & row).Value = record.TemperatureValue
        ws.Range("G" & row).Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Class: Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToRstlSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub


Public Sub WriteSelectedValuesToSideSheet_Change240604(templateSheetName As String, sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' ヘッダーを書き込み
    ws.Range("B1").Value = "ID"
    ws.Range("C1").Value = "Temperature"
    ws.Range("D1").Value = "Location"
    ws.Range("E1").Value = "DateValue"
    ws.Range("F1").Value = "TemperatureValue"
    ws.Range("G1").Value = "Force"

    row = 2

    ' データを書き込み前にコレクションの内容を確認
    For Each record In specificValues
        Debug.Print "Record ID: " & record.ID & ", Temp: " & record.Temperature & ", Location: " & record.Location & ", Date: " & record.DateValue & ", TempValue: " & record.TemperatureValue & ", Force: " & record.Force
    Next record

    ' データを書き込み
    For Each record In specificValues
        ws.Range("B" & row).Value = record.ID
        ws.Range("C" & row).Value = record.Temperature
        ws.Range("D" & row).Value = record.Location
        ws.Range("E" & row).Value = record.DateValue
        ws.Range("F" & row).Value = record.TemperatureValue
        ws.Range("G" & row).Value = record.Force
        row = row + 1
    Next record
    Debug.Print "Class: Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToRstlSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub


Public Sub WriteSelectedValuesToResultTempSheet_Change240604(sheetName As String, specificValues As Collection)
    Dim ws As Worksheet
    Dim row As Integer
    Dim record As Variant

    On Error GoTo ErrorHandler

    ' 新しいシートを参照
    Set ws = ThisWorkbook.Sheets(sheetName)

    Debug.Print "Writing data to sheet: " & ws.name

    ' ヘッダーを書き込み
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Temperature"
    ws.Range("C1").Value = "Location"
    ws.Range("D1").Value = "DateValue"
    ws.Range("E1").Value = "TemperatureValue"
    ws.Range("F1").Value = "Force"

    row = 2

    ' データを書き込み
    For Each record In specificValues
        ws.Range("A" & row).Value = record.ID
        ws.Range("B" & row).Value = record.Temperature
        ws.Range("C" & row).Value = record.Location
        ws.Range("D" & row).Value = record.DateValue
        ws.Range("E" & row).Value = record.TemperatureValue
        ws.Range("F" & row).Value = record.Force
        row = row + 1
    Next record

    Debug.Print "Class: Data written successfully to " & ws.name
    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteSelectedValuesToResultTempSheet: " & Err.Description
    Err.Clear
    Resume Next
End Sub
