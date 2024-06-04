VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Record クラスモジュール
Option Explicit

Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As Double
Public Force As Double
Public Values As Collection

Public Sub InitSingle()
    ' 初期化処理
End Sub

Public Sub InitMultiple()
    Set Values = New Collection
End Sub

Public Function GetSpecificValues() As Collection
    Set GetSpecificValues = Values
End Function


Public Sub Initialize( _
    ByVal ID As String, _
    ByVal Temperature As String, _
    ByVal Location As String, _
    ByVal DateValue As Date, _
    ByVal TemperatureValue As Double, _
    ByVal Force As Double)

    Me.ID = ID
    Me.Temperature = Temperature
    Me.Location = Location
    Me.DateValue = DateValue
    Me.TemperatureValue = TemperatureValue
    Me.Force = Force
End Sub


' データをロードし、分類およびグループ化を行うメソッド
Public Sub LoadData(ByVal ws As Worksheet, ByVal row As Integer)
    ID = ws.Cells(row, 1).Value
    Temperature = ws.Cells(row, 2).Value
    Location = ws.Cells(row, 3).Value
    DateValue = ws.Cells(row, 4).Value
    TemperatureValue = ws.Cells(row, 5).Value
    Force = ws.Cells(row, 6).Value

    ' IDを分析してカテゴリを決定
    Dim parts() As String
    parts = Split(ID, "-")

    ' 中間部分でのカテゴリ判定
    If InStr(parts(1), "F") > 0 Then
        Category = "MultiValue"
    Else
        Category = "SingleValue"
    End If

    ' 末尾部分でのグループ判定
    Select Case parts(3)
        Case "側"
            Group = "OtherValue." & parts(3) & "." & parts(2)
        Case Else
            Group = Category & "." & parts(3) & "." & parts(2)
    End Select
End Sub
