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


