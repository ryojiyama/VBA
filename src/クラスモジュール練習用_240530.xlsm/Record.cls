VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Record 郢ｧ�ｽｯ郢晢ｽｩ郢ｧ�ｽｹ郢晢ｽ｢郢ｧ�ｽｸ郢晢ｽ･郢晢ｽｼ郢晢ｽｫ
Option Explicit

Public ID As String
Public Temperature As String
Public Location As String
Public DateValue As Date
Public TemperatureValue As Double
Public Force As Double
Public Values As Collection
Public Category As String
Public Group As String

Public Sub InitSingle()
    ' 陋ｻ譎�謔�陋ｹ髢�ｿｽ�ｽｦ騾��ｿｽ
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


' 郢晢ｿｽ郢晢ｽｼ郢ｧ�ｽｿ郢ｧ蛛ｵﾎ溽ｹ晢ｽｼ郢晏ｳｨ�ｼ�邵ｲ竏晢ｿｽ�ｿｽ鬯俶ｧｭ笙郢ｧ蛹ｻ�ｿｽ�ｽｳ郢ｧ�ｽｰ郢晢ｽｫ郢晢ｽｼ郢晄懷密郢ｧ螳夲ｽ｡蠕娯鴬郢晢ｽ｡郢ｧ�ｽｽ郢晢ｿｽ郢晢ｿｽ
Public Sub LoadData(ByVal ws As Worksheet, ByVal row As Integer)
    ID = ws.Cells(row, 1).Value
    Temperature = ws.Cells(row, 2).Value
    Location = ws.Cells(row, 3).Value
    DateValue = ws.Cells(row, 4).Value
    TemperatureValue = ws.Cells(row, 5).Value
    Force = ws.Cells(row, 6).Value

    ' ID郢ｧ雋橸ｿｽ�ｿｽ隴ｫ闊鯉ｼ�邵ｺ�ｽｦ郢ｧ�ｽｫ郢晢ｿｽ郢ｧ�ｽｴ郢晢ｽｪ郢ｧ蜻茨ｽｱ�ｽｺ陞ｳ�ｿｽ
    Dim parts() As String
    parts = Split(ID, "-")

    ' 闕ｳ�ｽｭ鬮｢鬥ｴﾎ夊崕�ｿｽ邵ｺ�ｽｧ邵ｺ�ｽｮ郢ｧ�ｽｫ郢晢ｿｽ郢ｧ�ｽｴ郢晢ｽｪ陋ｻ�ｽ､陞ｳ�ｿｽ
    If InStr(parts(1), "F") > 0 Then
        Category = "MultiValue"
    Else
        Category = "SingleValue"
    End If

    ' 末尾部分でのグループ判定
    Select Case parts(3)
        Case "陋幢ｽｴ"
            Group = "OtherValue." & parts(3) & "." & parts(2)
        Case Else
            Group = Category & "." & parts(3) & "." & parts(2)
    End Select
End Sub
