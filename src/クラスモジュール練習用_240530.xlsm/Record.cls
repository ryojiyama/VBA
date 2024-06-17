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

' Public変数の定義
Public ID As String 'ID
Public sampleID As String '試料ID
Public itemNum As String '品番
Public testPart As String '試験位置
Public testDate As Date '検査日
Public testTemp As Double '温度
Public maxValue As Double '最大値
Public timeOfMax As Double '最大値を記録した時間
Public duration49kN As Double '4.9kNの継続時間
Public duration73kN As Double '7.3kNの継続時間
Public preProcess As String '前処理
Public sampleWeight As Double '重量
Public sampleTop As Double '天頂隙間
Public sampleColor As String '帽体色
Public sampleLotNum As String '製品ロット
Public sampleHelLot As String '帽体ロット
Public sampleBandLot As String '内装ロット
Public structureResult As String '構造検査
Public penetrationResult As String '貫通検査
Public testSection As String '試験区分


Public Values As Collection
Public sheetType As String
Public groupID As String

' 初期化メソッド
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


' データをロードし、分類およびグループ化を行うメソッド
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

    ' IDを分析してカテゴリを決定
    Dim parts() As String
    parts = Split(sampleID, "-")

    ' 中間部分でのカテゴリ判定
    If InStr(parts(1), "F") > 0 And InStr(parts(2), "側") > 0 Then
        sheetType = "Side"
    ElseIf InStr(parts(1), "F") > 0 Then
        sheetType = "Multi"
    Else
        sheetType = "Single"
    End If

    ' 末尾部分でのグループ判定
    Select Case parts(3)
        Case "側"
            groupID = parts(0) & "." & "SideValue." & parts(1) & "." & parts(3) & "." & parts(2)
        Case Else
            groupID = parts(0) & "." & parts(1) & "." & parts(3) & "." & parts(2)
    End Select
End Sub

' Record クラスに追加
Public Function Equals(other As record) As Boolean
    Equals = (Me.sampleID = other.sampleID) And (Me.groupID = other.groupID) And (Me.sheetType = other.sheetType)
End Function

' Recordクラスの'Clone'メソッド
'Public Function Clone() As record
'    Dim newRecord As New record
'
'    ' 各プロパティを新しいインスタンスにコピー
'    newRecord.ID = Me.ID
'    newRecord.sampleID = Me.sampleID
'    newRecord.itemNum = Me.itemNum
'    newRecord.testPart = Me.testPart
'    newRecord.testDate = Me.testDate
'    newRecord.testTemp = Me.testTemp
'    newRecord.maxValue = Me.maxValue
'    newRecord.timeOfMax = Me.timeOfMax
'    newRecord.duration49kN = Me.duration49kN
'    newRecord.duration73kN = Me.duration73kN
'    newRecord.preProcess = Me.preProcess
'    newRecord.sampleWeight = Me.sampleWeight
'    newRecord.sampleTop = Me.sampleTop
'    newRecord.sampleColor = Me.sampleColor
'    newRecord.sampleLotNum = Me.sampleLotNum
'    newRecord.sampleHelLot = Me.sampleHelLot
'    newRecord.sampleBandLot = Me.sampleBandLot
'    newRecord.structureResult = Me.structureResult
'    newRecord.penetrationResult = Me.penetrationResult
'    newRecord.testSection = Me.testSection
'
'    Set Clone = newRecord
'End Function


Public Sub LoadData_Test(ByVal ws As Worksheet, ByVal row As Integer)
    On Error GoTo ErrorHandler

    ID = ws.Cells(row, 2).Value ' IDが第1列にあると仮定
    testTemp = ws.Cells(row, 7).Value ' 温度が第2列にあると仮定
    testPart = ws.Cells(row, 5).Value ' 位置情報が第3列にあると仮定
    testDate = ws.Cells(row, 6).Value ' 日付が第4列にあると仮定
    sampleWeight = ws.Cells(row, 13).Value ' 温度値が第5列にあると仮定
    maxValue = ws.Cells(row, 8).Value ' 力が第6列にあると仮定

    Exit Sub

ErrorHandler:
    Debug.Print "Error loading data at column: " & Err.Source & " Row: " & row & " Error #" & Err.number & ": " & Err.Description
End Sub
