Sub Practice_match()
    Dim moveTimesSh As Worksheet
    Dim move2Sh As Worksheet
    Set moveTimesSh = ThisWorkbook.Worksheets("moveTimesSh")
    Set move2Sh = ThisWorkbook.Worksheets("Move2")

    Dim move2Row As Long
    move2Row = 156

    Dim move2Rng As Range
    Set move2Rng = Range(move2Sh.Cells(2, 3), move2Sh.Cells(move2Row, 3))

    Dim workEndR As Long
    Dim workTmpR As Long
    Dim moveName1 As String
    Dim move1stTime As Byte '技1のターン数
    Dim move1stCharge As Byte '技1のチャージEng
    Dim move2ndCharge As Byte '技2の必要Eng



    workEndR = moveTimesSh.Cells(Rows.Count, 1).End(xlUp).Row

    For workTmpR = 2 To workEndR
        moveName1 = moveTimesSh.Cells(workTmpR, 4).Value
        On Error Resume Next
        moveTimesSh.Cells(workTmpR, 2).Value = move2Sh.Cells(Application.WorksheetFunction.match(moveName1, move2Rng, 0) + 1, 6)
        If Err <> 0 Then
            moveTimesSh.Cells(workTmpR, 2).Value = "ER"
            Err.Clear
        End If
    Next
End Sub


Sub Charge_Table()

Dim i As Long 'カウンタ1
Dim j As Long 'カウンタ2

Dim move1stTime As Byte '技1のターン数
Dim move1stCharge As Byte '技1のチャージEng
Dim move2ndCharge As Byte '技2の必要Eng
Dim currentEng As Long '現在溜まっているEng
Dim currentTurn As Long '現在のターン数
Dim stackTurn As Long '積み重ねたターン
Dim stackTiming As Long '積み重ねたターン2

Dim countRow As Byte '技回数表の行番号
Dim timingRow As Byte '発動タイミング表の行番号



move1stTime = 2
move1stCharge = 9
move2ndCharge = 45
currentEng = 0

For j = 1 To 15
currentTurn = 0
    For i = 1 To 100
        If currentEng >= move2ndCharge Then
            Exit For
        End If

        currentEng = move1stCharge + currentEng
        currentTurn = move1stTime + currentTurn
    Next i

'Debug.Print "溜まったエネルギー:" & currentEng
'Debug.Print "経過したターン:" & currentTurn

countRow = countRow + 1


Debug.Print "timingRow:"; timingRow

currentEng = currentEng - move2ndCharge
stackTurn = stackTurn + currentTurn
Debug.Print "経過したターン:" & stackTurn
timingRow = stackTurn
'stackTiming = stackTiming + timingRow

With Sheets("技回数表")
    .Cells(17, countRow + 5).Value = currentTurn / move1stTime
End With


With Sheets("発動タイミング表")
    .Cells(17, timingRow + 4).Value = "○"
End With



'Debug.Print "いままでのターン:" & stackTurn


Next j

'Debug.Print "溜まったエネルギー:" & currentEng
'Debug.Print "いままでのターン:" & stackTurn


'    Dim i As Integer
'    For i = 1 To 5
'        Debug.Print i & "回目の繰り返し処理です。"
'
'        If i = 3 Then
'            Exit For
'        End If
'    Next


End Sub

Sub ReferenceVSVallue()
    Dim buf As String
    Dim buf2 As String

    buf = "tanaka"
    buf2 = "tanaka"

    Call Proc1(buf)
    MsgBox "Reference:" & buf
    Call proc2(buf2)
    MsgBox "Vallue:" & buf

End Sub

Sub Proc1(ByRef a As String) 'ByRef:参照渡し
    a = "suzuki"
End Sub

Sub proc2(ByVal a As String) 'ByVal:値渡し
    a = "suzuki"
End Sub

Sub Sample()
    Dim buf As String
    buf = "tanaka"      ''変数に文字列"tanaka"を入れる
    Call Proc1(buf)     ''プロシージャProc1の引数に変数を渡して呼び出す
    MsgBox buf          ''変数の値を表示する
End Sub
