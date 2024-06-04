Sub LINEUP_CALENDAR()

    Dim i As Integer, workDay As Integer
    Dim Nen As Integer, Tsuki As Integer
    Dim Red As Long
    Dim foundCell As Range, firstCell As Range, calArea As Range

    Call InputDate

    Set foundCell = Range("A1:A200").Find(What:=myDate & "月")
    If foundCell Is Nothing Then
        MsgBox "検索に失敗しました"
    Else
        foundCell.Select
    End If

    Nen = Selection.Row
    Tsuki = Selection.Column

    Dim foundFontColor As Range
    With Application.FindFormat
            .Clear
            .Font.ColorIndex = 0
    End With

    Set calArea = Range(Cells(Nen, Tsuki + 1), Cells(Nen + 5, Tsuki + 7))

    Do while i >= 31
    foundFontColor = calArea.Find(What:="*", searchformat:=True)
        If foundFontColor Is Nothing Then
            MsgBox "検索に失敗しました"
        Else
            Set firstCell = foundFontColor
            foundFontColor.Copy Sheets("日報テンプレート").Cells(3, i)
            i = i + 1

        End If
    Loop

    MsgBox Nen & vbCrLf & Tsuki & vbCrLf & workDay & vbCrLf & WorksheetFunction.Sum(calArea) & myDate
    & vbCrLf & firstCell

End Sub
