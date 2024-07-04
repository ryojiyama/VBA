Attribute VB_Name = "ArrangementData"
Option Explicit
' Impact_Topのデータを並べ直す。
Public Sub Rearrangement_Top()
    Dim copyDict As Object
    Dim cellValue As String
    Dim pseudoSpace As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Set copyDict = CreateObject("Scripting.Dictionary")
    
    ' コピー元とコピー先の対応関係を設定
    copyDict.Add "D16", "B2"
    copyDict.Add "H16", "C6"
    copyDict.Add "H17", "C8"
    copyDict.Add "H18", "C10"
    copyDict.Add "H19", "E6"
    copyDict.Add "H20", "E8"
    copyDict.Add "H21", "E10"
    copyDict.Add "H22", "G6"
    copyDict.Add "H23", "G8"
    copyDict.Add "H24", "G10"

    ' B16の値があるかどうかを確認
    If Not IsEmpty(ws.Range("B16").value) Then
        Dim key As Variant
        For Each key In copyDict.Keys
            ws.Range(copyDict(key)).value = ws.Range(key).value
        Next key
    End If

    ' B2セルに値があるかどうかを確認し、値を変更
    If Not IsEmpty(ws.Range("B2").value) Then
        pseudoSpace = ChrW(12288)
        cellValue = ws.Range("B2").value
        
        ' 既に "No." と " AB" が含まれていないか確認
        If Left(cellValue, 3) <> "No." Or Right(cellValue, 3) <> " AB" Then
            ws.Range("B2").value = "No." & cellValue & pseudoSpace & " AB"
        End If
    End If
    
    ' ディクショナリの解放
    Set copyDict = Nothing
End Sub


' Impact_Front, Backのデータを並べ直す。
Public Sub Rearrangement_FrontAndBack()
    Dim copyDict As Object
    Dim cellValue As String
    Dim pseudoSpace As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Set copyDict = CreateObject("Scripting.Dictionary")
    
    ' 品番の設定
    copyDict.Add "D16", "B2"
    ' 衝撃値の設定
    copyDict.Add "H16", "C6"
    copyDict.Add "H17", "C9"
    copyDict.Add "H18", "C12"
    copyDict.Add "H19", "E6"
    copyDict.Add "H20", "E9"
    copyDict.Add "H21", "E12"
    copyDict.Add "H22", "G6"
    copyDict.Add "H23", "G9"
    copyDict.Add "H24", "G12"
    ' 4.9kN時の継続時間の設定
    copyDict.Add "J16", "D6"
    copyDict.Add "J17", "D9"
    copyDict.Add "J18", "D12"
    copyDict.Add "J19", "F6"
    copyDict.Add "J20", "F9"
    copyDict.Add "J21", "F12"
    copyDict.Add "J22", "H6"
    copyDict.Add "J23", "H9"
    copyDict.Add "J24", "H12"
    ' 7.3kN時の継続時間の設定
    copyDict.Add "K16", "D7"
    copyDict.Add "K17", "D10"
    copyDict.Add "K18", "D13"
    copyDict.Add "K19", "F7"
    copyDict.Add "K20", "F10"
    copyDict.Add "K21", "F13"
    copyDict.Add "K22", "H7"
    copyDict.Add "K23", "H10"
    copyDict.Add "K24", "H13"

    ' B16の値があるかどうかを確認
    If Not IsEmpty(ws.Range("B16").value) Then
        Dim key As Variant
        For Each key In copyDict.Keys
            ws.Range(copyDict(key)).value = ws.Range(key).value
        Next key
    End If
    
    ' B2セルに値があるかどうかを確認し、値を変更
    If Not IsEmpty(ws.Range("B2").value) Then
        pseudoSpace = ChrW(12288)
        cellValue = ws.Range("B2").value
        
        ' 既に "No." と " AB" が含まれていないか確認
        If Left(cellValue, 3) <> "No." Or Right(cellValue, 3) <> " AB" Then
            ws.Range("B2").value = "No." & cellValue & pseudoSpace & " AB"
        End If
    End If
    
    ' ディクショナリの解放
    Set copyDict = Nothing
End Sub

'並べ直した後の書式設定
Public Sub ApplyFormattingCondition()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' B16の値があるかどうかを確認
    If Not IsEmpty(ws.Range("B16").value) Then
        ' 範囲A1:H13のフォントを設定
        ws.Range("A1:H13").Font.name = "游明朝"
        
        ' 指定セルのNumberFormatとフォントサイズを設定
        ws.Range("C6, C9, C12, E6, E9, E12, G6, G9, G12").NumberFormat = "0.00 ""kN"""
        ws.Range("C6, C9, C12, E6, E9, E12, G6, G9, G12").Font.size = 10
        
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").NumberFormat = """    4.90kN   ""0.0 ""ms"""
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").Font.size = 8
        ws.Range("D6, D9, D12, F6, F9, F12, H6, H9, H12").HorizontalAlignment = xlLeft
        
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").NumberFormat = """    7.30kN   ""0.0 ""ms"""
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").Font.size = 8
        ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13").HorizontalAlignment = xlLeft
        
        ' 値が0の場合のNumberFormatを設定
        Dim rng As Range
        For Each rng In ws.Range("D7, D10, D13, F7, F10, F13, H7, H10, H13")
            If rng.value = 0 Then
                rng.NumberFormat = """    7.30kN    ― """
            End If
        Next rng
    End If
End Sub


