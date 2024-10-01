Attribute VB_Name = "Hel_RequestedTestArrange"


' 依頼試験用にLOG_Helmetに新しいIDを作成する。
Sub GenereteRequestsID()

    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim id As String

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row

    ' 各行に対してIDを生成
    For i = 2 To lastRow ' 1行目はヘッダと仮定
        id = GenerateID(ws, i)
        ' B列にIDをセット
        ws.Cells(i, 2).value = id
    Next i
End Sub

Function GenerateID(ws As Worksheet, rowIndex As Long) As String
' GenereteRequestsID()のサブプロシージャ
    Dim id As String

    ' C列: 2桁以下の数字
    id = GetColumnCValue(ws.Cells(rowIndex, 3).value)
    id = id & "-" ' C列とD列の間に"-"
    ' D列の処理を変更
    id = id & ExtractNumberWithF(ws.Cells(rowIndex, 4).value)
    id = id & "-"
    id = id & GetColumnEValue(ws.Cells(rowIndex, 5).value) ' E列の条件
    id = id & "-"
    id = id & GetColumnLValue(ws.Cells(rowIndex, 12).value) ' L列の条件

    GenerateID = id
End Function
Function ExtractNumberWithF(value As String) As String
    Dim numPart As String
    Dim hasF As Boolean
    Dim regex As Object
    Dim matches As Object

    ' 正規表現オブジェクトの作成
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{3,6}" ' 1桁以上の数字を抽出
    regex.Global = True

    ' 数字部分を抽出
    Set matches = regex.Execute(value)
    If matches.count > 0 Then
        numPart = matches(0).value ' 最初に見つかった数字を取得
    Else
        numPart = "000000" ' デフォルト値またはエラーハンドリング
    End If

    ' "F"の存在チェック
    hasF = InStr(value, "F") > 0

    ' Fがある場合は数字の後に"F"をつける
    If hasF Then
        ExtractNumberWithF = numPart & "F"
    Else
        ExtractNumberWithF = numPart
    End If
End Function


Function GetColumnCValue(value As Variant) As String
' GenerateIDのサブ関数
    If Len(value) <= 2 Then
        GetColumnCValue = Right("00" & value, 2)
    Else
        GetColumnCValue = "??"
    End If
End Function

Function GetColumnEValue(value As Variant) As String
    ' GenerateIDのサブ関数
    If InStr(value, "天頂") > 0 Then
        GetColumnEValue = "天"
    ElseIf InStr(value, "前頭部") > 0 Then
        GetColumnEValue = "前"
    ElseIf InStr(value, "後頭部") > 0 Then
        GetColumnEValue = "後"
    ElseIf InStr(value, "側面") > 0 Then
        Dim parts() As String
        parts = Split(value, "_")

        If UBound(parts) >= 1 Then
            Dim angle As String
            Dim direction As String

            ' 角度を抽出
            angle = Replace(parts(0), "側面", "")

            ' 方向を抽出と整形
            direction = parts(1)
            direction = Replace(direction, "前", "前")
            direction = Replace(direction, "後", "後")
            direction = Replace(direction, "左", "左")
            direction = Replace(direction, "右", "右")

            GetColumnEValue = "側" & angle & direction
        Else
            GetColumnEValue = "側"
        End If
    Else
        GetColumnEValue = "?"
    End If
End Function

Function GetColumnLValue(value As Variant) As String
' GenerateIDのサブ関数
    Select Case value
        Case "高温"
            GetColumnLValue = "Hot"
        Case "低温"
            GetColumnLValue = "Cold"
        Case "浸せき"
            GetColumnLValue = "Wet"
        Case Else
            GetColumnLValue = "?"
    End Select
End Function

' グループごとに色分け_グループが正確にできているかを確認する
Sub ColorGroups()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentColorIndex As Long
    Dim i As Long
    Dim currentGroup As String
    Dim previousGroup As String
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet") ' シート名を必要に応じて変更
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row
    
    ' 初期設定
    currentColorIndex = 1
    previousGroup = ""
    
    Dim colorArray(1 To 20) As Long
    colorArray(1) = RGB(204, 255, 255) ' Light Cyan
    colorArray(2) = RGB(255, 204, 204) ' Light Red
    colorArray(3) = RGB(204, 255, 204) ' Light Green
    colorArray(4) = RGB(255, 255, 204) ' Light Yellow
    colorArray(5) = RGB(204, 204, 255) ' Light Blue
    colorArray(6) = RGB(255, 229, 204) ' Light Orange
    colorArray(7) = RGB(204, 255, 229) ' Light Aqua
    colorArray(8) = RGB(229, 204, 255) ' Light Purple
    colorArray(9) = RGB(255, 204, 229) ' Light Pink
    colorArray(10) = RGB(255, 255, 153) ' Light Yellow 2
    colorArray(11) = RGB(204, 255, 153) ' Light Lime
    colorArray(12) = RGB(153, 204, 255) ' Light Sky Blue
    colorArray(13) = RGB(255, 204, 153) ' Light Peach
    colorArray(14) = RGB(204, 153, 255) ' Light Lavender
    colorArray(15) = RGB(255, 153, 204) ' Light Rose
    colorArray(16) = RGB(204, 255, 255) ' Light Mint
    colorArray(17) = RGB(255, 255, 204) ' Light Cream
    colorArray(18) = RGB(204, 229, 255) ' Light Denim
    colorArray(19) = RGB(255, 204, 255) ' Light Fuchsia
    colorArray(20) = RGB(255, 204, 229) ' Light Rose 2

    ' グループごとに色をつける
    For i = 2 To lastRow
        currentGroup = ws.Cells(i, 3).value
        
        ' グループが変わったら次の色に切り替え
        If currentGroup <> previousGroup Then
            currentColorIndex = currentColorIndex + 1
            If currentColorIndex > UBound(colorArray) Then
                currentColorIndex = 1
            End If
        End If
        
        ' B列からE列に色を設定
        ws.Range(ws.Cells(i, 2), ws.Cells(i, 5)).Interior.color = colorArray(currentColorIndex)
        
        ' 現在のグループを記録
        previousGroup = currentGroup
    Next i
End Sub

