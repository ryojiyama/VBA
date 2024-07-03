Attribute VB_Name = "Test"


Sub TransferDataBasedOnID()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim nextRow As Long
    Dim data As Collection
    Dim dataItem As Variant
    
    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割
        idParts = Split(wsSource.Cells(i, 3).value, "-")
        If UBound(idParts) >= 2 Then
            ' グループ名（部位）を取得
            group = idParts(2)
            
            ' グループ名に基づいてシート名を設定
            Select Case group
                Case "天"
                    targetSheetName = "Impact_Top"
                Case "前"
                    targetSheetName = "Impact_Front"
                Case "後"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' 対応するグループがない場合はスキップ
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).value
                    GoTo NextIteration
            End Select
            
            ' データをコレクションに追加
            data.Add Array(i, targetSheetName)
        End If
NextIteration:
    Next i
    
    ' コレクションから各シートにデータを転記
    For Each dataItem In data
        Dim rowIndex As Long
        rowIndex = dataItem(0)
        targetSheetName = dataItem(1)
        
        ' 目的のシートを作成
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            wsDest.name = targetSheetName
        End If
        On Error GoTo 0
        
        ' ヘッダー行を設定（B15セルに設定）
        If wsDest.Range("B15").value = "" Then
            wsSource.Range("B1:Z1").Copy Destination:=wsDest.Range("B15")
        End If
        
        ' 最終行を見つけ、次の行からデータの転記を開始します
        nextRow = wsDest.Cells(wsDest.Rows.count, "B").End(xlUp).row + 1
        If nextRow < 16 Then
            nextRow = 16 ' 最初のデータ転記開始位置をB16に設定
        End If
        
        ' データ範囲を転記
        wsSource.Range("B" & rowIndex & ":Z" & rowIndex).Copy Destination:=wsDest.Range("B" & nextRow)
    Next dataItem

    ' リソースを解放
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub


Sub TransferDataBasedOnID_07031500()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    Dim idParts() As String
    Dim group As String
    Dim targetSheetName As String
    Dim groupName As String
    Dim maxValue As Double, duration49kN As Double, duration73kN As Double
    Dim nextRow As Long
    Dim tempArray As Variant
    Dim data As Collection
    Dim dataItem As Variant
    
    ' ソースシートを設定
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Set data = New Collection

    ' ソースシートの最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    ' 各行をループ処理
    For i = 1 To lastRow
        ' IDを分割
        idParts = Split(wsSource.Cells(i, 3).value, "-")
        If UBound(idParts) >= 2 Then
            ' グループ名（部位）を取得
            group = idParts(2)
            
            ' グループ名に基づいてシート名を設定
            Select Case group
                Case "天"
                    targetSheetName = "Impact_Top"
                Case "前"
                    targetSheetName = "Impact_Front"
                Case "後"
                    targetSheetName = "Impact_Back"
                Case Else
                    ' 対応するグループがない場合はスキップ
                    Debug.Print "No matching group for: " & wsSource.Cells(i, 3).value
                    GoTo NextIteration
            End Select
            
            groupName = "Group:" & idParts(0) & group
            maxValue = wsSource.Range("H" & i).value
            duration49kN = wsSource.Range("J" & i).value
            duration73kN = wsSource.Range("K" & i).value

            ' グループ名とシート名の対応を確認
            Debug.Print "Group: " & groupName & "; Sheet: " & targetSheetName
            Debug.Print "Max Value: " & Format(maxValue, "0.00") & " 49kN Duration: " & Format(duration49kN, "0.00") & " 73kN Duration: " & Format(duration73kN, "0.00")

            ' データをコレクションに追加
            tempArray = Array( _
            groupName, _
            targetSheetName, _
            Format(maxValue, "0.00"), _
            Format(duration49kN, "0.00"), _
            Format(duration73kN, "0.00") _
            )
            data.Add tempArray
        End If
NextIteration:
    Next i
    
    ' コレクションから各シートにデータを転記
    For Each dataItem In data
        groupName = dataItem(0)
        targetSheetName = dataItem(1)
        maxValue = dataItem(2)
        duration49kN = dataItem(3)
        duration73kN = dataItem(4)
        ' 目的のシートを作成
        On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(targetSheetName)
        If wsDest Is Nothing Then
            Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            wsDest.name = targetSheetName
        End If
        On Error GoTo 0
        
        ' ヘッダー行を設定（14行目）
        If wsDest.Range("A14").value = "" Then
            wsDest.Range("A14").value = "Group"
            wsDest.Range("B14").value = "Max"
            wsDest.Range("C14").value = "4.9kN"
            wsDest.Range("D14").value = "7.3kN"
        End If
        nextRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row + 1
        If nextRow < 15 Then
            nextRow = 15
        End If
        
        'データを転記
        wsDest.Range("A" & nextRow).value = groupName
        wsDest.Range("B" & nextRow).value = maxValue
        wsDest.Range("C" & nextRow).value = duration49kN
        wsDest.Range("D" & nextRow).value = duration73kN
    Next dataItem

    ' リソースを解放
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub

Sub GroupAndListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim part0 As String
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    ' アクティブシートのチャートオブジェクトをループ処理
    For Each chartObj In ActiveSheet.ChartObjects
        ' グラフにタイトルがあるかどうかをチェック
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"  ' タイトルがない場合
        End If

        ' chartNameを"-"で分割し、part(0)を取得
        part0 = Split(chartObj.name, "-")(0)

        ' グループがまだ存在しない場合、新規作成
        If Not groups.Exists(part0) Then
            groups.Add part0, New Collection
        End If

        ' グループにチャート名とタイトルを追加
        groups(part0).Add "Chart Name: " & chartObj.name & "; Title: " & chartTitle
    Next chartObj

    ' 各グループの内容をイミディエイトウィンドウに出力
    Dim key As Variant
    For Each key In groups.Keys
        Debug.Print "Group: " & key
        Dim item As Variant
        For Each item In groups(key)
            Debug.Print item
        Next item
    Next key
End Sub

Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim sheetName As String
    Dim parts() As String
    Dim groups As Object
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    ' "LOG_Helmet"シートを対象にする
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' "LOG_Helmet"シートのチャートオブジェクトをグループ分け
    For Each chartObj In ws.ChartObjects
        If chartObj.chart.HasTitle Then
            chartTitle = chartObj.chart.chartTitle.text
        Else
            chartTitle = "No Title"
        End If
        
        ' chartNameを"-"で分割し、sheetNameを取得
        parts = Split(chartObj.name, "-")
        If UBound(parts) >= 1 Then
            sheetName = parts(0) & "-" & parts(1)
        Else
            sheetName = parts(0)
        End If
        
        If Not groups.Exists(sheetName) Then
            groups.Add sheetName, New Collection
        End If
        
        groups(sheetName).Add chartObj
    Next chartObj
    
    ' グループごとにチャートを対応するシートに移動
    Dim key As Variant
    For Each key In groups.Keys
        ' シートの存在を確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(key)
        On Error GoTo 0
        
        ' シートが存在しない場合、チャートを移動しない
        If Not targetSheet Is Nothing Then
            Debug.Print "NewSheetName: " & key
            
            ' チャートの移動
            Dim chart As ChartObject
            For Each chart In groups(key)
                chart.chart.Location Where:=xlLocationAsObject, name:=targetSheet.name
            Next chart
            
            Set targetSheet = Nothing
        Else
            Debug.Print "Sheet " & key & " does not exist. Charts not moved."
        End If
    Next key
End Sub



