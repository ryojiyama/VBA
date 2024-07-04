Attribute VB_Name = "TransferData"
' "LOG_Helmet"のデータを各シートに分配する
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

Sub GroupAndListChartNamesAndTitles()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim partEnd As String
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

        ' chartNameを"-"で分割し、最後の部分を取得
        partEnd = Split(chartObj.name, "-")(UBound(Split(chartObj.name, "-")))

        ' グループがまだ存在しない場合、新規作成
        If Not groups.Exists(partEnd) Then
            groups.Add partEnd, New Collection
        End If

        ' グループにチャート名とタイトルを追加
        groups(partEnd).Add "Chart Name: " & chartObj.name & "; Title: " & chartTitle
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


'Sub GroupAndListChartNamesAndTitles()
'    Dim chartObj As ChartObject
'    Dim chartTitle As String
'    Dim part0 As String
'    Dim groups As Object
'    Set groups = CreateObject("Scripting.Dictionary")
'
'    ' アクティブシートのチャートオブジェクトをループ処理
'    For Each chartObj In ActiveSheet.ChartObjects
'        ' グラフにタイトルがあるかどうかをチェック
'        If chartObj.chart.HasTitle Then
'            chartTitle = chartObj.chart.chartTitle.text
'        Else
'            chartTitle = "No Title"  ' タイトルがない場合
'        End If
'
'        ' chartNameを"-"で分割し、part(0)を取得
'        part0 = Split(chartObj.name, "-")(0)
'
'        ' グループがまだ存在しない場合、新規作成
'        If Not groups.Exists(part0) Then
'            groups.Add part0, New Collection
'        End If
'
'        ' グループにチャート名とタイトルを追加
'        groups(part0).Add "Chart Name: " & chartObj.name & "; Title: " & chartTitle
'    Next chartObj
'
'    ' 各グループの内容をイミディエイトウィンドウに出力
'    Dim key As Variant
'    For Each key In groups.Keys
'        Debug.Print "Group: " & key
'        Dim item As Variant
'        For Each item In groups(key)
'            Debug.Print item
'        Next item
'    Next key
'End Sub

' グラフの並べ直し
Sub DistributeChartsToSheets()
    Dim chartObj As ChartObject
    Dim chartTitle As String
    Dim groupName As String
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
        
        ' chartNameを"-"で分割し、最後の部分を取得
        groupName = Split(chartObj.name, "-")(UBound(Split(chartObj.name, "-")))
        If Not groups.Exists(groupName) Then
            groups.Add groupName, New Collection
        End If
        groups(groupName).Add chartObj
    Next chartObj
    
    ' グループごとにチャートを対応するシートに移動
    Dim key As Variant
    'Dim SheetName As String
    For Each key In groups.Keys
        ' グループ名に基づいてシート名を決定
        Select Case key
            Case "天"
                SheetName = "Impact_Top"
            Case "前"
                SheetName = "Impact_Front"
            Case "後"
                SheetName = "Impact_Back"
            Case Else
                SheetName = "" ' 該当しない場合は空のシート名
        End Select
        
        If SheetName <> "" Then
            ' シートの存在を確認
            On Error Resume Next
            Set targetSheet = ThisWorkbook.Sheets(SheetName)
            On Error GoTo 0
            
            ' シートが存在する場合のみチャートを移動
            If Not targetSheet Is Nothing Then
                ' チャートの移動
                Dim chart As ChartObject
                For Each chart In groups(key)
                    chart.chart.Location Where:=xlLocationAsObject, name:=targetSheet.name
                Next chart
                Set targetSheet = Nothing
            Else
                Debug.Print "Sheet " & SheetName & " does not exist. Charts not moved."
            End If
        Else
            Debug.Print "Group " & key & " does not have a corresponding sheet. Charts not moved."
        End If
    Next key
End Sub

