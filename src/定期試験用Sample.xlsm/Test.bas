Attribute VB_Name = "Test"
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



