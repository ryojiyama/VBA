Attribute VB_Name = "Module1"
Sub ExampleChartNameAndTitle()
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects(1)
    
    ' グラフの名前を設定
    chartObj.Name = "MyCustomChartName"
    
    ' グラフのタイトルを設定
    If Not chartObj.chart.HasTitle Then
        chartObj.chart.SetElement msoElementChartTitleAboveChart
    End If
    chartObj.chart.chartTitle.Text = "My Chart Title"
End Sub

