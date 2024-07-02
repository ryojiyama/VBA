Attribute VB_Name = "Module1"
Sub ExampleChartNameAndTitle()
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects(1)
    
    ' �O���t�̖��O��ݒ�
    chartObj.Name = "MyCustomChartName"
    
    ' �O���t�̃^�C�g����ݒ�
    If Not chartObj.chart.HasTitle Then
        chartObj.chart.SetElement msoElementChartTitleAboveChart
    End If
    chartObj.chart.chartTitle.Text = "My Chart Title"
End Sub

