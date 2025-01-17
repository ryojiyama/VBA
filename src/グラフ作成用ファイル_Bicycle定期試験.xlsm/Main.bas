Attribute VB_Name = "Main"

Public Sub ShowForm()
    Call Form_Helmet.Show
End Sub






'
'            Array("試験のレポートを複製", "SetupInspectionReport", 159), _
'            Array("レポートの様式を整える", "ManageInspectionRecords", 212), _
'            Array("試験データの転記", "TransferBicycleTestData", 250), _
'            Array("チャートを各シートに配置", "ProcessChartDistribution", 620), _


Sub MainProceed()
    Call SetupInspectionReport
    Call ManageInspectionRecords
    Call TransferBicycleTestData
End Sub
