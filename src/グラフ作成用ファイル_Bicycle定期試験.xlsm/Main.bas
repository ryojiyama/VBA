Attribute VB_Name = "Main"

Public Sub ShowForm()
    Call Form_Helmet.Show
End Sub






'
'            Array("�����̃��|�[�g�𕡐�", "SetupInspectionReport", 159), _
'            Array("���|�[�g�̗l���𐮂���", "ManageInspectionRecords", 212), _
'            Array("�����f�[�^�̓]�L", "TransferBicycleTestData", 250), _
'            Array("�`���[�g���e�V�[�g�ɔz�u", "ProcessChartDistribution", 620), _


Sub MainProceed()
    Call SetupInspectionReport
    Call ManageInspectionRecords
    Call TransferBicycleTestData
End Sub
