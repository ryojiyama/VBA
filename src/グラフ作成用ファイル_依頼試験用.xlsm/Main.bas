Attribute VB_Name = "Main"
Sub GenerateTestReportWithGraphs()

    ' Hel_RequestedTestArrange���W���[������GenereteRequestsID�v���V�[�W�����Ăяo��
    Call GenereteRequestsID

    ' TransferDaraftDatatoSheet���W���[������TransferDataBasedOnID�v���V�[�W�����Ăяo��
    Call TransferDataBasedOnID

    ' TransferDaraftDatatoSheet���W���[������ProcessImpactSheets�v���V�[�W�����Ăяo��
    Call ProcessImpactSheets

    ' ArrangeData���W���[������ArrangeDataByGroup�v���V�[�W�����Ăяo��
    Call ArrangeDataByGroup

    ' ArrangeData���W���[������InsertTextInMergedCells�v���V�[�W�����Ăяo��
    ' �쐬����"Impact"�V�[�g�ɒl����
    Call InsertTextInMergedCells

    ' ArrangeData���W���[������DistributeChartsToRequestedSheets�v���V�[�W�����Ăяo��
    ' �`���[�g���e�V�[�g�ɕ��z
    Call DistributeChartsToRequestedSheets
    
    ' TransferToReport���W���[������TransferDataWithMappingAndFormatting�v���V�[�W�����Ăяo��
    ' �u���|�[�g�{���v�V�[�g���̕\�Ɍ��ʂ�}������B
    Call TransferDataWithMappingAndFormatting
    
End Sub
