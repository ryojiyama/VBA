Attribute VB_Name = "Main"
Sub Main()

    ' Hel_RequestedTestArrange���W���[������GenereteRequestsID�v���V�[�W�����Ăяo��
    Call GenereteRequestsID

    ' TransferDaraftDatatoSheet���W���[������TransferDataBasedOnID�v���V�[�W�����Ăяo��
    Call TransferDataBasedOnID

    ' TransferDaraftDatatoSheet���W���[������ProcessImpactSheets�v���V�[�W�����Ăяo��
    Call ProcessImpactSheets

    ' ArrangeData���W���[������ArrangeDataByGroup�v���V�[�W�����Ăяo��
    Call ArrangeDataByGroup

    ' ArrangeData���W���[������InsertTextInMergedCells�v���V�[�W�����Ăяo��
    Call InsertTextInMergedCells

    ' ArrangeData���W���[������DistributeChartsToRequestedSheets�v���V�[�W�����Ăяo��
    Call DistributeChartsToRequestedSheets

End Sub
