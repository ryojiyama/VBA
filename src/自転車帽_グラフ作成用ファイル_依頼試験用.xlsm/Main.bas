Attribute VB_Name = "Main"
Sub GenerateTestReportWithGraphs()

    ' SpecSheet���W���[������createID�v���V�[�W�����Ăяo��
    Call createID

    ' SpecSheet���W���[������SyncSpecSheetToLOGBicycle�v���V�[�W�����Ăяo��
    ' SpecSheet�ɓ]�L����v���V�[�W���̖{�́B
    Call SyncSpecSheetToLOGBicycle

    ' TransferDraftDatatoSheet���W���[������TransferDataBasedOnID�v���V�[�W�����Ăяo��
    '  "���|�[�g�O���t"�Ȃǂ̃V�[�g���쐬���A�l��]�L����v���V�[�W��
    Call TransferDataBasedOnID
    ' TransferDraftDatatoSheet���W���[������ProcessImpactSheets�v���V�[�W�����Ăяo��
    ' "���|�[�g�O���t"�V�[�g��"�e���v���[�g"�V�[�g����s���R�s�[���đ}��
    Call ProcessImpactSheets
    ' TransferDraftDatatoSheet���W���[������SetCellDimensions�v���V�[�W�����Ăяo��
    ' �s�E��̃T�C�Y�𐮂���
    Call SetCellDimensions

    ' ArrangeToImpactSheet���W���[������ArrangeDataByGroup�v���V�[�W�����Ăяo��
    ' �o���オ����"���|�[�g�O���t"�V�[�g�Ɋe�l��z�u����
    Call ArrangeDataByGroup
    ' ArrangeToImpactSheet���W���[������MoveChartsFromLOGToReport�v���V�[�W�����Ăяo��
    Call MoveChartsFromLOGToReport

    ' TransferToReport���W���[������TransferDataWithMappingAndFormatting�v���V�[�W�����Ăяo��
    ' "���|�[�g�{��"�̕\�Ɍ��ʂ�}������B
    Call TransferDataWithMappingAndFormatting
    ' TransferToReport���W���[������InsertTextToReport�v���V�[�W�����Ăяo��
    ' "���|�[�g�{��"�̕\�����Ƀe�L�X�g��}������B
    Call InsertTextToReport

End Sub
