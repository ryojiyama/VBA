Attribute VB_Name = "Main"
' ���{���ɕR�Â�
Sub ShowFormInspectionType()
    Form_InspectionType.Show
    ' �ȉ���4���܂܂��B
    'CreateGraphHelmet selectedType
    'Call InspectHelmetDurationTime
    'Call AdjustImpactValuesWithCustomFormatForAllLOGSheets
End Sub

Sub ShowFormSpecSheetStylerHelmet()
    Form_SelectedSheetHelmet.Show '�w�����b�g�̃O���t�f�[�^��I�������u�b�N�ɓ]�L
    ' �ȉ���3���܂܂��B
    'Call CreateID(sheetName)
    'Call TransferValuesBetweenSheets
    'Call SyncSpecSheetToLogHel
End Sub

'LOG_Helmet�̃V�[�g��"���", "�^��", "�˗�" �̖ړI�ʂɂ���Ĕz�z����B
Sub ShowSelectedSheetHelmet()
    Dim frm As New Form_SpecSheetStylerHelmet
    frm.Show
End Sub
