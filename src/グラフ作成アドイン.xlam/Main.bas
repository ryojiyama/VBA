Attribute VB_Name = "Main"
' リボンに紐づけ
Sub ShowFormInspectionType()
    Form_InspectionType.Show
    ' 以下の4つが含まれる。
    'CreateGraphHelmet selectedType
    'Call InspectHelmetDurationTime
    'Call AdjustImpactValuesWithCustomFormatForAllLOGSheets
End Sub

Sub ShowFormSpecSheetStylerHelmet()
    Form_SelectedSheetHelmet.Show 'ヘルメットのグラフデータを選択したブックに転記
    ' 以下の3つが含まれる。
    'Call CreateID(sheetName)
    'Call TransferValuesBetweenSheets
    'Call SyncSpecSheetToLogHel
End Sub

'LOG_Helmetのシートを"定期", "型式", "依頼" の目的別によって配布する。
Sub ShowSelectedSheetHelmet()
    Dim frm As New Form_SpecSheetStylerHelmet
    frm.Show
End Sub
