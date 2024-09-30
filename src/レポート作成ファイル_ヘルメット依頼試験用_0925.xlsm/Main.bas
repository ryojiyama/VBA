Attribute VB_Name = "Main"
Sub Main()

    ' Hel_RequestedTestArrangeモジュール内のGenereteRequestsIDプロシージャを呼び出す
    Call GenereteRequestsID

    ' TransferDaraftDatatoSheetモジュール内のTransferDataBasedOnIDプロシージャを呼び出す
    Call TransferDataBasedOnID

    ' TransferDaraftDatatoSheetモジュール内のProcessImpactSheetsプロシージャを呼び出す
    Call ProcessImpactSheets

    ' ArrangeDataモジュール内のArrangeDataByGroupプロシージャを呼び出す
    Call ArrangeDataByGroup

    ' ArrangeDataモジュール内のInsertTextInMergedCellsプロシージャを呼び出す
    Call InsertTextInMergedCells

    ' ArrangeDataモジュール内のDistributeChartsToRequestedSheetsプロシージャを呼び出す
    Call DistributeChartsToRequestedSheets

End Sub
