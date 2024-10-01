Attribute VB_Name = "Main"
Sub GenerateTestReportWithGraphs()

    ' Hel_RequestedTestArrangeモジュール内のGenereteRequestsIDプロシージャを呼び出す
    Call GenereteRequestsID

    ' TransferDaraftDatatoSheetモジュール内のTransferDataBasedOnIDプロシージャを呼び出す
    Call TransferDataBasedOnID

    ' TransferDaraftDatatoSheetモジュール内のProcessImpactSheetsプロシージャを呼び出す
    Call ProcessImpactSheets

    ' ArrangeDataモジュール内のArrangeDataByGroupプロシージャを呼び出す
    Call ArrangeDataByGroup

    ' ArrangeDataモジュール内のInsertTextInMergedCellsプロシージャを呼び出す
    ' 作成した"Impact"シートに値を代入
    Call InsertTextInMergedCells

    ' ArrangeDataモジュール内のDistributeChartsToRequestedSheetsプロシージャを呼び出す
    ' チャートを各シートに分配
    Call DistributeChartsToRequestedSheets
    
    ' TransferToReportモジュール内のTransferDataWithMappingAndFormattingプロシージャを呼び出す
    ' 「レポート本文」シート内の表に結果を挿入する。
    Call TransferDataWithMappingAndFormatting
    
End Sub
