Attribute VB_Name = "Main"
Sub GenerateTestReportWithGraphs()

    ' SpecSheetモジュール内のcreateIDプロシージャを呼び出す
    Call createID

    ' SpecSheetモジュール内のSyncSpecSheetToLOGBicycleプロシージャを呼び出す
    ' SpecSheetに転記するプロシージャの本体。
    Call SyncSpecSheetToLOGBicycle

    ' TransferDraftDatatoSheetモジュール内のTransferDataBasedOnIDプロシージャを呼び出す
    '  "レポートグラフ"などのシートを作成し、値を転記するプロシージャ
    Call TransferDataBasedOnID
    ' TransferDraftDatatoSheetモジュール内のProcessImpactSheetsプロシージャを呼び出す
    ' "レポートグラフ"シートに"テンプレート"シートから行をコピーして挿入
    Call ProcessImpactSheets
    ' TransferDraftDatatoSheetモジュール内のSetCellDimensionsプロシージャを呼び出す
    ' 行・列のサイズを整える
    Call SetCellDimensions

    ' ArrangeToImpactSheetモジュール内のArrangeDataByGroupプロシージャを呼び出す
    ' 出来上がった"レポートグラフ"シートに各値を配置する
    Call ArrangeDataByGroup
    ' ArrangeToImpactSheetモジュール内のMoveChartsFromLOGToReportプロシージャを呼び出す
    Call MoveChartsFromLOGToReport

    ' TransferToReportモジュール内のTransferDataWithMappingAndFormattingプロシージャを呼び出す
    ' "レポート本文"の表に結果を挿入する。
    Call TransferDataWithMappingAndFormatting
    ' TransferToReportモジュール内のInsertTextToReportプロシージャを呼び出す
    ' "レポート本文"の表下部にテキストを挿入する。
    Call InsertTextToReport

End Sub
