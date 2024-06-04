Attribute VB_Name = "ExportAllProcedures"
Sub ExportAllVBAProcedures()

    Dim VBComp As VBIDE.VBComponent
    Dim SaveToFolder As String
    Dim FolderChosen As Boolean
    Dim ExportFileName As String

    ' エクスポート先のフォルダを選択
    ' "C:\Users\QC07\OneDrive - トーヨーセフティホールディングス株式会社\品質管理部_試験グラフ作成\VBA_Log\"をとりあえず選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        FolderChosen = .Show
        If FolderChosen Then
            SaveToFolder = .SelectedItems(1)
        Else
            Exit Sub 'フォルダ選択がキャンセルされた場合は終了
        End If
    End With

    ' 参照設定に「Microsoft Visual Basic for Applications Extensibility」を追加する必要があります
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents

        ' エクスポートするファイル名を設定
        Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ExportFileName = VBComp.name & ".bas"
            Case vbext_ct_MSForm
                ExportFileName = VBComp.name & ".frm"
            Case vbext_ct_Document
                ' ThisWorkbookやシートのコードは.clsとしてエクスポートされる
                ExportFileName = VBComp.name & ".cls"
        End Select
        
        ' 実際にエクスポートする
        VBComp.Export SaveToFolder & "\" & ExportFileName

    Next VBComp

    MsgBox "VBAプロシージャがすべてエクスポートされました!", vbInformation

End Sub
