Attribute VB_Name = "RightClickAction"
' セルの右クリックメニューにカスタムメニューを追加する
Public Sub CustomizeRightClickMenu_IraiBicycle()


    Const MENU_CAPTION As String = "Custom Menu"

    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_CAPTION).Delete
    On Error GoTo 0

    Dim customMenu As CommandBarPopup
    Set customMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    customMenu.caption = MENU_CAPTION

    AddMenuItem customMenu, "結果からレポート作成", "GenerateTestReportWithGraphs", 233
    AddMenuItem customMenu, "ヘッダーの追加", "AddHeader", 512
    AddMenuItem customMenu, "一連のレポートの印刷", "PrintSheet", 1764
    AddMenuItem customMenu, "作成したシートとレポート内の表の削除", "DeleteReport", 358
End Sub

Private Sub AddMenuItem(menu As CommandBarPopup, caption As String, onAction As String, faceId As Long)
    ' 指定されたメニューにメニュー項目を追加する

    Dim newItem As CommandBarButton
    Set newItem = menu.Controls.Add(Type:=msoControlButton)
    With newItem
        .caption = caption
        .onAction = onAction
        .faceId = faceId
    End With
End Sub
' 右クリックメニューを削除する
Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub

