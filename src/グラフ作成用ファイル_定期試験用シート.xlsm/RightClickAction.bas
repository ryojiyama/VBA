Attribute VB_Name = "RightClickAction"
Public Sub CustomizeRightClickMenu()
    On Error GoTo ErrorHandler

    Const MENU_NAME As String = "検査レポート管理"
    Dim Menu As CommandBarPopup
    Dim MenuItem As CommandBarButton

    ' メニュー項目の定義
    Dim menuItems As Variant
    menuItems = Array( _
        Array("試験のレポートを作成", "MakingInspectionSheets", 159), _
        Array("複製したシートを印刷", "PrintFirstPageOfUniqueListedSheets", 5682), _
        Array("複製したシートを削除", "DeleteCopiedSheets", 358) _
    )

    ' 既存のメニューを削除
    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_NAME).Delete
    On Error GoTo ErrorHandler

    ' 新しいメニューを追加
    Set Menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    Menu.Caption = MENU_NAME

    ' メニュー項目を追加
    Dim item As Variant
    For Each item In menuItems
        Set MenuItem = Menu.Controls.Add(Type:=msoControlButton)
        With MenuItem
            .Caption = item(0)
            .OnAction = item(1)
            .FaceId = item(2)
        End With
    Next item

CleanUp:
    Set MenuItem = Nothing
    Set Menu = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "メニューの作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.number & vbCrLf & _
           "エラーの説明: " & Err.Description, _
           vbCritical, "エラー"
    Resume CleanUp
End Sub


Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub




