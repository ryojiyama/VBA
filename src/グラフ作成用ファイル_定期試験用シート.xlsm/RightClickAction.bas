Attribute VB_Name = "RightClickAction"
Public Sub CustomizeRightClickMenu()
    On Error GoTo ErrorHandler

    Const MENU_NAME As String = "�������|�[�g�Ǘ�"
    Dim Menu As CommandBarPopup
    Dim MenuItem As CommandBarButton

    ' ���j���[���ڂ̒�`
    Dim menuItems As Variant
    menuItems = Array( _
        Array("�����̃��|�[�g���쐬", "MakingInspectionSheets", 159), _
        Array("���������V�[�g�����", "PrintFirstPageOfUniqueListedSheets", 5682), _
        Array("���������V�[�g���폜", "DeleteCopiedSheets", 358) _
    )

    ' �����̃��j���[���폜
    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_NAME).Delete
    On Error GoTo ErrorHandler

    ' �V�������j���[��ǉ�
    Set Menu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    Menu.Caption = MENU_NAME

    ' ���j���[���ڂ�ǉ�
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
    MsgBox "���j���[�̍쐬���ɃG���[���������܂����B" & vbCrLf & _
           "�G���[�ԍ�: " & Err.number & vbCrLf & _
           "�G���[�̐���: " & Err.Description, _
           vbCritical, "�G���["
    Resume CleanUp
End Sub


Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub




