Attribute VB_Name = "AddedRightClickMenu"
' �Z���̉E�N���b�N���j���[�ɃJ�X�^�����j���[��ǉ�����
Public Sub CustomizeRightClickMenu()


    Const MENU_CAPTION As String = "Custom Menu"

    On Error Resume Next
    Application.CommandBars("Cell").Controls(MENU_CAPTION).Delete
    On Error GoTo 0

    Dim customMenu As CommandBarPopup
    Set customMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
    customMenu.caption = MENU_CAPTION

    AddMenuItem customMenu, "���ʂ��烌�|�[�g�쐬", "GenerateTestReportWithGraphs", 233
    ' AddMenuItem customMenu, "Transfer data to SpecSheet", "TransferProgressToHelmetSpec", 498
    AddMenuItem customMenu, "�O���[�v�����̐F�h��", "ShapingSheets", 512
    AddMenuItem customMenu, "��A�̃��|�[�g�̈��", "PrintedReportSheets", 1764
    AddMenuItem customMenu, "�쐬�����V�[�g�ƃ��|�[�g���̕\�̍폜", "CleanUpSheetsByName", 358
End Sub

Private Sub AddMenuItem(menu As CommandBarPopup, caption As String, onAction As String, faceId As Long)
    ' �w�肳�ꂽ���j���[�Ƀ��j���[���ڂ�ǉ�����

    Dim newItem As CommandBarButton
    Set newItem = menu.Controls.Add(Type:=msoControlButton)
    With newItem
        .caption = caption
        .onAction = onAction
        .faceId = faceId
    End With
End Sub
' �E�N���b�N���j���[���폜����
Public Sub RemoveRightClickMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Custom Menu").Delete
    On Error GoTo 0
End Sub
