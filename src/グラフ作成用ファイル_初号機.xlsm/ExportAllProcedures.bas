Attribute VB_Name = "ExportAllProcedures"
Sub ExportAllVBAProcedures()

    Dim VBComp As VBIDE.VBComponent
    Dim SaveToFolder As String
    Dim FolderChosen As Boolean
    Dim ExportFileName As String

    ' �G�N�X�|�[�g��̃t�H���_��I��
    ' "C:\Users\QC07\OneDrive - �g�[���[�Z�t�e�B�z�[���f�B���O�X�������\�i���Ǘ���_�����O���t�쐬\VBA_Log\"���Ƃ肠�����I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        FolderChosen = .Show
        If FolderChosen Then
            SaveToFolder = .SelectedItems(1)
        Else
            Exit Sub '�t�H���_�I�����L�����Z�����ꂽ�ꍇ�͏I��
        End If
    End With

    ' �Q�Ɛݒ�ɁuMicrosoft Visual Basic for Applications Extensibility�v��ǉ�����K�v������܂�
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents

        ' �G�N�X�|�[�g����t�@�C������ݒ�
        Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ExportFileName = VBComp.name & ".bas"
            Case vbext_ct_MSForm
                ExportFileName = VBComp.name & ".frm"
            Case vbext_ct_Document
                ' ThisWorkbook��V�[�g�̃R�[�h��.cls�Ƃ��ăG�N�X�|�[�g�����
                ExportFileName = VBComp.name & ".cls"
        End Select
        
        ' ���ۂɃG�N�X�|�[�g����
        VBComp.Export SaveToFolder & "\" & ExportFileName

    Next VBComp

    MsgBox "VBA�v���V�[�W�������ׂăG�N�X�|�[�g����܂���!", vbInformation

End Sub
