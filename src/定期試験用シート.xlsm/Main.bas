Attribute VB_Name = "Main"

Public Sub ShowForm()
    Call Form_Helmet.Show
End Sub



Sub MultiplyValues()
    ' 1�s�ڂ�ms�̒l�ɕϊ�����B(1000�������邾��)
    ' ���[�N�V�[�g��ݒ�
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' �J�n��ƏI�����ݒ�
    Dim startColumn As Integer: startColumn = 22 ' V��
    Dim endColumn As Integer: endColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column ' �ŏI��

    ' �񂲂ƂɃ��[�v
    Dim i As Integer
    For i = startColumn To endColumn
        ' �Z���̒l��1000���|����
        ws.Cells(1, i).value = ws.Cells(1, i).value * 1000
    Next i
End Sub






Sub ClickIconAttheTop()
    ' �E�̃V�[�g�Ɉړ�����
    On Error Resume Next
    ActiveSheet.Next.Select
    If Err.number <> 0 Then
        MsgBox "This is the last sheet."
    End If
    On Error GoTo 0
End Sub

Sub ClickUSBIcon()
    'USB�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClickGraphIcon()
    '�O���t�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClickPhotoIcon()
    '�摜�̃A�C�R�����N���b�N����
    UserForm1.Show
End Sub


Sub ClicIconAttheBottom()
    ' ���̃V�[�g�Ɉړ�����
    On Error Resume Next
    ActiveSheet.Previous.Select
    If Err.number <> 0 Then
        MsgBox "This is the first sheet."
    End If
    On Error GoTo 0
End Sub





