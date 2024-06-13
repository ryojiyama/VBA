Attribute VB_Name = "TestModule_Classify"
Sub ClassifyKeys()
    Dim keys As Variant
keys = Array( _
    "Single.110.�V.Cold", _
    "Single.110.�V.Cold", _
    "Multi.110F.�V.Wet", _
    "Multi.110F.�V.Wet", _
    "Multi.110F.�O.Hot", _
    "Single.210.�O.Cold", _
    "Single.210.��.Cold", _
    "Multi.215.�O.Hot", _
    "Multi.215.�V.Hot", _
    "Single.310.��.Wet", _
    "Single.320.�V.Cold", _
    "Single.320.�O.Hot", _
    "Multi.320F.�V.Cold", _
    "Multi.320F.�O.Wet", _
    "Multi.325F.��.Hot", _
    "Multi.325F.��.Hot", _
    "Single.330F.��.Cold", _
    "Single.330F.��.Wet", _
    "Multi.340F.�V.Wet", _
    "Multi.340F.�O.Wet" _
)
    
    ' 1. singleGroups��multiGroups��K�؂ɏ���������
    Dim singleGroups As Scripting.Dictionary
    Set singleGroups = CreateObject("Scripting.Dictionary")
    Dim multiGroups As Scripting.Dictionary
    Set multiGroups = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In keys
        Dim segments As Variant
        segments = Split(key, ".")
        Dim recordType As String
        Dim number As String
        Dim position As String
        Dim condition As String
        recordType = segments(0)
        number = segments(1)
        position = segments(2)
        condition = segments(3)
        
        Dim groupKey As String
        
        ' �ʒu���u���v�̏ꍇ�̓��ʂȈ���
        If position = "��" Then
            groupKey = number & "-" & condition & "-��"
        Else
            groupKey = number & "-" & condition
        End If
        
            ' groupKey�̒l���o��
        Dim tempDict As Scripting.Dictionary
        ' ���R�[�h�^�C�v��SingleValue��MultiValue���ɂ���ď����𕪊�
        If recordType = "Single" Then
            If Not singleGroups.Exists(groupKey) Then
                singleGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = singleGroups(groupKey)
            AddToGroup tempDict, key, position
        ElseIf recordType = "Multi" Then
            If Not multiGroups.Exists(groupKey) Then
                multiGroups.Add groupKey, CreateObject("Scripting.Dictionary")
            End If
            Set tempDict = multiGroups(groupKey)
            AddToGroup tempDict, key, position
        End If
    Next key
    
    ' ���ʂ�\��
    Debug.Print "SingleValue Groups:"
    PrintGroups singleGroups
    
    Debug.Print "MultiValue Groups:"
    PrintGroups multiGroups
End Sub

Sub AddToGroup(ByVal group As Scripting.Dictionary, ByVal key As String, ByVal position As String)
    ' �w�肳�ꂽ�ʒu�Ƀ��X�g�����ɑ��݂��邩�`�F�b�N���A���݂��Ȃ��ꍇ�͐V�������X�g���쐬
    If Not group.Exists(position) Then
        group.Add position, New Collection  ' Collection���g�p���ĕ����̗v�f��ێ�
    End If
    group(position).Add key  ' �ʒu�Ɋւ�炸�L�[�����X�g�ɒǉ�
End Sub

Function GroupHasPosition(group As Scripting.Dictionary, position As String) As Boolean
    ' Check if the position exists in the group
    GroupHasPosition = group.Exists(position)
End Function

Function GroupHasOtherPositions(group As Scripting.Dictionary) As Boolean
    ' Check for the presence of any other position keys in the group
    Dim positions As Variant
    positions = Array("�V", "�O", "��")
    Dim pos As Variant
    For Each pos In positions
        If group.Exists(pos) Then
            GroupHasOtherPositions = True
            Exit Function
        End If
    Next pos
    GroupHasOtherPositions = False
End Function

Sub PrintGroups(ByVal groups As Scripting.Dictionary)
    Dim groupKey As Variant
    For Each groupKey In groups.keys
        Debug.Print "Group " & groupKey & ":"
        Dim position As Variant
        For Each position In groups(groupKey).keys
            Debug.Print "  " & position & ":"
            Dim key As Variant
            For Each key In groups(groupKey)(position)
                Debug.Print "    " & key  ' �e�L�[���o��
            Next key
        Next position
    Next groupKey
End Sub


