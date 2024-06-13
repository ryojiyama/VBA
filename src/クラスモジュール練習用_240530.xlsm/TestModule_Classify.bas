Attribute VB_Name = "TestModule_Classify"
Sub ClassifyKeys()
    Dim keys As Variant
keys = Array( _
    "Single.110.天.Cold", _
    "Single.110.天.Cold", _
    "Multi.110F.天.Wet", _
    "Multi.110F.天.Wet", _
    "Multi.110F.前.Hot", _
    "Single.210.前.Cold", _
    "Single.210.後.Cold", _
    "Multi.215.前.Hot", _
    "Multi.215.天.Hot", _
    "Single.310.側.Wet", _
    "Single.320.天.Cold", _
    "Single.320.前.Hot", _
    "Multi.320F.天.Cold", _
    "Multi.320F.前.Wet", _
    "Multi.325F.後.Hot", _
    "Multi.325F.側.Hot", _
    "Single.330F.後.Cold", _
    "Single.330F.側.Wet", _
    "Multi.340F.天.Wet", _
    "Multi.340F.前.Wet" _
)
    
    ' 1. singleGroupsとmultiGroupsを適切に初期化する
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
        
        ' 位置が「側」の場合の特別な扱い
        If position = "側" Then
            groupKey = number & "-" & condition & "-側"
        Else
            groupKey = number & "-" & condition
        End If
        
            ' groupKeyの値を出力
        Dim tempDict As Scripting.Dictionary
        ' レコードタイプがSingleValueかMultiValueかによって処理を分岐
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
    
    ' 結果を表示
    Debug.Print "SingleValue Groups:"
    PrintGroups singleGroups
    
    Debug.Print "MultiValue Groups:"
    PrintGroups multiGroups
End Sub

Sub AddToGroup(ByVal group As Scripting.Dictionary, ByVal key As String, ByVal position As String)
    ' 指定された位置にリストが既に存在するかチェックし、存在しない場合は新しいリストを作成
    If Not group.Exists(position) Then
        group.Add position, New Collection  ' Collectionを使用して複数の要素を保持
    End If
    group(position).Add key  ' 位置に関わらずキーをリストに追加
End Sub

Function GroupHasPosition(group As Scripting.Dictionary, position As String) As Boolean
    ' Check if the position exists in the group
    GroupHasPosition = group.Exists(position)
End Function

Function GroupHasOtherPositions(group As Scripting.Dictionary) As Boolean
    ' Check for the presence of any other position keys in the group
    Dim positions As Variant
    positions = Array("天", "前", "後")
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
                Debug.Print "    " & key  ' 各キーを出力
            Next key
        Next position
    Next groupKey
End Sub


