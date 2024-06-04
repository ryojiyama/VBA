Attribute VB_Name = "DataMigrationForDatabase"
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim oneDrivePath As String
    Dim myPath As String

    ' OneDrive�̃p�X���擾
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\" & "QC_�����O���t�쐬" & "\" & "��������_�f�[�^�x�[�X.xlsm"

    ' sourceWorkbook���J��
    Set sourceWorkbook = OpenWorkbook("C:\Users\QC07\OneDrive - �g�[���[�Z�t�e�B�z�[���f�B���O�X�������\QC_�����O���t�쐬\", "�O���t�쐬�p�t�@�C��.xlsm")

    ' myPath���g�p����targetWorkbook���J��
    Set targetWorkbook = OpenWorkbook(myPath, "")

    ' �]�L����
    MigrateData sourceWorkbook, targetWorkbook

    Application.ScreenUpdating = True
End Sub

Sub MigrateData(ByRef sourceWB As Workbook, ByRef targetWB As Workbook)
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    sourceSheets = Array("LOG_Helmet")
    targetSheets = Array("HeLmetTestData", "BaseBallTestData", "biCycleHelmetTestData", "FallArrestTestData")
    IDPrefixes = Array("HBT-", "BBT-", "CHT-", "FAT-")

    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWB.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWB.Sheets(targetSheets(i))

        CopyData sourceSheet, targetSheet, IDPrefixes(i)
    Next i
End Sub

Sub CopyData_CopyPaste(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim IDGenRow As Long
    Dim currentID As String

    ' �]�L���̃V�[�g�̍ŏI�s�ƍŏI����擾
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).column

    ' �]�L��̃V�[�g�̍ŏI�s���擾
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row + 1

    For IDGenRow = 2 To lastRow
        ' �V����ID�𐶐����ē]�L��̃V�[�g��C��ɃZ�b�g
        currentID = GetNewID(targetSheet, IDPrefix)
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' �]�L������]�L��փf�[�^���R�s�[
        sourceSheet.Range(sourceSheet.Cells(IDGenRow, "C"), sourceSheet.Cells(IDGenRow, "U")).Copy _
            Destination:=targetSheet.Cells(targetLastRow, "D")

        ' D�񂩂�ŏI��܂ł�D�񂩂�ŏI��փR�s�[
        If lastColumn > 4 Then ' 4���葽���ꍇ�̂ݎ��s
            sourceSheet.Range(sourceSheet.Cells(IDGenRow, "D"), sourceSheet.Cells(IDGenRow, lastColumn)).Copy _
                Destination:=targetSheet.Cells(targetLastRow, "E")
        End If

        targetLastRow = targetLastRow + 1
    Next IDGenRow

    ' �]�L�����͈͂��폜�i�f�[�^�݂̂��폜�j
    sourceSheet.Range(sourceSheet.Cells(2, "B"), sourceSheet.Cells(lastRow, lastColumn)).ClearContents

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description & " in " & sourceSheet.name, vbCritical
    Application.ScreenUpdating = True
End Sub
Sub CopyData(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim dataRange As Range
    Dim data As Variant
    Dim targetLastRow As Long
    Dim i As Long
    Dim currentID As String

    ' �]�L���̃V�[�g�̍ŏI�s���擾
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row

    ' �]�L����f�[�^�͈͂�ݒ�i�����ŗ�͈͂�K�X�������Ă��������j
    Set dataRange = sourceSheet.Range("C2:U" & lastRow) ' ��: C�񂩂�U��܂�

    ' �f�[�^�͈͂�z��ɓǂݍ���
    data = dataRange.Value

    ' �]�L��̃V�[�g�̍ŏI�s���擾
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row + 1

    ' �z��̃f�[�^��]�L��ɓ]�L
    For i = LBound(data, 1) To UBound(data, 1)
        ' �V����ID�𐶐�
        currentID = GetNewID(targetSheet, IDPrefix)

        ' ID���Z�b�g
        targetSheet.Cells(targetLastRow, "C").Value = currentID

        ' �z�񂩂�f�[�^��]�L
        Dim j As Long
        For j = LBound(data, 2) To UBound(data, 2)
            targetSheet.Cells(targetLastRow, j + 3).Value = data(i, j) ' D�񂩂�J�n
        Next j

        targetLastRow = targetLastRow + 1
    Next i

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub


Function OpenWorkbook(ByVal path As String, ByVal name As String) As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    If name = "" Then
        fullPath = path
    Else
        fullPath = path & "\" & name
    End If
    Debug.Print "fullPath" & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

Function GetNewID(ByVal targetSheet As Worksheet, ByVal IDPrefix As String) As String
    Dim lastRow As Long
    Dim currentID As String
    Dim currentNumber As Integer

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "C").End(xlUp).row
    If lastRow > 1 Then
        currentID = targetSheet.Cells(lastRow, "C").Value
        currentNumber = Val(Mid(currentID, Len(IDPrefix) + 1)) + 1
    Else
        currentNumber = 1
    End If
    GetNewID = IDPrefix & Format(currentNumber, "00000")
End Function

