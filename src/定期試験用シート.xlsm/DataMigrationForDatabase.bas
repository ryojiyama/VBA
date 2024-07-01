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
    Set sourceWorkbook = OpenWorkbook("C:\Users\QC07\OneDrive - �g�[���[�Z�t�e�B�z�[���f�B���O�X�������\QC_�����O���t�쐬\", "�O���t�쐬�p�t�@�C��_�ی�X��������p.xlsm")

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

Sub CopyData(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String)
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim IDGenRow As Long
    Dim currentID As String

    ' �]�L���̃V�[�g�̍ŏI�s�ƍŏI����擾
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, "B").End(xlUp).row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column

    ' �]�L��̃V�[�g�̍ŏI�s���擾
    targetLastRow = targetSheet.Cells(targetSheet.Rows.count, "C").End(xlUp).row + 1

    For IDGenRow = 2 To lastRow
        ' �V����ID�𐶐����ē]�L��̃V�[�g��C��ɃZ�b�g
        currentID = GetNewID(targetSheet, IDPrefix)
        targetSheet.Cells(targetLastRow, "C").value = currentID

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

    lastRow = targetSheet.Cells(targetSheet.Rows.count, "C").End(xlUp).row
    If lastRow > 1 Then
        currentID = targetSheet.Cells(lastRow, "C").value
        currentNumber = Val(Mid(currentID, Len(IDPrefix) + 1)) + 1
    Else
        currentNumber = 1
    End If
    GetNewID = IDPrefix & Format(currentNumber, "00000")
End Function
