Attribute VB_Name = "DataMigrate"
'Database�t�H���_��"��������_�f�[�^�x�[�X.xlsm"�Ɏ����f�[�^��]�L����B
Sub DataMigration_GraphToTestDB_FromGraphbook()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim relativePath As String
    Dim localPath As String
    Dim wb As Workbook
    Dim isOpen As Boolean

    '"OneDriveGraph:C:\Users\QC07\TS�z�[���f�B���O�X�������\OfficeScript�̐��� - �h�L�������g\QC_�O���t�쐬"
    localPath = Environ("OneDriveGraph") ' & "\Database\Database��������_�f�[�^�x�[�X.xlsm"

    ' ���݂̃f�B���N�g������ɑ��΃p�X��ݒ�
    relativePath = localPath & "\Database\��������_�f�[�^�x�[�X.xlsm"
    Set sourceWorkbook = ThisWorkbook
    Set targetWorkbook = Workbooks.Open(relativePath)


    ' ��������_�f�[�^�x�[�X.xlsm�����ɊJ����Ă��邩���m�F
    isOpen = False
    For Each wb In Application.Workbooks
        If wb.FullName = relativePath Then
            Set targetWorkbook = wb
            isOpen = True
            Exit For
        End If
    Next wb

    ' �J����Ă��Ȃ��ꍇ��OpenWorkbook�֐����g�p���ĊJ��
    If Not isOpen Then
        Set targetWorkbook = OpenWorkbook(relativePath)
    End If

    On Error GoTo ErrorHandler

    ' �f�[�^�̓]�L���������s
    MigrateData sourceWorkbook, targetWorkbook
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True
End Sub

Function OpenWorkbook(ByVal fullPath As String) As Workbook
    'DataMigration_GraphToTestDB_FromGraphbook()�̃T�u�v���V�[�W��
    Dim wb As Workbook
    ' Debug.Print "fullPath: " & fullPath

    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

Sub MigrateData(ByRef sourceWorkbook As Workbook, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()�̃T�u�v���V�[�W��
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer

    ' ���̃V�[�g���A�^�[�Q�b�g�V�[�g���AID�v���t�B�b�N�X��z��Ƃ��Đݒ�
    sourceSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    targetSheets = Array("HeLmetTestData", "FallArrestTestData", "biCycleHelmetTestData", "BaseBallTestData")
    IDPrefixes = Array("HBT-", "FAT-", "CHT-", "BBT-")

    ' �z��̊e�v�f�ɑ΂��ăf�[�^�̓]�L���s��
    For i = LBound(sourceSheets) To UBound(sourceSheets)
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWorkbook.Sheets(sourceSheets(i))

        Dim targetSheet As Worksheet
        Set targetSheet = targetWorkbook.Sheets(targetSheets(i))

        ' �f�[�^�̃R�s�[�����s
        CopyData_CopyPaste sourceSheet, targetSheet, IDPrefixes(i), targetWorkbook
    Next i
End Sub

Sub CopyData_CopyPaste(ByRef sourceSheet As Worksheet, ByRef targetSheet As Worksheet, ByVal IDPrefix As String, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()�̃T�u�v���V�[�W��
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim targetLastRow As Long
    Dim currentID As String
    Dim newIDCollection As Collection
    Dim numRecords As Long
    Dim i As Long

    ' �]�L���̃V�[�g�̍ŏI�s�ƍŏI����擾�i�w�b�_�[�s�����O�j
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "�]�L���̃V�[�g " & sourceSheet.Name & " �Ƀf�[�^������܂���B", vbExclamation
        Exit Sub
    End If
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).row

    ' �]�L���̃��R�[�h�����v�Z
    numRecords = lastRow - 1 ' �w�b�_�[�s�����O
    currentID = targetSheet.Cells(targetLastRow, "B").value

    ' �V����ID�𐶐�
    Set newIDCollection = GetNewID(currentID, IDPrefix, numRecords)

    ' �]�L���̃f�[�^�͈͂��R�s�[�i�w�b�_�[�s�����O�j
    sourceSheet.Range(sourceSheet.Cells(2, 1), sourceSheet.Cells(lastRow, lastColumn)).Copy
    ' �V�����f�[�^���y�[�X�g����ꏊ
    targetSheet.Cells(targetLastRow + 1, 1).PasteSpecial Paste:=xlPasteValues

    ' �V����ID��ǉ�
    For i = 1 To numRecords
        targetSheet.Cells(targetLastRow + i, "B").value = newIDCollection(i)
    Next i

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True
End Sub

Function GetNewID(ByVal currentID As String, ByVal IDPrefix As String, ByVal numRecords As Long) As Collection
    'DataMigration_GraphToTestDB_FromGraphbook()�̃T�u�v���V�[�W��
    Dim newIDCollection As Collection
    Set newIDCollection = New Collection

    Dim currentNumber As Long
    Dim i As Long
    Dim idNumberPart As String

    ' �v���t�B�b�N�X����菜���Đ��l�����𒊏o
    idNumberPart = Replace(currentID, IDPrefix, "")
    currentNumber = Val(idNumberPart)

    ' �����̐V����ID�𐶐�
    For i = 1 To numRecords
        currentNumber = currentNumber + 1
        newIDCollection.Add IDPrefix & Format(currentNumber, "00000")
    Next i

    Set GetNewID = newIDCollection
    ' Debug.Print "Generated " & numRecords & " new IDs starting from " & currentNumber - numRecords + 1
End Function
' B��̒l���Q�l��"LOG"�V�[�g�𑼃u�b�N�Ɉړ�����B
Sub CopySheetsToOtherWorkbooks(selectedButton As String)
    Dim sheetNames As Variant
    Dim folderNames As Variant
    Dim sheetName As Variant
    Dim folderName As Variant
    Dim ws As Worksheet
    Dim destWb As Workbook
    Dim destFile As String
    Dim destDir As String
    Dim file As String
    Dim fileCount As Integer
    Dim copySheetName As String
    Dim oneDrivePath As String
    
    Application.ScreenUpdating = False

    ' ���ϐ�����OneDrive�̃p�X���擾
    oneDrivePath = Environ("OneDriveGraph")
    
    ' �ΏۃV�[�g���ƃt�H���_���̃��X�g
    sheetNames = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    folderNames = Array("��Helmet", "��FallArrest", "��Bicycle", "��BaseBall")

    ' �V�[�g���Ƃɏ���
    For i = LBound(sheetNames) To UBound(sheetNames)
        sheetName = sheetNames(i)
        folderName = folderNames(i)
        
        ' �ΏۃV�[�g�̃I�u�W�F�N�g��ݒ�
        Set ws = ThisWorkbook.Sheets(sheetName)

        ' B2�Z�����󂩂ǂ����m�F
        If ws.Range("B2").value <> "" Then
            ' �R�s�[��f�B���N�g����ݒ�
            destDir = oneDrivePath & "\" & folderName & "\"
            Debug.Print "DestDir:" & destDir
            
            ' �R�s�[��t�@�C�������[�v�ŊJ��
            file = Dir(destDir & "*.xls*")
            Do While file <> ""
                ' selectedButton�̓��e�Ɋ�Â��ăt�B���^�����O
                If InStr(file, selectedButton) > 0 Then
                    destFile = destDir & file
                    Set destWb = Workbooks.Open(destFile)
                    
                    ' �A�Ԃ����ăR�s�[
                    fileCount = 1
                    copySheetName = sheetName & "-" & fileCount
                    Do While SheetExists(copySheetName, destWb)
                        fileCount = fileCount + 1
                        copySheetName = sheetName & "-" & fileCount
                    Loop
                    
                    ' �V�[�g���R�s�[
                    ws.Copy After:=destWb.Sheets(destWb.Sheets.Count)
                    destWb.Sheets(destWb.Sheets.Count).Name = copySheetName
                    destWb.Close SaveChanges:=True
                End If
                
                ' ���̃t�@�C����
                file = Dir
            Loop
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

' �V�[�g�����݂��邩�`�F�b�N����֐�
Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

