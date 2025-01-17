Attribute VB_Name = "DataMigrate"
'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�FDatabase�t�H���_�̎������ʃf�[�^�x�[�X�t�@�C���Ƀf�[�^��]�L
' �����F�Ȃ�
'*******************************************************************************
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
    Set sourceWorkbook = ActiveWorkbook
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

'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�w�肳�ꂽ�p�X�̃��[�N�u�b�N���J��
' �����FfullPath - �J�����[�N�u�b�N�̊��S�p�X
' �ߒl�F�J����Workbook�I�u�W�F�N�g
'*******************************************************************************
Function OpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    ' Debug.Print "fullPath: " & fullPath
    On Error Resume Next
    Set wb = Workbooks.Open(fullPath)
    On Error GoTo 0

    Set OpenWorkbook = wb
End Function

'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�\�[�X���[�N�u�b�N����^�[�Q�b�g���[�N�u�b�N�փf�[�^��]�L
' �����FsourceWorkbook - �]�L���̃��[�N�u�b�N
'       targetWorkbook - �]�L��̃��[�N�u�b�N
'*******************************************************************************
Sub MigrateData(ByRef sourceWorkbook As Workbook, ByRef targetWorkbook As Workbook)
    'DataMigration_GraphToTestDB_FromGraphbook()�̃T�u�v���V�[�W��
    Dim sourceSheets As Variant
    Dim targetSheets As Variant
    Dim IDPrefixes As Variant
    Dim i As Integer
    Dim sheetExists As Boolean
    
    ' ���̃V�[�g���A�^�[�Q�b�g�V�[�g���AID�v���t�B�b�N�X��z��Ƃ��Đݒ�
    sourceSheets = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall")
    targetSheets = Array("HeLmetTestData", "FallArrestTestData", "biCycleHelmetTestData", "BaseBallTestData")
    IDPrefixes = Array("HBT-", "FAT-", "CHT-", "BBT-")
    
    Debug.Print "�f�[�^�]�L�����J�n: " & Now
    
    ' �z��̊e�v�f�ɑ΂��ăf�[�^�̓]�L���s��
    For i = LBound(sourceSheets) To UBound(sourceSheets)
        ' �\�[�X�V�[�g�̑��݃`�F�b�N
        sheetExists = False
        On Error Resume Next
        sheetExists = Not sourceWorkbook.Sheets(sourceSheets(i)) Is Nothing
        On Error GoTo 0
        
        If Not sheetExists Then
            Debug.Print "�x��: �\�[�X�V�[�g '" & sourceSheets(i) & "' ��������܂���B�X�L�b�v���܂��B"
            GoTo NextIteration
        End If
        
        ' �^�[�Q�b�g�V�[�g�̑��݃`�F�b�N
        sheetExists = False
        On Error Resume Next
        sheetExists = Not targetWorkbook.Sheets(targetSheets(i)) Is Nothing
        On Error GoTo 0
        
        If Not sheetExists Then
            Debug.Print "�x��: �^�[�Q�b�g�V�[�g '" & targetSheets(i) & "' ��������܂���B�X�L�b�v���܂��B"
            GoTo NextIteration
        End If
        
        Dim sourceSheet As Worksheet
        Set sourceSheet = sourceWorkbook.Sheets(sourceSheets(i))
        
        Dim targetSheet As Worksheet
        Set targetSheet = targetWorkbook.Sheets(targetSheets(i))
        
        Debug.Print "������: " & sourceSheets(i) & " �� " & targetSheets(i)
        
        ' �f�[�^�̃R�s�[�����s
        CopyData_CopyPaste sourceSheet, targetSheet, IDPrefixes(i), targetWorkbook
        
NextIteration:
    Next i
    
    Debug.Print "�f�[�^�]�L��������: " & Now
End Sub
'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�w�肳�ꂽ�V�[�g�ԂŃf�[�^���R�s�[���y�[�X�g
' �����FsourceSheet - �]�L���̃��[�N�V�[�g
'       targetSheet - �]�L��̃��[�N�V�[�g
'       IDPrefix - �V�KID�������̃v���t�B�b�N�X
'       targetWorkbook - �]�L��̃��[�N�u�b�N
'*******************************************************************************
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
        Debug.Print "�x��: �V�[�g '" & sourceSheet.Name & "' �Ƀf�[�^������܂���B"
        GoTo ExitHandler
    End If
    lastColumn = sourceSheet.Cells(1, sourceSheet.columns.Count).End(xlToLeft).column
    targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).row

    ' �]�L���̃��R�[�h�����v�Z
    numRecords = lastRow - 1 ' �w�b�_�[�s�����O
    currentID = targetSheet.Cells(targetLastRow, "B").value
    
    Debug.Print "�����J�n: �V�[�g '" & sourceSheet.Name & "'"
    Debug.Print "  - �]�L�����R�[�h��: " & numRecords
    Debug.Print "  - �]�L�J�n�ʒu: " & (targetLastRow + 1)
    Debug.Print "  - �ŏIID: " & currentID

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

    Debug.Print "��������: �V�[�g '" & sourceSheet.Name & "'"
    Debug.Print "  - �V�KID�͈�: " & newIDCollection(1) & " �` " & newIDCollection(numRecords)

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "�G���[���� - �V�[�g '" & sourceSheet.Name & "'"
    Debug.Print "  - �G���[���e: " & Err.Description
    Debug.Print "  - �G���[�R�[�h: " & Err.Number
    Resume ExitHandler
End Sub
'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�w�肳�ꂽ�v���t�B�b�N�X�ŘA�Ԃ̐V�KID�𐶐�
' �����FcurrentID - ���݂̍ŐVID
'       IDPrefix - ID�̃v���t�B�b�N�X
'       numRecords - ��������ID��
' �ߒl�F�������ꂽID��Collection
'*******************************************************************************
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
'*******************************************************************************
' ���C���v���V�[�W��
' �@�\�FLOG�V�[�g���w�肳�ꂽ���i�J�e�S���̃t�H���_���̃��[�N�u�b�N�ɃR�s�[
' �����FselectedButton - �R�s�[����w�肷��{�^����
'*******************************************************************************
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
        Set ws = ActiveWorkbook.Sheets(sheetName)

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
                    Do While sheetExists(copySheetName, destWb)
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

'*******************************************************************************
' �T�u�v���V�[�W��
' �@�\�F�w�肳�ꂽ���[�N�u�b�N���ɃV�[�g�����݂��邩���m�F
' �����FsheetName - �m�F����V�[�g��
'       wb - �m�F�Ώۂ̃��[�N�u�b�N
' �ߒl�F�V�[�g�����݂���ꍇ��True
'*******************************************************************************
Function sheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    sheetExists = Not ws Is Nothing
End Function

