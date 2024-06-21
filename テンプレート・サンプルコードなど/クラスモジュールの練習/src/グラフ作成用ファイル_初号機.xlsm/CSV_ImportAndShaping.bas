Attribute VB_Name = "CSV_ImportAndShaping"
Public DEFAULT_SHEET_ORDER() As Variant

Sub ImportCSV()
    Call ImportCSVsAndSortSheets
    Call Shapig_CSVData
    ' �J���Ă���u�b�N�̈�ԍ��̃V�[�g��I��
    ThisWorkbook.Sheets(1).Select

    ' A1�Z���ɃJ�[�\�����ړ�
    Range("A1").Select
    
    ' �I�����b�Z�[�W��\��
    MsgBox "CSV�𐳏�ɓǂݍ��݂܂����B", vbInformation, "Operation Complete"
End Sub





Sub ImportCSVsAndSortSheets()
    ' DEFAULT_SHEET_ORDER�̏������𒼐ڍs��
    ReDim DEFAULT_SHEET_ORDER(0 To 5) '�z��̃T�C�Y���w��
    DEFAULT_SHEET_ORDER = Array("LOG_Helmet", "LOG_FallArrest", "LOG_Bicycle", "LOG_BaseBall", "Setting", "Hel_SpecSheet")

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oneDrivePath As String
    Dim myPath As String
    Dim myFile As String
    Dim csvFile As String
    Dim sheetNames As Collection
    Dim idx As Integer

    Set wb = ThisWorkbook
    
    ' OneDrive�̃��[�J���p�X�����ϐ�����擾
    oneDrivePath = Environ("OneDriveCommercial")
    myPath = oneDrivePath & "\QC_�����O���t�쐬\CSV\"
    
    Set sheetNames = New Collection
    
    myFile = Dir(myPath & "*.csv")

    Do While myFile <> ""
        csvFile = myPath & myFile
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = Left(myFile, InStr(myFile, ".") - 1)
        ImportCSVToSheet ws, csvFile
        sheetNames.Add ws.name
        myFile = Dir()
    Loop

    Application.ScreenUpdating = False
    SortSheetsByOrder wb, sheetNames, DEFAULT_SHEET_ORDER
    Application.ScreenUpdating = True
End Sub

Sub ImportCSVToSheet(ByRef ws As Worksheet, ByVal csvFile As String)
    With ws.QueryTables.Add(Connection:="TEXT;" & csvFile, Destination:=ws.Range("A1"))
        .FieldNames = True
        .RefreshOnFileOpen = False
        .TextFilePlatform = 932
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileTabDelimiter = True
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub SortSheetsByOrder(ByRef wb As Workbook, ByVal sheetNames As Collection, ByVal defaultOrder As Variant)
    ' Debug
    Debug.Print "Start of SortSheetsByOrder"
    Debug.Print "defaultOrder type: " & TypeName(defaultOrder)
    ' Debug
    Dim sheetOrder() As String
    Dim i As Integer

    ReDim sheetOrder(sheetNames.Count - 1)
    For i = 1 To sheetNames.Count
        sheetOrder(i - 1) = sheetNames(i)
    Next i

    Call BubbleSort(sheetOrder)

    For i = 1 To UBound(sheetOrder) + 1
        Sheets(sheetOrder(i - 1)).Move After:=Sheets(wb.Sheets.Count)
    Next i

    For i = LBound(defaultOrder) To UBound(defaultOrder)
        Sheets(defaultOrder(i)).Move Before:=Sheets(i + 1)
    Next i
End Sub

Sub BubbleSort(arr As Variant)
    Dim strTemp As String
    Dim i As Integer, j As Integer
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CustomCompare(arr(i), arr(j)) > 0 Then
                strTemp = arr(i)
                arr(i) = arr(j)
                arr(j) = strTemp
            End If
        Next j
    Next i
End Sub

Function CustomCompare(ByVal str1 As String, ByVal str2 As String) As Integer
    Dim numPart1 As String, numPart2 As String
    Dim restPart1 As String, restPart2 As String
    
    ' ���������Ǝc��̕����𕪊�
    numPart1 = Left(str1, 4)
    numPart2 = Left(str2, 4)
    restPart1 = Mid(str1, 5)
    restPart2 = Mid(str2, 5)
    
    ' �ŏ��ɖ����̕������r
    If restPart1 < restPart2 Then
        CustomCompare = 1
    ElseIf restPart1 > restPart2 Then
        CustomCompare = -1
    Else
        ' �����̕����������ꍇ�A�����������t�̏����Ŕ�r
        If numPart1 < numPart2 Then
            CustomCompare = 1
        ElseIf numPart1 > numPart2 Then
            CustomCompare = -1
        Else
            CustomCompare = 0
        End If
    End If
End Function



Sub ImportCSVsAndSortSheets_0926Before()
    ' ���̃t�@�C���Ɠ����f�B���N�g������CSV�t�H���_�Ɋi�[����Ă���CSV�t�@�C����ǂݍ���
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim myPath As String
    Dim myFile As String
    Dim csvFile As String
    Dim sheetNames As Collection
    Dim i As Integer

    Set wb = ThisWorkbook
    myPath = ThisWorkbook.path & "\CSV\"   ' Path changed to subfolder 'CSV'

    Set sheetNames = New Collection

    myFile = Dir(myPath & "*.csv")   ' get the first CSV file

    i = 1
    Do While myFile <> ""
        csvFile = myPath & myFile

        ' Create a new worksheet with the name of the file (without extension)
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = Left(myFile, InStr(myFile, ".") - 1)

        ' Import the CSV file into the new worksheet
        With ws.QueryTables.Add(Connection:="TEXT;" & csvFile, Destination:=ws.Range("A1"))
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 932
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With

        sheetNames.Add ws.name
        myFile = Dir()    ' get the next CSV file
        i = i + 1
    Loop

    ' Sort sheets
    Application.ScreenUpdating = False
    Dim sheetOrder As Variant
    sheetOrder = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest", "Setting", "SpecSheet")

    Dim X As Long
    For X = UBound(sheetOrder) To LBound(sheetOrder) Step -1
        Sheets(sheetOrder(X)).Move Before:=Sheets(1)
    Next X

    For X = 1 To sheetNames.Count
        Sheets(sheetNames(X)).Move After:=Sheets(5 + sheetNames.Count - X)
    Next X

    Application.ScreenUpdating = True
End Sub





Sub Shapig_CSVData()
    '�ǂݍ���CSV�t�@�C���𐮌`���A���ꂼ��̃V�[�g�ɕ��ג����܂��B
    Dim ws As Worksheet
    Dim logSheet As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    Dim dataRange As Range
    Dim targetRange As Range
    Dim lastColumn As Long

    ' ���[�N�u�b�N���̃V�[�g���t���ɏ������܂��B'Setting'��'LOG'�V�[�g�͖������܂��B
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)

            ' �V�[�g����"Setting"�ƈقȂ�A"LOG"���܂܂��A"SpecSheet"���܂܂Ȃ��V�[�g�ɑ΂��Ă̂ݏ������s��
            If ws.name <> "Setting" And InStr(ws.name, "LOG") = 0 And InStr(ws.name, "SpecSheet") = 0 Then

            ' �V�[�g���ɂ���āA���O�V�[�g��ύX���܂�
            If InStr(UCase(ws.name), "HEL") > 0 Then
                Set logSheet = ThisWorkbook.Sheets("LOG_Helmet")
            ElseIf InStr(UCase(ws.name), "BASEBALL") > 0 Then
                Set logSheet = ThisWorkbook.Sheets("LOG_BaseBall")
            ElseIf InStr(UCase(ws.name), "BICYCLE") > 0 Then
                Set logSheet = ThisWorkbook.Sheets("LOG_Bicycle")
            ElseIf InStr(UCase(ws.name), "FALLARR") > 0 Then
                Set logSheet = ThisWorkbook.Sheets("LOG_FallArrest")
            Else
                ' Skip this sheet if it does not match any criteria
                GoTo NextSheet
            End If

            ' �������̃V�[�g����LOG�V�[�g�̍Ō�̍s�ɒǉ����܂��B
            lastRow = logSheet.Cells(logSheet.Rows.Count, "B").End(xlUp).row + 1
            logSheet.Cells(lastRow, "B").Value = ws.name

            ' �������̃V�[�g����f�[�^���R�s�[���܂��B
            ws.Range("A3:D3").Copy
            logSheet.Cells(lastRow, "D").PasteSpecial xlPasteAll

            ws.Range("A6:I6").Copy
            logSheet.Cells(lastRow, "G").PasteSpecial xlPasteAll

            ' B�񂩂�9�s�ڂ܂ł̓��e���ƍs��ϊ�����O�񂩂���ג����܂��B
            lastRowInWs = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            Set dataRange = ws.Range("B9:B" & lastRowInWs)
            Set targetRange = logSheet.Cells(lastRow, "P")

            dataRange.Copy
            targetRange.PasteSpecial Paste:=xlPasteAll, Transpose:=True

            ' �\��t�����f�[�^�̍ŏI��������܂��B
            lastColumn = logSheet.Cells(lastRow, logSheet.Columns.Count).End(xlToLeft).column

            ' ���l��W���`���ŕ\�����܂��B
            logSheet.Range(logSheet.Cells(lastRow, "P"), logSheet.Cells(lastRow, lastColumn)).NumberFormat = "0.0000"
        
            ' ���O�V�[�g��G�񂩂�O��܂ł̃f�[�^���폜���܂��B
            logSheet.Range("G2:O" & logSheet.Rows.Count).ClearContents
        End If

NextSheet:
    Next i

    ' �R�s�[���[�h���I�����܂��B
    Application.CutCopyMode = False
End Sub



