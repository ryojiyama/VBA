Attribute VB_Name = "TranscribeInspectionReport"
Sub CopyFromExcelToWordBookmark()
    
    On Error GoTo ErrorHandler ' �G���[�n���h�����O
    
    ' Excel�̃V�[�g��ݒ�
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    
    ' Word�A�v���P�[�V�����ƃh�L�������g��ݒ�
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim filePath As String
    filePath = ThisWorkbook.path & "\PeriodicInspectionReport\�l�i�`�U�p�|�O�X�|�P�S�|�O�Q�@�Г��^�����莎���[_AutoTenki.docm"
    
    Set WordApp = New Word.Application
    
    ' Word�t�@�C�������ɊJ���Ă���ꍇ�A����
    Dim docOpen As Boolean
    docOpen = False
    Dim doc As Word.Document
    For Each doc In WordApp.Documents
        If doc.FullName = filePath Then
            doc.Close SaveChanges:=wdSaveChanges
            docOpen = True
            Exit For
        End If
    Next doc
    
    ' Word�t�@�C�����J��
    If docOpen Then
        Set WordDoc = WordApp.Documents.Open(filePath)
    Else
        Set WordDoc = WordApp.Documents.Open(filePath)
    End If
    
    ' �_�C�A���O��ID�����
    Dim ID As String
    ID = InputBox("Enter the ID to process", "ID Input")
    
    ' ID����ɍs������
    Dim rng As Range
    Set rng = ws.Columns("B").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' ID��������Ȃ��ꍇ�A�������I��
    If rng Is Nothing Then
        MsgBox "ID not found"
        Exit Sub
    End If
    
    ' ID�����������s���擾
    Dim i As Long
    i = rng.row
    Dim productNumber As String
    productNumber = ws.Cells(i, "C").Value
    
    With WordDoc
        ' �u�b�N�}�[�N�ɒl��]�L
        If .Bookmarks.Exists("InspectionDate") Then .Bookmarks("InspectionDate").Range.Text = ws.Cells(i, "F").Value
        If .Bookmarks.Exists("ProductNumber") Then .Bookmarks("ProductNumber").Range.Text = productNumber
        If .Bookmarks.Exists("Color") Then .Bookmarks("Color").Range.Text = ws.Cells(i, "N").Value
        If .Bookmarks.Exists("LotNumber") Then .Bookmarks("LotNumber").Range.Text = ws.Cells(i, "O").Value
        If .Bookmarks.Exists("TestContent") Then .Bookmarks("TestContent").Range.Text = ws.Cells(i, "T").Value
        If .Bookmarks.Exists("NaisouLot") Then .Bookmarks("NaisouLot").Range.Text = ws.Cells(i, "Q").Value
        If .Bookmarks.Exists("BoutaiLot") Then .Bookmarks("BoutaiLot").Range.Text = ws.Cells(i, "P").Value
        If .Bookmarks.Exists("Ondo") Then .Bookmarks("Ondo").Range.Text = ws.Cells(i, "G").Value
        If .Bookmarks.Exists("ResultA") Then .Bookmarks("ResultA").Range.Text = ws.Cells(i, "R").Value
        If .Bookmarks.Exists("ResultB") Then .Bookmarks("ResultB").Range.Text = ws.Cells(i, "S").Value
        If .Bookmarks.Exists("Pretreatment") Then .Bookmarks("Pretreatment").Range.Text = ws.Cells(i, "K").Value
        If .Bookmarks.Exists("Weight") Then .Bookmarks("Weight").Range.Text = ws.Cells(i, "L").Value
        If .Bookmarks.Exists("HeadClearance") Then .Bookmarks("HeadClearance").Range.Text = ws.Cells(i, "M").Value
        ' �h�L�������g��ۑ����ĕ���
        .SaveAs filePath & productNumber & .name
        .Close
    End With
    
    ' Word�A�v���P�[�V�������I��
    WordApp.Quit
    
    Exit Sub ' Clean-up �ƃG���[�n���h���̊ԂɈʒu���܂��B

ErrorHandler: ' �G���[�n���h��
    MsgBox "An error has occurred: " & Err.Description
    ' �I�u�W�F�N�g�����
    Set WordDoc = Nothing
    If Not WordApp Is Nothing Then WordApp.Quit
    Set WordApp = Nothing
    Set ws = Nothing
    Set rng = Nothing
End Sub



Sub ExportAllGraphsToWordAsPicture()

    Dim WordApp As Object
    Dim WordDoc As Object
    Dim ExcelWb As Workbook
    Dim ExcelWs As Worksheet
    Dim ExcelChart As ChartObject

    ' Word�A�v���P�[�V�������J�n����
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    ' Word�A�v���P�[�V���������ɂ���
    WordApp.Visible = True

    ' �V����Word�h�L�������g���쐬
    Set WordDoc = WordApp.Documents.Add

    ' Excel�̎w�肳�ꂽ���[�N�u�b�N�ƃ��[�N�V�[�g���J��
    Set ExcelWb = Workbooks.Open("�O���t�쐬�p�t�@�C��.xlsm")
    Set ExcelWs = ExcelWb.Sheets("LOG_Helmet")

    ' �V�[�g���̂��ׂẴO���t���R�s�[����Word�Ƀy�[�X�g
    For Each ExcelChart In ExcelWs.ChartObjects
        ' �O���t�͈̔͂��摜�Ƃ��ăR�s�[
        ExcelChart.chart.CopyPicture Format:=xlPicture
    
        ' Word�̃h�L�������g�̖����ɃJ�[�\�����ړ�
        Dim rng As Object
        Set rng = WordDoc.Content
        rng.Collapse Direction:=wdCollapseEnd  ' �J�[�\���𖖔��Ɉړ�
    
        ' �O���t���摜�Ƃ��ăy�[�X�g
        rng.Paste
        
        ' �y�[�X�g�����摜�̎Q�Ƃ��擾
        Dim InlineShape As Object
        Set InlineShape = WordDoc.InlineShapes(WordDoc.InlineShapes.Count)
        
        ' �摜�̑傫���𒲐�
        InlineShape.LockAspectRatio = True   ' �A�X�y�N�g���ێ�
        InlineShape.Width = 200               ' �����ł̒l�i200�j�͗�Ƃ��Ă��܂��B���ۂ̒l���w�肵�Ă��������B
        
        ' ����ɁA�摜�̃��C�A�E�g�I�v�V�������u�O�ʁv�ɐݒ�
        InlineShape.ConvertToShape.WrapFormat.Type = wdWrapFront
    
        rng.InsertParagraphAfter
    Next ExcelChart

    ' ���ׂẴI�u�W�F�N�g���N���A
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set ExcelChart = Nothing
    Set ExcelWs = Nothing
    Set ExcelWb = Nothing

End Sub


Sub OpenWordTemplate()

    Dim WordApp As Object
    Dim WordDoc As Object
    Dim oneDrivePath As String
    Dim templatePath As String
    
    ' OneDrive�̃��[�J���p�X�����ϐ�����擾
    oneDrivePath = Environ("OneDriveCommercial")

    ' OneDrive�̃p�X�ƕK�v�ȃT�u�t�H���_�E�t�@�C������g�ݍ��킹�ăe���v���[�g�̃p�X�𐶐�
    templatePath = oneDrivePath & "\�i���Ǘ����̏���\�`�F�ی�X\�˗����Q�R�|�ی�X����_�e���v���[�g.docx"

    ' Word�A�v���P�[�V�����̃I�u�W�F�N�g�𐶐��iWord�����ɊJ���Ă���ꍇ�͂�����g�p�j
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    ' Word��\��
    WordApp.Visible = True

    ' �e���v���[�g�t�@�C�����J��
    Set WordDoc = WordApp.Documents.Open(templatePath)

End Sub
