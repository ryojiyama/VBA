Attribute VB_Name = "Module1"

Sub CopyAndPopulateSheet( _
    sourceSheetName As String, _
    prefix As String, _
    index As Integer, _
    customPropertyName As String, _
    customPropertyValue As String, _
    dataCollection As Collection, _
    writeMethod As String)

    Dim sheetName As String
    Dim ws As Worksheet
    Dim DataSetManager As New DataSetManager

    ' シート名を生成
    sheetName = GenerateSheetName(prefix, index)
    Debug.Print "Generated sheet name: " & sheetName

    ' シートの存在確認と作成
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' コピーするソースシートが存在するか確認
        If Not SheetExists(sourceSheetName) Then
            Debug.Print "Source sheet not found: " & sourceSheetName
            Exit Sub
        End If
        
        On Error Resume Next
        Sheets(sourceSheetName).Copy After:=Sheets(Sheets.Count)
        Set ws = ActiveSheet
        ws.Name = sheetName
        On Error GoTo 0
        
        ' シート名の変更が成功したか確認
        If ws.Name <> sheetName Then
            Debug.Print "Failed to rename the sheet correctly."
            Exit Sub
        End If
    End If

    If Not ws Is Nothing Then
        ' カスタムプロパティの設定
        ws.CustomProperties.Add Name:=customPropertyName, Value:=customPropertyValue

        ' データの転記
        Select Case writeMethod
            Case "WriteSelectedValuesToOutputSheet"
                DataSetManager.WriteSelectedValuesToOutputSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToRstlSheet"
                DataSetManager.WriteSelectedValuesToRstlSheet sourceSheetName, ws.Name, dataCollection
            Case "WriteSelectedValuesToResultTempSheet"
                DataSetManager.WriteSelectedValuesToResultTempSheet ws.Name, dataCollection
            Case Else
                Debug.Print "Unknown write method: " & writeMethod
        End Select
    Else
        Debug.Print "Failed to create or find the sheet: " & sheetName
    End If

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim tmpSheet As Worksheet
    On Error Resume Next
    Set tmpSheet = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not tmpSheet Is Nothing
    On Error GoTo 0
End Function

Function GenerateSheetName(prefix As String, index As Integer) As String
    GenerateSheetName = prefix & Format(index, "00")
End Function





' 実行するプロシージャ：2024-05-30：シートのコピーはできるが転機できてない。
Sub TestSheetCreationAndDataWriting()
    Dim sheetIndex As Integer
    Dim i As Integer '自分で追加
    Dim resultTempIndex As Integer
    Dim testValues As Collection
    Dim Record As Record
    Dim resultTempValues As Collection
    Dim outputValues As Collection
    Dim rstlValues As Collection
    
    ' DataSetManagerの初期化
    Dim DataSetManager As DataSetManager
    Set DataSetManager = New DataSetManager
    DataSetManager.Init
    Debug.Print "records initialized"
    
    ' テストデータの準備
    Set testValues = New Collection
    Set resultTempValues = New Collection
    Set outputValues = New Collection
    Set rstlValues = New Collection
    
    ' サンプルレコードの追加とフィルタリング
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-天", "110F", "天頂", DateValue("2024/5/17"), 29, 3.07
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
        outputValues.Add Record
        rstlValues.Add Record
    End If
    
    Set Record = New Record
    Record.Initialize "02-110-Cold-天", "110", "天頂", DateValue("2024/5/17"), 26, 4.91
    Record.Initialize "03-F110F-Wet-天", "110F", "天頂", DateValue("2024/5/17"), 26, 2.89
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-前", "110F", "前頭部", DateValue("2024/5/17"), 26, 5.25
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-前", "110F", "前頭部", DateValue("2024/5/17"), 29, 5.64
    Else
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Else
    Record.Initialize "01-F110F-Hot-後", "110F", "後頭部", DateValue("2024/5/17"), 26, 5.12
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
        rstlValues.Add Record
    End If
    
    ' インデックス初期化
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-後", "110F", "後頭部", DateValue("2024/5/17"), 29, 5.19
    testValues.Add Record
    Set Record = New Record
    Record.Initialize "01-F110F-Hot-後", "110F", "後頭部", DateValue("2024/5/17"), 26, 5.12
    testValues.Add Record
    outputValues.Add Record
    rstlValues.Add Record
    
    Set Record = New Record
    Record.Initialize "03-F110F-Wet-後", "110F", "後頭部", DateValue("2024/5/17"), 29, 5.19
    testValues.Add Record
    If Record.TemperatureValue = 29 Then
        resultTempValues.Add Record
    Else
        outputValues.Add Record
    sheetIndex = 1
    resultTempIndex = 1
    
    ' OutputSingle/OutputSheet シートの作成とデータの書き込み
    For i = 1 To 5
        CopyAndPopulateSheet "申請_飛来", "申請_飛来_", sheetIndex, "Temp_Shinsei", "申請_飛来", outputValues, "WriteSelectedValuesToOutputSheet"
        CopyAndPopulateSheet "申請_墜落", "申請_墜落_", sheetIndex, "Temp_Shinsei", "申請_墜落", outputValues, "WriteSelectedValuesToOutputSheet"
        sheetIndex = sheetIndex + 1
    Next i
    
    ' Rstl_Single/Rstl_Triple シートの作成とデータの書き込み
    For i = 1 To 5
        CopyAndPopulateSheet "定期_飛来", "定期_飛来_", sheetIndex, "Temp_Teiki", "定期_飛来", rstlValues, "WriteSelectedValuesToRstlSheet"
        CopyAndPopulateSheet "定期_墜落", "定期_墜落_", sheetIndex, "Temp_Teiki", "定期_墜落", rstlValues, "WriteSelectedValuesToRstlSheet"
        sheetIndex = sheetIndex + 1
    Next i
    
    ' Result_Tempシートの作成とデータの書き込み
    CopyAndPopulateSheet "依頼試験", "依頼試験_", resultTempIndex, "Temp_Irai", "依頼試験", resultTempValues, "WriteSelectedValuesToResultTempSheet"
End Sub


