Attribute VB_Name = "InspectionSheet"
Sub FilterAndGroupDataByF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row

    Dim groupedDataF As Object
    Set groupedDataF = CreateObject("Scripting.Dictionary")
    Dim groupedDataNonF As Object
    Set groupedDataNonF = CreateObject("Scripting.Dictionary")
    Dim copiedSheets As Object
    Set copiedSheets = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = ws.Cells(i, 3).value

        Dim helmetData As New helmetData
        Set helmetData = ParseHelmetData(cellValue)

        Dim productNameKey As String
        productNameKey = helmetData.GroupNumber & "-" & helmetData.ProductName

        If Right(helmetData.ProductName, 1) = "F" Then
            If Not groupedDataF.Exists(helmetData.GroupNumber) Then
                groupedDataF.Add helmetData.GroupNumber, New Collection
            End If
            groupedDataF(helmetData.GroupNumber).Add helmetData

            If helmetData.ImpactPosition = "天" Then
                If Not copiedSheets.Exists(productNameKey) Then
                    ThisWorkbook.Sheets("InspectionSheet").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                    ActiveSheet.name = CreateUniqueName(productNameKey)
                    copiedSheets.Add productNameKey, Nothing
                End If
            End If
        Else
            If Not groupedDataNonF.Exists(helmetData.GroupNumber) Then
                groupedDataNonF.Add helmetData.GroupNumber, New Collection
            End If
            groupedDataNonF(helmetData.GroupNumber).Add helmetData

            If Not copiedSheets.Exists(productNameKey) Then
                ThisWorkbook.Sheets("InspectionSheet").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                ActiveSheet.name = CreateUniqueName(productNameKey)
                copiedSheets.Add productNameKey, Nothing
            End If
        End If
    Next i

    Debug.Print "Data with ProductName ending in 'F':"
    PrintGroupedData groupedDataF
    Debug.Print "Data without ProductName ending in 'F':"
    PrintGroupedData groupedDataNonF
End Sub
Function ParseHelmetData(value As String) As helmetData
    Dim parts() As String
    parts = Split(value, "-")
    Dim result As New helmetData
    
    If UBound(parts) >= 4 Then
        result.GroupNumber = parts(0)
        result.ProductName = parts(1)
        result.ImpactPosition = parts(2)
        result.ImpactTemp = parts(3)
        result.Color = parts(4)
    End If
    
    Set ParseHelmetData = result
End Function

Function CreateUniqueName(baseName As String) As String
    Dim uniqueName As String
    uniqueName = baseName
    Dim count As Integer
    count = 1
    While SheetExists(uniqueName)
        uniqueName = baseName & count
        count = count + 1
    Wend
    CreateUniqueName = uniqueName ' 正しい戻り値の設定
End Function
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sheet Is Nothing ' 正しい戻り値の設定
End Function


Private Sub PrintGroupedData(groupedData As Object)
    Dim key As Variant, item As helmetData
    For Each key In groupedData.Keys
        Debug.Print "GroupNumber: " & key
        For Each item In groupedData(key)
            Debug.Print "  ProductName: " & item.ProductName
            Debug.Print "  ImpactPosition: " & item.ImpactPosition
            Debug.Print "  ImpactTemp: " & item.ImpactTemp
            Debug.Print "  Color: " & item.Color
            Debug.Print "----------------------------"
        Next item
        Debug.Print "============================"
    Next key
End Sub


Sub TransferDataToAppropriateSheets()
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("LOG_Helmet")
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.count, "C").End(xlUp).row

    Dim wsTarget As Worksheet
    Dim i As Long
    Dim productNameKey As String
    Dim dataRange As Range
    Dim targetRow As Long

    ' LOG_Helmetシートの各行をループして処理します
    For i = 2 To lastRow
        ' GroupNumberとProductNameからproductNameKeyを構築します
        productNameKey = wsSource.Cells(i, 3).value
        productNameKey = Split(productNameKey, "-")(0) & "-" & Split(productNameKey, "-")(1)
        
        ' productNameKeyに基づいて対応するシートを検索します
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(productNameKey)
        On Error GoTo 0
        
        ' シートが存在する場合、データを転記します
        If Not wsTarget Is Nothing Then
            ' ターゲットシートにヘッダーを転記する処理
            If wsTarget.Range("B28").value = "" Then ' ヘッダーが未転記であれば転記
                wsSource.Range("B1:Z1").Copy Destination:=wsTarget.Range("B28")
            End If
            
            ' 最終行を見つけ、次の行からデータの転記を開始します
            targetRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row + 1
            If targetRow < 29 Then
                targetRow = 29 ' 最初のデータ転記開始位置をB29に設定
            End If
            Set dataRange = wsSource.Range("B" & i & ":Z" & i)
            dataRange.Copy Destination:=wsTarget.Range("B" & targetRow)
        End If
        
        ' 次のイテレーションのためにwsTargetをリセットします
        Set wsTarget = Nothing
    Next i
End Sub




