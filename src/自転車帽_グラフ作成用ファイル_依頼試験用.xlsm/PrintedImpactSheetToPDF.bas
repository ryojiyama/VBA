Attribute VB_Name = "PrintedImpactSheetToPDF"
' "レポートグラフ"シートの内容を4つのレコードずつ�T枚のPDFで出力する。ヘルメットのものをそのままコピー修正が必要
Sub GeneratePDFsWithGroupedData()
    Dim ws As Worksheet
    Dim testResults As Object
    Dim colorArray As Variant
    Dim lastRow As Long
    Dim groupCount As Long
    Dim groupNumber As Long
    Dim groupStartRow As Long
    Dim groupInfo As Variant
    Dim pdfFileName As String
    Dim wsRange As Range
    Dim i As Long
    Dim headerText As String
    
    ' 全ワークシートをループ
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' シート名に基づいてページヘッダーを設定
            Select Case ws.Name
                Case "Impact_Top"
                    headerText = "天頂部衝撃試験"
                Case "Impact_Front"
                    headerText = "前頭部衝撃試験"
                Case "Impact_Back"
                    headerText = "後頭部衝撃試験"
                Case Else
                    headerText = "衝撃試験"
            End Select
            
            ' ページヘッダーに設定
            ws.PageSetup.CenterHeader = headerText
            
            ' グループ情報を取得
            Set testResults = CreateObject("Scripting.Dictionary")
            GetGroupInfo ws, testResults
            
            ' グループ数を取得
            groupCount = testResults.Count
            Debug.Print "グループ数：" & testResults.Count
            
            ' グループ情報を基にPDFを出力
            ApplyColorsAndExportPDF ws, testResults, groupCount, colorArray
        End If
    Next ws
End Sub



Sub GetGroupInfo(ws As Worksheet, testResults As Object)
' GeneratePDFsWithGroupedDataのサブルーチン。ワークシートからグループ情報を取得する
    Dim lastRow As Long
    Dim groupStartRow As Long
    Dim groupNumber As Long
    Dim currentGroup As String
    Dim i As Long
    Dim groupCount As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    
    groupCount = 0
    currentGroup = ""
    groupStartRow = 0
    
    For i = 2 To lastRow
        If ws.Cells(i, "I").value Like "Insert*" Then
            If ws.Cells(i, "I").value <> currentGroup Then
                groupCount = groupCount + 1
                currentGroup = ws.Cells(i, "I").value
                groupStartRow = i
                groupNumber = Val(Mid(currentGroup, 7))
                
                ' グループ情報をDictionaryに保存
                testResults.Add groupCount, Array(groupNumber, groupStartRow)
            End If
        End If
    Next i
End Sub


Sub ApplyColorsAndExportPDF(ws As Worksheet, testResults As Object, groupCount As Long, colorArray As Variant)
    'GeneratePDFsWithGroupedDataのサブルーチン。グループ情報を基にPDFを出力
    Dim i As Long
    Dim groupInfo As Variant
    Dim groupNumber As Long
    Dim groupStartRow As Long
    Dim lastGroupRow As Long
    Dim colorIndex As Long
    Dim pdfFileName As String
    Dim firstGroupRow As Long
    Dim currentColorIndex As Long
    Dim wsRange As Range
    Dim filePath As String
    Dim lastColorGroupRow As Long
    Dim j As Long
    
    filePath = ThisWorkbook.Path
    Debug.Print "FilePath:" & filePath
    
    currentColorIndex = -1
    
    ' 全行を表示状態にする
    ws.Rows.Hidden = False
    
    For i = 1 To groupCount
        groupInfo = testResults(i)
        groupNumber = groupInfo(0)
        groupStartRow = groupInfo(1)
        
        ' 次のグループの開始行を取得
        If i < groupCount Then
            lastGroupRow = testResults(i + 1)(1) - 1
        Else
            lastGroupRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        End If
        
        ' 色分けのインデックスを計算
        colorIndex = (i - 1) \ 4
        If colorIndex > 2 Then colorIndex = 2
        
        ' グループの開始行に色を付ける
        'ws.Range(ws.Cells(groupStartRow, "A"), ws.Cells(groupStartRow, "G")).Interior.color = colorArray(colorIndex)
        
        ' 初回または色が変わった場合の処理
        If currentColorIndex <> colorIndex Then
            ' 前の色のグループがあればPDFを出力
            If currentColorIndex <> -1 Then
                ' 印刷範囲を設定
                Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
                ws.PageSetup.PrintArea = wsRange.Address
                
                ' 不要な行を非表示にする
                For j = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
                    If j < firstGroupRow Or j > lastColorGroupRow Then
                        ws.Rows(j).Hidden = True
                    End If
                Next j
                
                ' PDFファイル名を設定
                pdfFileName = filePath & "�" & ws.Name & "-" & currentColorIndex & ".pdf"
                
                ' PDFを出力
                ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFileName
                
                ' 非表示にした行を再表示
                ws.Rows.Hidden = False
            End If
            
            ' 新しい色のグループの開始行を設定
            firstGroupRow = groupStartRow
            currentColorIndex = colorIndex
            Debug.Print "CurrentColorIndex:" & currentColorIndex
        End If
        
        ' 現在の色のグループの最終行を更新
        lastColorGroupRow = lastGroupRow
    Next i
    
    ' 最後の色のグループをPDF出力
    If currentColorIndex <> -1 Then
        Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
        ws.PageSetup.PrintArea = wsRange.Address
        
        ' 不要な行を非表示にする
        For j = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
            If j < firstGroupRow Or j > lastColorGroupRow Then
                ws.Rows(j).Hidden = True
            End If
        Next j
        
        pdfFileName = filePath & "\" & ws.Name & "-" & currentColorIndex & ".pdf"
        
        Debug.Print "Exporting PDF: " & pdfFileName
        ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFileName
        Debug.Print "Export Complete"

        ws.Rows.Hidden = False
    End If
End Sub

' 薄くおしゃれな色を返す関数
Function GetColorArray() As Variant
    GetColorArray = Array(RGB(255, 182, 193), RGB(173, 216, 230), RGB(240, 230, 140)) ' 薄いピンク、薄い青、薄い黄色
End Function






