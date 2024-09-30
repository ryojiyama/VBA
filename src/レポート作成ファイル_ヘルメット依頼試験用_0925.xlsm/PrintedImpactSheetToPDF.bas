Attribute VB_Name = "PrintedImpactSheetToPDF"
' "Impact_"ƒV[ƒg‚Ì“à—e‚ğ4‚Â‚ÌƒŒƒR[ƒh‚¸‚Â‡T–‡‚ÌPDF‚Åo—Í‚·‚éB
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
    
    ' ‘Sƒ[ƒNƒV[ƒg‚ğƒ‹[ƒv
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' ƒV[ƒg–¼‚ÉŠî‚Ã‚¢‚Äƒy[ƒWƒwƒbƒ_[‚ğİ’è
            Select Case ws.Name
                Case "Impact_Top"
                    headerText = "“V’¸•”ÕŒ‚Œ±"
                Case "Impact_Front"
                    headerText = "‘O“ª•”ÕŒ‚Œ±"
                Case "Impact_Back"
                    headerText = "Œã“ª•”ÕŒ‚Œ±"
                Case Else
                    headerText = "ÕŒ‚Œ±"
            End Select
            
            ' ƒy[ƒWƒwƒbƒ_[‚Éİ’è
            ws.PageSetup.CenterHeader = headerText
            
            ' ƒOƒ‹[ƒvî•ñ‚ğæ“¾
            Set testResults = CreateObject("Scripting.Dictionary")
            GetGroupInfo ws, testResults
            
            ' ƒOƒ‹[ƒv”‚ğæ“¾
            groupCount = testResults.count
            
            ' ƒOƒ‹[ƒvî•ñ‚ğŠî‚ÉPDF‚ğo—Í
            ApplyColorsAndExportPDF ws, testResults, groupCount, colorArray
        End If
    Next ws
End Sub



Sub GetGroupInfo(ws As Worksheet, testResults As Object)
' GeneratePDFsWithGroupedData‚ÌƒTƒuƒ‹[ƒ`ƒ“Bƒ[ƒNƒV[ƒg‚©‚çƒOƒ‹[ƒvî•ñ‚ğæ“¾‚·‚é
    Dim lastRow As Long
    Dim groupStartRow As Long
    Dim groupNumber As Long
    Dim currentGroup As String
    Dim i As Long
    Dim groupCount As Long
    
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    
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
                
                ' ƒOƒ‹[ƒvî•ñ‚ğDictionary‚É•Û‘¶
                testResults.Add groupCount, Array(groupNumber, groupStartRow)
            End If
        End If
    Next i
End Sub


Sub ApplyColorsAndExportPDF(ws As Worksheet, testResults As Object, groupCount As Long, colorArray As Variant)
    'GeneratePDFsWithGroupedData‚ÌƒTƒuƒ‹[ƒ`ƒ“BƒOƒ‹[ƒvî•ñ‚ğŠî‚ÉPDF‚ğo—Í
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
    
    currentColorIndex = -1
    
    ' ‘Ss‚ğ•\¦ó‘Ô‚É‚·‚é
    ws.Rows.Hidden = False
    
    For i = 1 To groupCount
        groupInfo = testResults(i)
        groupNumber = groupInfo(0)
        groupStartRow = groupInfo(1)
        
        ' Ÿ‚ÌƒOƒ‹[ƒv‚ÌŠJns‚ğæ“¾
        If i < groupCount Then
            lastGroupRow = testResults(i + 1)(1) - 1
        Else
            lastGroupRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
        End If
        
        ' F•ª‚¯‚ÌƒCƒ“ƒfƒbƒNƒX‚ğŒvZ
        colorIndex = (i - 1) \ 4
        If colorIndex > 2 Then colorIndex = 2
        
        ' ƒOƒ‹[ƒv‚ÌŠJns‚ÉF‚ğ•t‚¯‚é
        'ws.Range(ws.Cells(groupStartRow, "A"), ws.Cells(groupStartRow, "G")).Interior.color = colorArray(colorIndex)
        
        ' ‰‰ñ‚Ü‚½‚ÍF‚ª•Ï‚í‚Á‚½ê‡‚Ìˆ—
        If currentColorIndex <> colorIndex Then
            ' ‘O‚ÌF‚ÌƒOƒ‹[ƒv‚ª‚ ‚ê‚ÎPDF‚ğo—Í
            If currentColorIndex <> -1 Then
                ' ˆóü”ÍˆÍ‚ğİ’è
                Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
                ws.PageSetup.printArea = wsRange.Address
                
                ' •s—v‚Ès‚ğ”ñ•\¦‚É‚·‚é
                For j = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
                    If j < firstGroupRow Or j > lastColorGroupRow Then
                        ws.Rows(j).Hidden = True
                    End If
                Next j
                
                ' PDFƒtƒ@ƒCƒ‹–¼‚ğİ’è
                pdfFileName = filePath & "€" & ws.Name & "-" & currentColorIndex & ".pdf"
                
                ' PDF‚ğo—Í
                ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName
                
                ' ”ñ•\¦‚É‚µ‚½s‚ğÄ•\¦
                ws.Rows.Hidden = False
            End If
            
            ' V‚µ‚¢F‚ÌƒOƒ‹[ƒv‚ÌŠJns‚ğİ’è
            firstGroupRow = groupStartRow
            currentColorIndex = colorIndex
        End If
        
        ' Œ»İ‚ÌF‚ÌƒOƒ‹[ƒv‚ÌÅIs‚ğXV
        lastColorGroupRow = lastGroupRow
    Next i
    
    ' ÅŒã‚ÌF‚ÌƒOƒ‹[ƒv‚ğPDFo—Í
    If currentColorIndex <> -1 Then
        Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
        ws.PageSetup.printArea = wsRange.Address
        
        ' •s—v‚Ès‚ğ”ñ•\¦‚É‚·‚é
        For j = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
            If j < firstGroupRow Or j > lastColorGroupRow Then
                ws.Rows(j).Hidden = True
            End If
        Next j
        
        pdfFileName = filePath & "€" & ws.Name & "-" & currentColorIndex & ".pdf"
        
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName
        
        ws.Rows.Hidden = False
    End If
End Sub

' ”–‚­‚¨‚µ‚á‚ê‚ÈF‚ğ•Ô‚·ŠÖ”
Function GetColorArray() As Variant
    GetColorArray = Array(RGB(255, 182, 193), RGB(173, 216, 230), RGB(240, 230, 140)) ' ”–‚¢ƒsƒ“ƒNA”–‚¢ÂA”–‚¢‰©F
End Function





