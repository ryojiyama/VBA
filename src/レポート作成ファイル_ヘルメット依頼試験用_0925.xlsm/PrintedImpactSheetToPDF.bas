Attribute VB_Name = "PrintedImpactSheetToPDF"
' "Impact_"�V�[�g�̓��e��4�̃��R�[�h���T����PDF�ŏo�͂���B
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
    
    ' �S���[�N�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Impact") > 0 Then
            ' �V�[�g���Ɋ�Â��ăy�[�W�w�b�_�[��ݒ�
            Select Case ws.Name
                Case "Impact_Top"
                    headerText = "�V�����Ռ�����"
                Case "Impact_Front"
                    headerText = "�O�����Ռ�����"
                Case "Impact_Back"
                    headerText = "�㓪���Ռ�����"
                Case Else
                    headerText = "�Ռ�����"
            End Select
            
            ' �y�[�W�w�b�_�[�ɐݒ�
            ws.PageSetup.CenterHeader = headerText
            
            ' �O���[�v�����擾
            Set testResults = CreateObject("Scripting.Dictionary")
            GetGroupInfo ws, testResults
            
            ' �O���[�v�����擾
            groupCount = testResults.count
            
            ' �O���[�v�������PDF���o��
            ApplyColorsAndExportPDF ws, testResults, groupCount, colorArray
        End If
    Next ws
End Sub



Sub GetGroupInfo(ws As Worksheet, testResults As Object)
' GeneratePDFsWithGroupedData�̃T�u���[�`���B���[�N�V�[�g����O���[�v�����擾����
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
                
                ' �O���[�v����Dictionary�ɕۑ�
                testResults.Add groupCount, Array(groupNumber, groupStartRow)
            End If
        End If
    Next i
End Sub


Sub ApplyColorsAndExportPDF(ws As Worksheet, testResults As Object, groupCount As Long, colorArray As Variant)
    'GeneratePDFsWithGroupedData�̃T�u���[�`���B�O���[�v�������PDF���o��
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
    
    ' �S�s��\����Ԃɂ���
    ws.Rows.Hidden = False
    
    For i = 1 To groupCount
        groupInfo = testResults(i)
        groupNumber = groupInfo(0)
        groupStartRow = groupInfo(1)
        
        ' ���̃O���[�v�̊J�n�s���擾
        If i < groupCount Then
            lastGroupRow = testResults(i + 1)(1) - 1
        Else
            lastGroupRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
        End If
        
        ' �F�����̃C���f�b�N�X���v�Z
        colorIndex = (i - 1) \ 4
        If colorIndex > 2 Then colorIndex = 2
        
        ' �O���[�v�̊J�n�s�ɐF��t����
        'ws.Range(ws.Cells(groupStartRow, "A"), ws.Cells(groupStartRow, "G")).Interior.color = colorArray(colorIndex)
        
        ' ����܂��͐F���ς�����ꍇ�̏���
        If currentColorIndex <> colorIndex Then
            ' �O�̐F�̃O���[�v�������PDF���o��
            If currentColorIndex <> -1 Then
                ' ����͈͂�ݒ�
                Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
                ws.PageSetup.printArea = wsRange.Address
                
                ' �s�v�ȍs���\���ɂ���
                For j = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
                    If j < firstGroupRow Or j > lastColorGroupRow Then
                        ws.Rows(j).Hidden = True
                    End If
                Next j
                
                ' PDF�t�@�C������ݒ�
                pdfFileName = filePath & "�" & ws.Name & "-" & currentColorIndex & ".pdf"
                
                ' PDF���o��
                ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName
                
                ' ��\���ɂ����s���ĕ\��
                ws.Rows.Hidden = False
            End If
            
            ' �V�����F�̃O���[�v�̊J�n�s��ݒ�
            firstGroupRow = groupStartRow
            currentColorIndex = colorIndex
        End If
        
        ' ���݂̐F�̃O���[�v�̍ŏI�s���X�V
        lastColorGroupRow = lastGroupRow
    Next i
    
    ' �Ō�̐F�̃O���[�v��PDF�o��
    If currentColorIndex <> -1 Then
        Set wsRange = ws.Range(ws.Cells(firstGroupRow, "A"), ws.Cells(lastColorGroupRow, "G"))
        ws.PageSetup.printArea = wsRange.Address
        
        ' �s�v�ȍs���\���ɂ���
        For j = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
            If j < firstGroupRow Or j > lastColorGroupRow Then
                ws.Rows(j).Hidden = True
            End If
        Next j
        
        pdfFileName = filePath & "�" & ws.Name & "-" & currentColorIndex & ".pdf"
        
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName
        
        ws.Rows.Hidden = False
    End If
End Sub

' �����������ȐF��Ԃ��֐�
Function GetColorArray() As Variant
    GetColorArray = Array(RGB(255, 182, 193), RGB(173, 216, 230), RGB(240, 230, 140)) ' �����s���N�A�����A�������F
End Function





