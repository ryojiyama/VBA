VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SpecSheetStylerHelmet 
   Caption         =   "Template Form"
   ClientHeight    =   5136
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   5724
   OleObjectBlob   =   "Form_SpecSheetStylerHelmet.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "Form_SpecSheetStylerHelmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub RunButton_Click()
    Dim sheetName As String
    sheetName = ComboBox_Type.value ' ComboBox_Typeに選択されたシート名を格納

    ' シート名が選択されている場合にCreateIDを実行
    If sheetName <> "" Then
        Call CreateID(sheetName)
        Call TransferValuesBetweenSheets
        Call SyncSpecSheetToLogHel
    Else
        MsgBox "シート名が選択されていません。", vbExclamation
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet

    ' コンボボックスに "Hel_SpecSheet" を含むシート名を追加
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(1, ws.Name, "SpecSheet", vbTextCompare) > 0 Then
            ComboBox_Type.AddItem ws.Name
        End If
    Next ws

End Sub




