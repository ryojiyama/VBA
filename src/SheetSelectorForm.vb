'*******************************************************************************
' モジュール名：SheetSelector
' 目的：TestDataシートの選択機能を提供する
' 作成：2024/12/27
'*******************************************************************************
Option Explicit

'定数定義
Private Const SHEET_KEYWORD As String = "TestData"
Private mSelectedSheet As String

'*******************************************************************************
' GetTestDataSheets
' 概要：TestDataを含むシート名の配列を取得する
' 戻値：Variant - シート名の配列、シートがない場合はEmpty
'*******************************************************************************
Private Function GetTestDataSheets() As Variant
    Dim ws As Worksheet
    Dim sheetNames As New Collection

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, SHEET_KEYWORD, vbTextCompare) > 0 Then
            sheetNames.Add ws.Name
        End If
    Next ws

    If sheetNames.Count = 0 Then
        GetTestDataSheets = Empty
        Exit Function
    End If

    ' コレクションを配列に変換
    Dim result() As String
    ReDim result(1 To sheetNames.Count)

    Dim i As Long
    For i = 1 To sheetNames.Count
        result(i) = sheetNames(i)
    Next i

    GetTestDataSheets = result
End Function

'*******************************************************************************
' ShowSheetSelector
' 概要：シート選択フォームを表示する
' 戻値：String - 選択されたシート名、キャンセル時は空文字
'*******************************************************************************
Public Function ShowSheetSelector() As String
    Dim sheets As Variant
    sheets = GetTestDataSheets()

    If IsEmpty(sheets) Then
        MsgBox "TestDataを含むシートが見つかりません。", vbExclamation
        ShowSheetSelector = ""
        Exit Function
    End If

    ' UserFormの作成と表示
    With SheetSelectorForm
        .Initialize sheets
        .Show
        ShowSheetSelector = .SelectedSheet
    End With
End Function

' UserFormのコード（SheetSelectorForm）
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetSelectorForm
   Caption         =   "データシートの選択"
   ClientHeight    =   2280
   ClientWidth     =   4560
   OleObjectBlob   =   "SheetSelectorForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End

Option Explicit

Private mSelectedSheet As String

Private Sub UserForm_Initialize()
    Me.Caption = "データシートの選択"
    lstSheets.MultiSelect = False
End Sub

Public Sub Initialize(sheets As Variant)
    Dim i As Long
    lstSheets.Clear

    For i = LBound(sheets) To UBound(sheets)
        lstSheets.AddItem sheets(i)
    Next i

    If lstSheets.ListCount > 0 Then
        lstSheets.Selected(0) = True
    End If
End Sub

Private Sub cmdOK_Click()
    If lstSheets.ListIndex = -1 Then
        MsgBox "シートを選択してください。", vbExclamation
        Exit Sub
    End If

    mSelectedSheet = lstSheets.Value
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mSelectedSheet = ""
    Me.Hide
End Sub

Public Property Get SelectedSheet() As String
    SelectedSheet = mSelectedSheet
End Property

'*******************************************************************************
' TransferSortedData
' 概要：ソートされたデータをChartSheetに移行する
' 種別：メインプロシージャ
' 対象：入力=LOG_Helmet, 出力=ChartSheet
' 戻値：Boolean - 処理成功時はTrue、失敗時はFalse
'*******************************************************************************
Public Function TransferSortedData() As Boolean
    ' シート選択
    Dim sourceSheetName As String
    sourceSheetName = ShowSheetSelector()

    If sourceSheetName = "" Then
        MsgBox "シートが選択されていません。", vbExclamation
        TransferSortedData = False
        Exit Function
    End If

    ' 以下、既存のコードを修正
    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long

    ' 初期化
    TransferSortedData = False

    ' シートの存在確認と取得
    If Not CheckAndGetSheets(wsSource, wsTarget, sourceSheetName) Then
        MsgBox "必要なシートの確認に失敗しました。", vbExclamation
        Exit Function
    End If

    ' データ件数の確認
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "データが存在しません。", vbExclamation
        Exit Function
    End If

    ' レコード数チェック
    If (lastRow - 1) > MAX_RECORDS Then
        MsgBox "データが" & MAX_RECORDS & "件を超えています。" & vbNewLine & _
               "さらにソートして件数を絞ってください。" & vbNewLine & _
               "現在のレコード数: " & (lastRow - 1) & "件", vbExclamation
        Exit Function
    End If

    ' ターゲットシートのクリア
    ClearTargetSheet wsTarget

    ' データの移行
    If Not CopyDataToTarget(wsSource, wsTarget, lastRow) Then
        Exit Function
    End If

    TransferSortedData = True
    Exit Function

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbNewLine & _
           "エラー番号: " & Err.Number & vbNewLine & _
           "エラーの説明: " & Err.Description, vbCritical
    TransferSortedData = False
End Function

'*******************************************************************************
' CheckAndGetSheets
' 概要：必要なシートの存在確認と取得
' 種別：サブプロシージャ
' 引数：wsSource - 元シート, wsTarget - 移行先シート
' 戻値：Boolean - 成功時True
'*******************************************************************************
Private Function CheckAndGetSheets(ByRef wsSource As Worksheet, _
                                 ByRef wsTarget As Worksheet, _
                                 ByVal sourceSheetName As String) As Boolean
    On Error GoTo ErrorHandler

    CheckAndGetSheets = False

    ' ソースシートの確認
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    If wsSource Is Nothing Then
        MsgBox sourceSheetName & "シートが見つかりません。", vbExclamation
        Exit Function
    End If

    ' ターゲットシートの確認
    Set wsTarget = ThisWorkbook.Sheets("ChartSheet")
    If wsTarget Is Nothing Then
        MsgBox "ChartSheetが見つかりません。", vbExclamation
        Exit Function
    End If

    CheckAndGetSheets = True
    Exit Function

ErrorHandler:
    MsgBox "シートの確認中にエラーが発生しました。" & vbNewLine & _
           "エラー番号: " & Err.Number & vbNewLine & _
           "エラーの説明: " & Err.Description, vbCritical
    CheckAndGetSheets = False
End Function

'*******************************************************************************
' ClearTargetSheet
' 概要：移行先シートのクリア
' 種別：サブプロシージャ
' 引数：wsTarget - 移行先シート
'*******************************************************************************
Private Sub ClearTargetSheet(ByRef wsTarget As Worksheet)
    On Error Resume Next
    With wsTarget
        .Cells.Clear
        .Cells.Interior.ColorIndex = xlNone
    End With
End Sub

'*******************************************************************************
' CopyDataToTarget
' 概要：データを移行先シートにコピー
' 種別：サブプロシージャ
' 引数：wsSource - 元シート, wsTarget - 移行先シート, lastRow - 最終行
' 戻値：Boolean - 成功時True
'*******************************************************************************
Private Function CopyDataToTarget(ByRef wsSource As Worksheet, _
                                ByRef wsTarget As Worksheet, _
                                ByVal lastRow As Long) As Boolean
    On Error GoTo ErrorHandler

    CopyDataToTarget = False

    ' ヘッダー行のコピー
    wsSource.Rows(1).Copy Destination:=wsTarget.Rows(1)

    ' データ行のコピー
    wsSource.Range("A2:ZZ" & lastRow).Copy _
        Destination:=wsTarget.Range("A2")

    ' 書式の調整
    With wsTarget
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
    End With

    CopyDataToTarget = True
    Exit Function

ErrorHandler:
    MsgBox "データのコピー中にエラーが発生しました。" & vbNewLine & _
           "エラー番号: " & Err.Number & vbNewLine & _
           "エラーの説明: " & Err.Description, vbCritical
    CopyDataToTarget = False
End Function
