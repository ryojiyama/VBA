Attribute VB_Name = "Module1"
Sub ExampleChartNameAndTitle()
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects(1)
    
    ' グラフの名前を設定
    chartObj.Name = "MyCustomChartName"
    
    ' グラフのタイトルを設定
    If Not chartObj.chart.HasTitle Then
        chartObj.chart.SetElement msoElementChartTitleAboveChart
    End If
    chartObj.chart.chartTitle.text = "My Chart Title"
End Sub

Sub TransferValues_old()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    
    ' シートをセット
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' ヘッダーの列番号を取得
    colHinban = 0
    colBoutai = 0
    colTencho = 0
    
    For Each cell In wsHelSpec.Rows(1).Cells
        If cell.value = "品番(D)" Then
            colHinban = cell.column
        ElseIf cell.value = "天頂肉厚" Then
            colTencho = cell.column
        End If
        If colHinban > 0 And colTencho > 0 Then Exit For
    Next cell
    
    For Each cell In wsSetting.Rows(1).Cells
        If cell.value = "帽体No." Then
            colBoutai = cell.column
            Exit For
        End If
    Next cell
    
    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "品番(D)" 列の値を探索し、転記
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
End Sub

Sub CopyAndSubtractValues_old()
    Dim wsHelSpec As Worksheet
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim i As Long
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim cell As Range
    
    ' シートをセット
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' ヘッダーの列番号を取得
    colTenchoSukima = 0
    colSokuteiSukima = 0
    colTenchoNikui = 0
    
    For Each cell In wsHelSpec.Rows(1).Cells
        If cell.value = "天頂すきま(N)" Then
            colTenchoSukima = cell.column
        ElseIf cell.value = "測定すきま" Then
            colSokuteiSukima = cell.column
        ElseIf cell.value = "天頂肉厚" Then
            colTenchoNikui = cell.column
        End If
    Next cell
    
    ' 必要な列が見つかったかを確認
    If colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "必要な列が見つかりません。ヘッダーを確認してください。", vbCritical
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colTenchoSukima).End(xlUp).row
    
    ' "天頂すきま(N)" の値を "測定すきま" にコピーし、値を計算
    For i = 2 To lastRowHelSpec
        ' 各セルの値を取得
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "天頂すきま(N)"の値を"測定すきま"にコピー
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "天頂すきま(N)"の値から"天頂肉厚"の値を引く
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q列とR列に"合格"の値を代入
        wsHelSpec.Cells(i, 17).value = "合格" ' Q列は17番目の列
        wsHelSpec.Cells(i, 18).value = "合格" ' R列は18番目の列
    Next i
End Sub



Sub TransferValues()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    
    ' シートをセット
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' ヘッダーの列番号を取得
    colHinban = GetColumnIndex(wsHelSpec, "品番(D)")
    colTencho = GetColumnIndex(wsHelSpec, "天頂肉厚")
    colBoutai = GetColumnIndex(wsSetting, "帽体No.")
    
    ' 必要な列が見つかったかを確認
    If colHinban = 0 Or colTencho = 0 Or colBoutai = 0 Then
        MsgBox "必要な列が見つかりません。ヘッダーを確認してください。", vbCritical
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "品番(D)" 列の値を探索し、転記
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
End Sub

Sub CopyAndSubtractValues()
    Dim wsHelSpec As Worksheet
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim i As Long
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim cell As Range
    
    ' シートをセット
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    
    ' ヘッダーの列番号を取得
    colTenchoSukima = GetColumnIndex(wsHelSpec, "天頂すきま(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "測定すきま")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "天頂肉厚")
    
    ' 必要な列が見つかったかを確認
    If colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "必要な列が見つかりません。ヘッダーを確認してください。", vbCritical
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colTenchoSukima).End(xlUp).row
    
    ' "天頂すきま(N)" の値を "測定すきま" にコピーし、値を計算
    For i = 2 To lastRowHelSpec
        ' 各セルの値を取得
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "天頂すきま(N)"の値を"測定すきま"にコピー
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "天頂すきま(N)"の値から"天頂肉厚"の値を引く
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q列とR列に"合格"の値を代入
        wsHelSpec.Cells(i, 17).value = "合格" ' Q列は17番目の列
        wsHelSpec.Cells(i, 18).value = "合格" ' R列は18番目の列
    Next i
End Sub




Sub UpdateCrownClearance()
    Dim wsHelSpec As Worksheet
    Dim wsSetting As Worksheet
    Dim colHinban As Integer
    Dim colBoutai As Integer
    Dim colTencho As Integer
    Dim colTenchoSukima As Integer
    Dim colSokuteiSukima As Integer
    Dim colTenchoNikui As Integer
    Dim lastRowHelSpec As Long
    Dim lastRowSetting As Long
    Dim cell As Range
    Dim tenSukima As Long
    Dim valueToFind As Variant
    Dim tenchoSukimaValue As Variant
    Dim tenchoNikuiValue As Variant
    Dim i As Long
    
    ' シートをセット
    Set wsHelSpec = ThisWorkbook.Sheets("Hel_SpecSheet")
    Set wsSetting = ThisWorkbook.Sheets("Setting")
    
    ' ヘッダーの列番号を取得
    colHinban = GetColumnIndex(wsHelSpec, "品番(D)")
    colBoutai = GetColumnIndex(wsSetting, "帽体No.")
    colTencho = GetColumnIndex(wsHelSpec, "天頂肉厚")
    colTenchoSukima = GetColumnIndex(wsHelSpec, "天頂すきま(N)")
    colSokuteiSukima = GetColumnIndex(wsHelSpec, "測定すきま")
    colTenchoNikui = GetColumnIndex(wsHelSpec, "天頂肉厚")
    
    ' 必要な列が見つかったかを確認
    If colHinban = 0 Or colBoutai = 0 Or colTencho = 0 Or colTenchoSukima = 0 Or colSokuteiSukima = 0 Or colTenchoNikui = 0 Then
        MsgBox "必要な列が見つかりません。ヘッダーを確認してください。", vbCritical
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRowHelSpec = wsHelSpec.Cells(wsHelSpec.Rows.Count, colHinban).End(xlUp).row
    lastRowSetting = wsSetting.Cells(wsSetting.Rows.Count, colBoutai).End(xlUp).row
    
    ' "品番(D)" 列の値を探索し、転記
    For Each cell In wsHelSpec.Range(wsHelSpec.Cells(2, colHinban), wsHelSpec.Cells(lastRowHelSpec, colHinban))
        valueToFind = cell.value
        For tenSukima = 2 To lastRowSetting
            If wsSetting.Cells(tenSukima, colBoutai).value = valueToFind Then
                wsHelSpec.Cells(cell.row, colTencho).value = wsSetting.Cells(tenSukima, "H").value
                Exit For
            End If
        Next tenSukima
    Next cell
    
    ' "天頂すきま(N)" の値を "測定すきま" にコピーし、値を計算
    For i = 2 To lastRowHelSpec
        ' 各セルの値を取得
        tenchoSukimaValue = wsHelSpec.Cells(i, colTenchoSukima).value
        tenchoNikuiValue = wsHelSpec.Cells(i, colTenchoNikui).value
        
        ' "天頂すきま(N)"の値を"測定すきま"にコピー
        If IsNumeric(tenchoSukimaValue) Then
            wsHelSpec.Cells(i, colSokuteiSukima).value = tenchoSukimaValue
        End If
        
        ' "天頂すきま(N)"の値から"天頂肉厚"の値を引く
        If IsNumeric(tenchoSukimaValue) And IsNumeric(tenchoNikuiValue) Then
            wsHelSpec.Cells(i, colTenchoSukima).value = tenchoSukimaValue - tenchoNikuiValue
        End If
        
        ' Q列とR列に"合格"の値を代入
        wsHelSpec.Cells(i, 17).value = "合格" ' Q列は17番目の列
        wsHelSpec.Cells(i, 18).value = "合格" ' R列は18番目の列
    Next i
End Sub

Function GetColumnIndex(sheet As Worksheet, headerName As String) As Integer
    Dim cell As Range
    For Each cell In sheet.Rows(1).Cells
        If cell.value = headerName Then
            GetColumnIndex = cell.column
            Exit Function
        End If
    Next cell
    GetColumnIndex = 0 ' 見つからない場合は0を返す
End Function


