Attribute VB_Name = "LOG_Decorate"

Sub SaveChartsAsPNG()
    ' グラフをPNGに変換しデスクトップのフォルダに保存する。
    ' ワークシートの名前を宣言
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    ' Windowsのデスクトップのパスを取得
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    ' 今日の日付を取得し、指定のフォルダ名を作成
    Dim folderName As String
    folderName = "Graph_" & Format(Date, "yyyy-mm-dd")

    ' フォルダのパスを作成
    Dim folderPath As String
    folderPath = desktopPath & "\" & folderName

    ' フォルダが存在しない場合、新たに作成
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    Dim ws As Worksheet
    Dim i As Integer
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
        
            ' チャートオブジェクトを宣言
            Dim ChartObj As ChartObject

            ' ファイル名を宣言
            Dim FileName As String

            ' チャートオブジェクトごとにループ
            For Each ChartObj In ws.ChartObjects

                ' グラフのタイトルを一時的に保存し、グラフからは削除
                FileName = ChartObj.chart.ChartTitle.Text
                ChartObj.chart.HasTitle = False

                ' ファイル名に ".png" を追加
                FileName = FileName & ".png"

                ' ファイルパスを設定（フォルダのパス + ファイル名）
                Dim filePath As String
                filePath = folderPath & "\" & FileName

                ' チャートの現在の幅と高さを保存
                Dim originalWidth As Double
                Dim originalHeight As Double
                originalWidth = ChartObj.Width
                originalHeight = ChartObj.Height

                ' チャートの幅を設定し、高さはアスペクト比を保持
                Dim aspectRatio As Double
                aspectRatio = ChartObj.Height / ChartObj.Width
                ChartObj.Width = 1000
                ChartObj.Height = 1000 * aspectRatio

                ' チャートをPNGファイルとしてエクスポート
                ChartObj.chart.Export FileName:=filePath, FilterName:="PNG"

                ' チャートの幅と高さを元の大きさに戻す
                ChartObj.Width = originalWidth
                ChartObj.Height = originalHeight

                ' グラフのタイトルを元に戻す
                ChartObj.chart.HasTitle = True
                ChartObj.chart.ChartTitle.Text = FileName
            Next ChartObj
        End If
        
        Set ws = Nothing
    Next i
End Sub

Sub ApplyColorToAllSheets()
    '各ログシートに色をつけたりする
    Dim sheetNames As Variant
    sheetNames = Array("LOG_Helmet", "LOG_BaseBall", "LOG_Bicycle", "LOG_FallArrest")
    
    Dim ws As Worksheet
    Dim i As Integer
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Call ColorCells(ws)
            Set ws = Nothing
        End If
    Next i
End Sub

Sub ColorCells(ws As Worksheet)
    'ApplyColorToALlSHeetsの関数
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim cellColor As Long

    ' A列の最終行を取得します
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' A列の2行目から最終行までの範囲を定義します
    Set rng = ws.Range("A2:A" & lastRow)

    ' 範囲内の各セルについてループします
    For Each cell In rng
        If InStr(cell.Value, "HEL") > 0 Then
            ' "HEL"がセルの値の一部である場合、G列とH列のセルの色をオレンジにします
            cellColor = RGB(255, 111, 56)
            ColorAndFont ws.Range("H" & cell.row & ":I" & cell.row), cellColor
        ElseIf InStr(cell.Value, "BICYCLE") > 0 Then
            ' "BICYCLE"がセルの値の一部である場合、I列のセルの色をブルーにします
            cellColor = RGB(8, 92, 255)
            ColorAndFont ws.Range("I" & cell.row), cellColor
        ElseIf InStr(cell.Value, "BASEBALL") > 0 Then
            ' "BASEBALL"がセルの値の一部である場合、K列のセルの色をグレーにします
            cellColor = RGB(218, 218, 218)
            ColorAndFont ws.Range("K" & cell.row), cellColor
        ElseIf InStr(cell.Value, "FALLARR") > 0 Then
            ' "FALLARR"がセルの値の一部である場合、L列からN列のセルの色を緑にします
            cellColor = RGB(22, 187, 98)
            ColorAndFont ws.Range("L" & cell.row & ":N" & cell.row), cellColor
        End If

        ' F列のセルの色も同様に変更します
        ColorAndFont ws.Range("F" & cell.row), cellColor
    Next cell
End Sub

Sub ColorAndFont(rng As Range, color As Long)
    'ColorCellsの関数
    rng.Interior.color = color
    rng.Font.color = RGB(255, 255, 255)
    rng.Font.Bold = True
End Sub

Sub DataMidrationAndCSVSheetDelete()
    Call FillColumnsQtoS
    Call CustomSort_Helmet
End Sub
Sub FillColumnsQtoS()
    ' 検査表の項目に便宜上の合格印を押す。
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' "LOG_Helmet"シートを指定
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row

    ' 上から最終行までループ
    For i = 2 To lastRow
        ' S,T列に "合格" を入力
        ws.Cells(i, "S").Value = "合格"
        ws.Cells(i, "T").Value = "合格"

        ' Q列に "更新" を入力
        'ws.Cells(i, "S").Value = "更新"
    Next i

    ' メモリの開放
    Set ws = Nothing

End Sub

Sub CustomSort_Helmet()
    'B列を天頂、前頭部、後頭部、その他の順にソートする。
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' データの範囲を指定します。1行目は無視するので2から始まります。
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim rng As Range
    Set rng = ws.Range("B2:B" & lastRow)
    
    ' 新しい列を追加して、ソートキーを設定します。
    ws.Columns("C").Insert
    Dim cell As Range
    For Each cell In rng
        If InStr(cell.Value, "HEL_TOP") > 0 Then
            cell.Offset(0, 1).Value = 10000 + CInt(Mid(cell.Value, 1, 4)) ' HEL_TOPの場合
        ElseIf InStr(cell.Value, "HEL_FRONT") > 0 Then
            cell.Offset(0, 1).Value = 20000 + CInt(Mid(cell.Value, 1, 4)) ' HEL_FRONTの場合
        ElseIf InStr(cell.Value, "HEL_BACK") > 0 Then
            cell.Offset(0, 1).Value = 30000 + CInt(Mid(cell.Value, 1, 4)) ' HEL_BACKの場合
        ElseIf InStr(cell.Value, "HEL_ZENGO") > 0 Then
            cell.Offset(0, 1).Value = 40000 + CInt(Mid(cell.Value, 1, 4)) ' HEL_ZENGOの場合
        End If
    Next cell
    
    ' 全ての列（A列から最後の列まで）でソートします。
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Sort Key1:=ws.Range("C2"), Order1:=xlAscending, Header:=xlNo
    
    ' ソートに使用した列を削除します。
    ws.Columns("C").Delete
End Sub



Sub GenerateSampleID()
    ' 試料用のIDを生成する。
    Dim ws As Worksheet
    Dim rng As Range
    Dim dic As Object
    Dim i As Long
    Dim key As String
    Dim prefix As String
    Dim idNum As Long
    Dim randChars As String

    ' "LOG_Helmet"ワークシートを指定する
    Set ws = ThisWorkbook.Sheets("LOG_Helmet")

    ' Scripting.Dictionaryを作成する
    Set dic = CreateObject("Scripting.Dictionary")

    ' データ範囲を指定する
    Set rng = ws.Range("D2:P" & ws.Cells(ws.Rows.Count, "D").End(xlUp).row)

    ' 接頭辞を設定する
    prefix = "_Hel"

    For i = 1 To rng.Rows.Count
        ' D列、M列、N列、O列、L列(前処理)の値を結合してキーを作成する
        key = ws.Cells(i + 1, "D").Value & "_" & ws.Cells(i + 1, "M").Value & "_" & ws.Cells(i + 1, "N").Value & "_" & ws.Cells(i + 1, "O").Value & "_" & ws.Cells(i + 1, "L").Value

        ' キーが既にdicに存在する場合、既存のIDを使用する。存在しない場合、新たなIDを生成する
        If dic.Exists(key) Then
            ws.Cells(i + 1, "C").Value = dic(key)
        Else
            idNum = idNum + 1
            ' ランダムなアルファベット2文字を生成する
            randChars = Chr(Int((90 - 65 + 1) * Rnd + 65)) & Chr(Int((90 - 65 + 1) * Rnd + 65))
            ' ランダムな文字を追加してIDを生成する
            dic.Add key, Format(idNum, "00000") & randChars & prefix & ws.Cells(i + 1, "D").Value
            ws.Cells(i + 1, "C").Value = dic(key)
        End If
    Next i
End Sub





