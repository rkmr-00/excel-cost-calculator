Attribute VB_Name = "CostCalculator"
' ============================================================
' 原価計算マクロ
' 入力シート: A列=品目, B列=数量, C列=単価（2行目から）
' 実行後: 「集計」シートに小計・合計・グラフを自動生成
' ============================================================

Sub CalculateCost()
    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalCost As Double

    Application.ScreenUpdating = False

    Set wsInput = ThisWorkbook.Sheets("入力")

    ' 集計シートを初期化（なければ作成）
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("集計")
    On Error GoTo 0
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=wsInput)
        wsReport.Name = "集計"
    Else
        wsReport.Cells.Clear
    End If

    ' 入力データの最終行を取得
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' ヘッダー
    With wsReport
        .Cells(1, 1).Value = "品目"
        .Cells(1, 2).Value = "数量"
        .Cells(1, 3).Value = "単価（円）"
        .Cells(1, 4).Value = "小計（円）"
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:D1").Font.Color = RGB(255, 255, 255)
    End With

    ' データ転記 & 小計計算
    totalCost = 0
    Dim reportRow As Long
    reportRow = 2
    For i = 2 To lastRow
        If wsInput.Cells(i, 1).Value = "" Then Exit For
        Dim qty As Double
        Dim price As Double
        Dim subtotal As Double
        qty = wsInput.Cells(i, 2).Value
        price = wsInput.Cells(i, 3).Value
        subtotal = qty * price

        wsReport.Cells(reportRow, 1).Value = wsInput.Cells(i, 1).Value
        wsReport.Cells(reportRow, 2).Value = qty
        wsReport.Cells(reportRow, 3).Value = price
        wsReport.Cells(reportRow, 4).Value = subtotal

        ' 数値書式
        wsReport.Cells(reportRow, 3).NumberFormat = "#,##0"
        wsReport.Cells(reportRow, 4).NumberFormat = "#,##0"

        totalCost = totalCost + subtotal
        reportRow = reportRow + 1
    Next i

    ' 合計行
    wsReport.Cells(reportRow, 3).Value = "合計"
    wsReport.Cells(reportRow, 4).Value = totalCost
    wsReport.Cells(reportRow, 3).Font.Bold = True
    wsReport.Cells(reportRow, 4).Font.Bold = True
    wsReport.Cells(reportRow, 4).NumberFormat = "#,##0"
    wsReport.Cells(reportRow, 4).Interior.Color = RGB(255, 255, 200)

    ' 列幅自動調整
    wsReport.Columns("A:D").AutoFit

    ' 棒グラフ作成（品目別小計）
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Set chartRange = wsReport.Range("A1:A" & (reportRow - 1))
    Set chartRange = Union(chartRange, wsReport.Range("D1:D" & (reportRow - 1)))

    Set chartObj = wsReport.ChartObjects.Add(Left:=320, Top:=20, Width:=400, Height:=260)
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "品目別コスト"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "品目"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "金額（円）"
    End With

    Application.ScreenUpdating = True
    MsgBox "集計完了！合計コスト: ¥" & Format(totalCost, "#,##0"), vbInformation, "原価計算完了"
End Sub


Sub SetupInputSheet()
    ' 入力シートのひな形を作成するセットアップマクロ
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("入力")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "入力"
    End If

    With ws
        .Cells(1, 1).Value = "品目"
        .Cells(1, 2).Value = "数量"
        .Cells(1, 3).Value = "単価（円）"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(200, 230, 200)

        ' サンプルデータ
        .Cells(2, 1).Value = "材料費A"
        .Cells(2, 2).Value = 10
        .Cells(2, 3).Value = 5000

        .Cells(3, 1).Value = "外注費B"
        .Cells(3, 2).Value = 2
        .Cells(3, 3).Value = 30000

        .Cells(4, 1).Value = "経費C"
        .Cells(4, 2).Value = 1
        .Cells(4, 3).Value = 8000

        .Columns("A:C").AutoFit
    End With

    MsgBox "入力シートを作成しました。データを入力後、「CalculateCost」を実行してください。", vbInformation
End Sub
