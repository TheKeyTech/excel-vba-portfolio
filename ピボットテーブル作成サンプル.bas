Attribute VB_Name = "Module1"
Option Explicit

Sub 実績取り込み()
    Dim filePath As String
    Dim saveBookPath As String

    ' ▼ダウンロード先でも動く相対パスに変更（このブックの隣に data\sales_data.csv を置く）
    filePath = ThisWorkbook.Path & "\data\sales_data.csv"
    saveBookPath = ThisWorkbook.Path & "\sales_data.xlsx"

    ' 画面更新と再計算を停止
    ToggleApplicationSettings False

    ' ▼CSVを開く
    Workbooks.Open filePath

    ' ▼xlsxとして保存（保存名を明示）→ 以後は "sales_data.xlsx" を前提に処理
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs _
        Filename:=saveBookPath, _
        FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    ' ▼データシート名を固定（後工程がシート名で参照するため）
    On Error Resume Next
    Workbooks("sales_data.xlsx").Worksheets(1).Name = "sales_data"
    On Error GoTo 0

    ' ▼「集計結果」シートが無ければ作成（あってもそのまま流用）
    Dim ws As Worksheet, hasTarget As Boolean
    hasTarget = False
    For Each ws In Workbooks("sales_data.xlsx").Worksheets
        If ws.Name = "集計結果" Then
            hasTarget = True
            Exit For
        End If
    Next ws
    If Not hasTarget Then
        Workbooks("sales_data.xlsx").Worksheets.Add(Before:=Workbooks("sales_data.xlsx").Worksheets(1)).Name = "集計結果"
    End If

    ' ▼ピボット作成
    Call CreatePivotTable

    ' 画面更新と再計算を再開
    ToggleApplicationSettings True
End Sub

Sub CreatePivotTable()
    ' ピボットテーブル作成
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim wb As Workbook
    Dim frWS As Worksheet
    Dim toWS As Worksheet

    Set wb = Workbooks("sales_data.xlsx")
    Set frWS = wb.Worksheets("sales_data")
    Set toWS = wb.Worksheets("集計結果")

    ' 既存ピボットがあれば領域をクリア（シート削除はしない）
    Dim ptOld As PivotTable
    On Error Resume Next
    Set ptOld = toWS.PivotTables("実績集計結果")
    On Error GoTo 0
    If Not ptOld Is Nothing Then
        ptOld.TableRange2.Clear
    Else
        toWS.Cells.Clear
    End If

    Set pc = ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=frWS.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
            TableDestination:=toWS.Range("A3"), _
            TableName:="実績集計結果")

    With pt
        .PivotFields("Product").Orientation = xlRowField
        .PivotFields("Product").Position = 1

        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 2

        .AddDataField .PivotFields("Sales"), "合計重量", xlSum
        .PivotFields("Sales").NumberFormat = "#,##0"
    End With

    pt.RowAxisLayout xlOutlineRow
    toWS.Range("A3").CurrentRegion.Columns.AutoFit

    ' ブックを保存して閉じる
    wb.Close SaveChanges:=True
End Sub

Sub ToggleApplicationSettings(enable As Boolean)
    ' enableがTrueの場合、画面更新と再計算を再開
    ' enableがFalseの場合、画面更新と再計算を停止
    Application.DisplayAlerts = enable
    Application.ScreenUpdating = enable
    If enable Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub


