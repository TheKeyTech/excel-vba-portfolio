Attribute VB_Name = "Module1"
Option Explicit

Sub ���ю�荞��()
    Dim filePath As String
    Dim saveBookPath As String

    ' ���_�E�����[�h��ł��������΃p�X�ɕύX�i���̃u�b�N�ׂ̗� data\sales_data.csv ��u���j
    filePath = ThisWorkbook.Path & "\data\sales_data.csv"
    saveBookPath = ThisWorkbook.Path & "\sales_data.xlsx"

    ' ��ʍX�V�ƍČv�Z���~
    ToggleApplicationSettings False

    ' ��CSV���J��
    Workbooks.Open filePath

    ' ��xlsx�Ƃ��ĕۑ��i�ۑ����𖾎��j�� �Ȍ�� "sales_data.xlsx" ��O��ɏ���
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs _
        Filename:=saveBookPath, _
        FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    ' ���f�[�^�V�[�g�����Œ�i��H�����V�[�g���ŎQ�Ƃ��邽�߁j
    On Error Resume Next
    Workbooks("sales_data.xlsx").Worksheets(1).Name = "sales_data"
    On Error GoTo 0

    ' ���u�W�v���ʁv�V�[�g��������΍쐬�i�����Ă����̂܂ܗ��p�j
    Dim ws As Worksheet, hasTarget As Boolean
    hasTarget = False
    For Each ws In Workbooks("sales_data.xlsx").Worksheets
        If ws.Name = "�W�v����" Then
            hasTarget = True
            Exit For
        End If
    Next ws
    If Not hasTarget Then
        Workbooks("sales_data.xlsx").Worksheets.Add(Before:=Workbooks("sales_data.xlsx").Worksheets(1)).Name = "�W�v����"
    End If

    ' ���s�{�b�g�쐬
    Call CreatePivotTable

    ' ��ʍX�V�ƍČv�Z���ĊJ
    ToggleApplicationSettings True
End Sub

Sub CreatePivotTable()
    ' �s�{�b�g�e�[�u���쐬
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim wb As Workbook
    Dim frWS As Worksheet
    Dim toWS As Worksheet

    Set wb = Workbooks("sales_data.xlsx")
    Set frWS = wb.Worksheets("sales_data")
    Set toWS = wb.Worksheets("�W�v����")

    ' �����s�{�b�g������Η̈���N���A�i�V�[�g�폜�͂��Ȃ��j
    Dim ptOld As PivotTable
    On Error Resume Next
    Set ptOld = toWS.PivotTables("���яW�v����")
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
            TableName:="���яW�v����")

    With pt
        .PivotFields("Product").Orientation = xlRowField
        .PivotFields("Product").Position = 1

        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 2

        .AddDataField .PivotFields("Sales"), "���v�d��", xlSum
        .PivotFields("Sales").NumberFormat = "#,##0"
    End With

    pt.RowAxisLayout xlOutlineRow
    toWS.Range("A3").CurrentRegion.Columns.AutoFit

    ' �u�b�N��ۑ����ĕ���
    wb.Close SaveChanges:=True
End Sub

Sub ToggleApplicationSettings(enable As Boolean)
    ' enable��True�̏ꍇ�A��ʍX�V�ƍČv�Z���ĊJ
    ' enable��False�̏ꍇ�A��ʍX�V�ƍČv�Z���~
    Application.DisplayAlerts = enable
    Application.ScreenUpdating = enable
    If enable Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub


