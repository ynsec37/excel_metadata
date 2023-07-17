Sub excel_metadata()

    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    Dim sheetname As String
   
    sheetname = "excel.metadata"
    If SheetExists(sheetname) = False Then
        wb.Sheets.Add(After:=wb.Worksheets(Worksheets.Count)).Name = sheetname
    End If
    
    Dim ws_meta As Worksheet
    Set ws_meta = Sheets(sheetname)
    
    ws_meta.Cells.Clear
    
    Dim ws As Worksheet
    Dim i, j, last_col As Long
    i = 0
    For Each ws In wb.Worksheets
        If ws.Name <> sheetname And ws.Name <> "" Then
            i = i + 1
            ws_meta.Cells(1, i).Value = ws.Name

            last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

            For j = 1 To last_col
                ws_meta.Cells(1 + j, i).Value = ws.Cells(1, j).Value
            Next
        End If
    Next
    With ws_meta
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        .Range(Cells(1, 1), Cells(1, i)).Columns.AutoFit
        Cells(1, 1).Select
    End With
End Sub