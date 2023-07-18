' check if sheet already exists
Function SheetExists(ByVal sheetname As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ActiveWorkbook.Sheets(sheetname)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function


Sub excel_metadata()
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    Dim sheetname As String
    sheetname = "excel.metadata"
    If SheetExists(sheetname) = False Then
        wb.Sheets.Add(After:=wb.Worksheets(1)).Name = sheetname
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
    
    ws_meta.Select
    
    With ws_meta
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        .Range(Cells(1, 1), Cells(1, i)).Columns.AutoFit
        Cells(1, 1).Select
    End With
End Sub


Sub combine_multiple_sheets()

    Application.ScreenUpdating = False
    
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    
    'create combined sheet
    Dim sheetname As String
    sheetname = "All"
    If SheetExists(sheetname) = False Then
        wb.Sheets.Add(After:=wb.Worksheets(1)).Name = sheetname
    End If
    ' destination sheet
    Dim ws_dst As Worksheet
    Set ws_dst = Sheets(sheetname)
    ws_dst.Select
    ws_dst.Cells.Clear
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        wb.Sheets(ws.Name).Select
        ' clear the filter and show all data
        With ws
            If .AutoFilterMode Then .AutoFilter.ShowAllData
            .Columns.Hidden = False
            .Rows.Hidden = False
            .Cells(1, 1).Select
        End With
        
        If ws.Name <> sheetname And ws.Name <> "excel.metadata" Then
        
            Dim last_row, last_col, dst_last_row As Long
            'last_row = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
            'last_row = Cells(Rows.Count, 1).End(xlUp).Row
            last_row = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
                    
            'last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
            
            last_col = Cells.Find(What:="*", _
                                After:=Range("A1"), _
                                LookAt:=xlPart, _
                                LookIn:=xlFormulas, _
                                SearchOrder:=xlByColumns, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False).Column
            
            'header row
            For i = 1 To last_col
                ws_dst.Cells(1, 1).Value = "Source Sheet Name"
                If (Not IsEmpty(ws.Cells(1, i).Value)) And Len(Trim(ws.Cells(1, i).Value)) > 0 Then
                    ws.Cells(1, i + 1).Copy
                    ws_dst.Cells(1, i + 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
                    ws_dst.Cells(1, i + 1).PasteSpecial Paste:=xlPasteColumnWidths
                End If
            Next

            Dim rng_src, rng_dst As Range
            
            Set rng_scr = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col))
            rng_scr.Copy
            wb.Sheets(ws_dst.Name).Select
            dst_last_row = ws_dst.UsedRange.Rows(ws_dst.UsedRange.Rows.Count).Row
            
            ws_dst.Cells(dst_last_row + 1, 2).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            ws_dst.Range(ws_dst.Cells(dst_last_row + 1, 1), ws_dst.Cells(dst_last_row + last_row - 1, 1)).Value = ws.Name
            

        End If
        
    Next ws
    dst_last_row = ws_dst.UsedRange.Rows(ws_dst.UsedRange.Rows.Count).Row
    ws_dst.Select
    ws_dst.Range("A1").AutoFilter
    ws_dst.Range("A1:A" & dst_last_row).Borders.LineStyle = xlContinuous
    ws_dst.Columns("A").EntireColumn.ColumnWidth = 15
    ws_dst.Columns("E").EntireColumn.ColumnWidth = 15
    ws_dst.Columns("M:V").EntireColumn.ColumnWidth = 15
    
    ws_dst.Rows("2:" & dst_last_row).RowHeight = 50
 
    MsgBox "sheet combined to Sheet [All]"
    

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub


' Check the sheets and colums

' Combine sheets