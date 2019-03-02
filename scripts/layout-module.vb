'Inserts a given number of columns before the active cell.
Sub InsertMultipleColumns()
    Dim i As Integer
    Dim j As Integer
    ActiveCell.EntireColumn.Select
    On Error GoTo Last
    i = InputBox("Enter number of columns to insert", "Insert Columns")
    For j = 1 To i
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove
    Next j
    Last:Exit Sub
End Sub

'Inserts a given number of rows before the active cell.
Sub InsertMultipleRows()
    Dim i As Integer
    Dim j As Integer
    ActiveCell.EntireRow.Select
    On Error GoTo Last
    i = InputBox("Enter number of columns to insert", "Insert Columns")
    For j = 1 To i
        Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromRightorAbove
    Next j
    Last:Exit Sub
End Sub

'Fits all columns.
Sub AutoFitColumns()
    Set cell = ActiveCell
    Cells.Select
    Cells.EntireColumn.AutoFit
    cell.Select
End Sub

'Fits all rows.
Sub AutoFitRows()
    Set cell = ActiveCell
    Cells.Select
    Cells.EntireRow.AutoFit
    cell.Select
End Sub

'Makes the selected cells equal size in the original size of the selected range.
Sub EqualizeColumns()
    Set iColumns = Selection.Columns
    On Error GoTo Last
    sum_size = 0
    For Each col In iColumns
        sum_size = sum_size + col.ColumnWidth
    Next
    size = sum_size / iColumns.Count
    iColumns.ColumnWidth = size
    Last:Exit Sub
End Sub