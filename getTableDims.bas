Attribute VB_Name = "Module2"
Sub getTableDimensions()

    '// get column and row count from specified sheet
    Dim dsht As Worksheet
    Set dsht = Sheets(1)
    numRows = dsht.UsedRange.Rows.Count
    numCols = dsht.UsedRange.Columns.Count

End Sub
