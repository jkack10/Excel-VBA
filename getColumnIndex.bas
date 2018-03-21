Attribute VB_Name = "getColumnIndex"

'//function - pass sheet (as sheet) and column name (as string) and get back the column number, A=1, B=2, etc.

Function getColumn(dsht As Worksheet, colName As String)

    numCols = dsht.UsedRange.Columns.Count
    For x = 1 To numCols
        If Cells(1, x).Value = colName Then
            getColumn = x
            Exit For
        Else
            getColumn = "Error: Column with that value not found."
        End If
    Next x

End Function
