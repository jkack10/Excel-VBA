Attribute VB_Name = "Module2"
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
