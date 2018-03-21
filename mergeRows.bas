
Sub Macro2()

' ********************************************************************
' ** Highlight cell range (only ONE(1) column wide will work)       **
' ** and then activate macro.                                       **
' **                                                                **
' ** Will result in merging the contents of all highlighted         **
' ** cells, separated by carraige returns                           **
' **                                                                **
' ** You must highligh from the top down.  Highlighting from        **
' ** bottom up will cause the output to be incorrect and located    **
' ** in the bottom cell instead of the top                          **
' ********************************************************************

str1 = ActiveCell.Text
 
'LOOP A
For cRow = 1 To Selection.Rows.Count - 1

    'build string of all rows
    str1 = str1 & Chr(10) & ActiveCell.Offset(cRow, 0).Text

    'delete un-needed text
    ActiveCell.Offset(cRow, 0).Value = ""

Next cRow
'END LOOP A

'Output
ActiveCell.Value = str1
ActiveCell.Select

End Sub
