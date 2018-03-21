Attribute VB_Name = "Module2"
Sub createTemplatesV2()

  Dim wb As Workbook
  Dim ThisSheet As Worksheet
  Dim NumOfColumns As Integer
  Dim RangeToCopy As Range
  Dim RangeOfHeader As Range
  Dim WorkbookCounter As Integer
  Dim RowsInFile

  Application.ScreenUpdating = False

  'Initialize data
  Set ThisSheet = ThisWorkbook.ActiveSheet
  NumOfColumns = ThisSheet.UsedRange.Columns.Count
  WorkbookCounter = 1
  RowsInFile = 101

  'Copy the data of the first row (header)
  Set RangeOfHeader = ThisSheet.Range(ThisSheet.Cells(1, 1), ThisSheet.Cells(1, NumOfColumns))

  For p = 2 To ThisSheet.UsedRange.Rows.Count Step RowsInFile - 1
    Set wb = Workbooks.Add

    'Paste the header row in new file
    RangeOfHeader.Copy wb.Sheets(1).Range("A1")

    'Paste the chunk of rows for this file
    Set RangeToCopy = ThisSheet.Range(ThisSheet.Cells(p, 1), ThisSheet.Cells(p + RowsInFile - 2, NumOfColumns))
    RangeToCopy.Copy wb.Sheets(1).Range("A2")

    'Save the new workbook, and close it
    wb.SaveAs ThisWorkbook.Path & "\LTOverride_Template_" & WorkbookCounter
    wb.Close

    'Increment file counter
    WorkbookCounter = WorkbookCounter + 1
  Next p

  Application.ScreenUpdating = True
  Set wb = Nothing
End Sub
