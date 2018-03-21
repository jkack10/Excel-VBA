Attribute VB_Name = "Module1"
Sub WBSNumbering()

'From http://j.modjeska.us/?p=31
'Renumber tasks on a project plan
'Associate this code with a button or other control on your spreadsheet

'Layout Assumptions:
'Row 1 contains column headings
'Column A contains WBS numbers
'Column B contains Task description, with appropriate indentation
'Some text (here we assume "END OF PROJECT") delimits the end of the task list

    On Error Resume Next

    'Hide page breaks and disable screen updating (speeds up processing)
    Application.ScreenUpdating = False
    ActiveSheet.DisplayPageBreaks = False
    'Format WBS column as text (so zeros are not truncated)
    ActiveSheet.Range("A:A").NumberFormat = "@"
    Dim r As Long                   'Row counter
    Dim depth As Long               'How many "decimal" places for each task
    Dim wbsarray() As Long          'Master array holds counters for each WBS level
    Dim basenum As Long             'Whole number sequencing variable
    Dim wbs As String               'The WBS string for each task
    Dim aloop As Long               'General purpose For/Next loop counter

    r = 3                           'Starting row
    basenum = 0                     'Initialize whole numbers
    ReDim wbsarray(0 To 0) As Long  'Initialize WBS ennumeration array

    'Loop through cells with project tasks and generate WBS
    Do While Cells(r, 2) <> ""

        'Ignore empty tasks in column B
        If Cells(r, 2) <> "" Then

           'Skip hidden rows
            If Rows(r).EntireRow.Hidden = False Then

                'Get indentation level of task in col B
                depth = Cells(r, 2).IndentLevel

                'Case if no depth (whole number master task)
                If depth = 0 Then

                    'increment WBS base number
                    basenum = basenum + 1
                    wbs = CStr(basenum)
                    ReDim wbsarray(0 To 0)

                'Case if task has WBS depth (is a subtask, sub-subtask, etc.)
                Else

                    'Resize the WBS array according to current depth
                    ReDim Preserve wbsarray(0 To depth) As Long

                    'Repurpose depth to refer to array size; arrays start at 0
                    depth = depth - 1

                    'Case if this is the first subtask
                    If wbsarray(depth) <> 0 Then

                        wbsarray(depth) = wbsarray(depth) + 1

                    'Case if we are incrementing a subtask
                    Else

                        wbsarray(depth) = 1

                    End If

                    'Only ennumerate WBS as deep as the indentation calls for;
                    'so we clear previous stored values for deeper levels
                    If wbsarray(depth + 1) <> 0 Then
                        For aloop = depth + 1 To UBound(wbsarray)
                            wbsarray(aloop) = 0
                        Next aloop
                    End If

                    'Assign contents of array to WBS string
                    wbs = CStr(basenum)

                    For aloop = 0 To depth
                        wbs = wbs & "." & CStr(wbsarray(aloop))
                    Next aloop

                End If

                'Populate target cell with WBS number
                Cells(r, 1).Value = wbs

                'Get rid of annoying "number stored as text" error
                Cells(r, 1).Errors(xlNumberAsText).Ignore = True

                'Apply text format: next row is deeper than current
                If Cells(r + 1, 2).IndentLevel > Cells(r, 2).IndentLevel Then

                    Cells(r, 1).Font.Bold = True
                    Cells(r, 2).Font.Bold = True
                'Else (next row is same/shallower than current) no format
                Else
                    Cells(r, 1).Font.Bold = False
                    Cells(r, 2).Font.Bold = False
                End If
                'Special formatting for master (whole number) tasks)
                If Cells(r, 2).IndentLevel = 0 Then
                    Cells(r, 1).Font.Bold = True
                    Cells(r, 2).Font.Bold = True
                    'Add whatever other formatting you want here

                End If

            End If

        End If

    'Go to the next row
    r = r + 1

    Loop

End Sub



