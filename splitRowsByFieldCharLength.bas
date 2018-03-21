Attribute VB_Name = "Module1"
Option Explicit

Sub breakStringByLength()

'// allows program to run faster - on error, run subroutine "reset()"
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

On Error GoTo done:

'// declare variables
Dim lastRow As Integer
Dim lastColumn As Integer
Dim i As Integer: Dim j As Integer
Dim rIns As Integer
Dim mStr As String
Dim numBreaks As Integer
Dim mStrLen As Integer
Dim foundSpace As Boolean
Dim spaceIndex As Integer
Dim maxChars As Integer
Dim subStr1 As String
Dim subStr2 As String

maxChars = InputBox("Input the maximum number of characters for a single row.", "Max Character Input", 72)
maxChars = CInt(maxChars)

lastRow = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

i = 2
While i < lastRow


    If i Mod 10 = 0 Then
        Application.StatusBar = Format((i / lastRow) * 100, "#.00") & "% Complete   |   Total Row Count = " & lastRow
        'MsgBox (i)
    End If

    '// set value
    mStr = Cells(i, 4).Value

    '//determine if greater than n characters
    mStrLen = Len(mStr)

    If mStrLen > maxChars Then
        '//find space
        spaceIndex = InStrRev(mStr, " ", maxChars)

        '//determine substrings 1 and 2 to paste into cells
        subStr1 = Mid(mStr, 1, spaceIndex)
        subStr2 = Mid(mStr, spaceIndex + 1)

        '// transcribe values down a row
        Cells(i + 1, 4).EntireRow.Insert
        For j = 1 To lastColumn
            Cells(i + 1, j).Value = Cells(i, j)
        Next j

        lastRow = lastRow + 1

        '// write new DESCRIPTION values from sub strings
        Cells(i, 4).Value = subStr1
        Cells(i + 1, 4).Value = subStr2

    End If
    i = i + 1
Wend

done:

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Application.StatusBar = Format((i / lastRow) * 100, "#.00") & "% Complete   |   Total Row Count = " & lastRow



If i = lastRow Then
    MsgBox ("Program successful.  " & i & " of " & lastRow & " rows completed.")
Else
    MsgBox ("Program ended early.  " & i & " of " & lastRow & " rows completed.")
End If

    

End Sub


Sub reset()

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic



End Sub


