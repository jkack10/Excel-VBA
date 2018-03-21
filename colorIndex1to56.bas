Attribute VB_Name = "Module1"
Sub ColorIndexList()

       ' Begin error trapping.
       On Error GoTo Done

       Range("A1").Value = "Color"
       Range("B1").Value = "Color Index Number"
       ActiveCell.Offset(1, 0).Activate

       ' Begin loop from 1 to 56.
       For xColor = 1 To 56

          ' Apply color and pattern properties to active cell.
          With ActiveCell.Interior
             .ColorIndex = xColor
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
          End With

          ' Put color index in cell to right of active cell.
          ActiveCell.Offset(0, 1).Formula = xColor

          ' Select next cell down.
          ActiveCell.Offset(1, 0).Activate

          ' Increment For loop.
       Next xColor

Done:

   End Sub


