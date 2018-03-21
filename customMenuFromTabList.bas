Sub AddMenus()

    Dim cMenu1 As CommandBarControl
'//bulit this long ago. Check for functionality. 

    Dim cbMainMenuBar As CommandBar
    Dim iHelpMenu As Integer
    Dim cbcCutomMenu As CommandBarControl

    On Error Resume Next
        Application.CommandBars("Worksheet Menu Bar").Controls("&Go To Script").Delete
    On Error GoTo 0

    'Set a CommandBar variable to Worksheet menu bar
    Set cbMainMenuBar = Application.CommandBars("Worksheet Menu Bar")

    'Return the Index number of the Help menu
    iHelpMenu = cbMainMenuBar.Controls("Help").Index

    'Add a Control to the "Worksheet Menu Bar" before Help.
    'Set a CommandBarControl variable to it
    Set cbcCutomMenu = cbMainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=iHelpMenu)

    With cbcCutomMenu
        .Caption = "&Go To Script"
    End With

For x = 1 To 36
    With cbcCutomMenu.Controls.Add(Type:=msoControlButton)
            .Caption = ShtASAPSheetIndex.Range("a" & x + 1).Text
            .OnAction = "'ProcessingRequest """ & .Caption & "'"
    End With
Next x
        
End Sub
Sub ProcessingRequest(ByVal shtrequest As String)
    Sheets(shtrequest).Select
End Sub

Sub DeleteMenu()
    On Error Resume Next
        Application.CommandBars("Worksheet Menu Bar").Controls("&Go To Script").Delete
    On Error GoTo 0
End Sub

