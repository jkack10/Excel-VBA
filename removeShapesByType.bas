Attribute VB_Name = "Module1"
Sub ReplacePictures()
    For Each wShape In ActiveSheet.Shapes
        If (wShape.Type = 13) Then
            wShape.TopLeftCell = "Yes"
            wShape.Delete
        End If
    Next wShape
End Sub
