Sub DeleteDoorsID_ExactLine()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' Looks specifically for "DOORS ID: OVER", followed by any numbers [0-9]@,
        ' and includes the paragraph break (^13) to delete the whole line blank space.
        .Text = "DOORS ID: OVER[0-9]@^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        
        ' Execute across the document instantly
        .Execute Replace:=wdReplaceAll
    End With
    
    MsgBox "Done! All 'DOORS ID: OVER' lines have been completely deleted.", vbInformation
End Sub