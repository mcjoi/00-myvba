
Sub releasePivot()
    
    Set rngHeader = Range("b2:f2")
    Set rngFactor = Range("b2:b15")
    
    rngHeaderCount = rngHeader.Count
    rngFactorCount = rngFactor.Count
    
    For y = 1 To rngHeaderCount
        For x = 1 To rngFactorCount
            If Cells(2, 2).Offset(x, y) <> "" Then
                ActiveCell = Cells(2, 2).Offset(x, 0).Value
                ActiveCell.Offset(0, 1) = Cells(2, 2).Offset(0, y).Value
                ActiveCell.Offset(0, 2) = Cells(2, 2).Offset(x, y).Value
                ActiveCell.Offset(1, 0).Activate
            End If
            
        Next x
    Next y
    
End Sub




