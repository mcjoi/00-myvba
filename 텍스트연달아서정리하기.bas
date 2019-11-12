Sub ADDCHAR()

    On Error GoTo DD
    
    Application.DisplayAlerts = False
        Set refRng = Application.InputBox("참조하실 영역을 선택하세요" & vbCr & _
        "*단, 참조영역의 바로 우측셀에 불러올 값이 있어야 합니다", , , , , , , 8)
        
    Application.DisplayAlerts = True
    
    tRow = ActiveCell.Row
    tCol = ActiveCell.Column
    
    For Each x In refRng
        For y = tRow To ActiveCell.End(xlDown).Row
            If x = Cells(y, tCol) Then
                Cells(y, tCol).Offset(0, 1) = Cells(y, tCol).Offset(0, 1) & " " _
                & x.Offset(0, 1)
            End If
        Next y
    Next x
    
DD:

End Sub



Sub myloop()

On Error GoTo DD

Application.DisplayAlerts = False
    Set myRng = Application.InputBox("참조영역을선택하세요" & vbCr & vbCr & "*참조영역 우측셀에 찾을 값이 있어야 합니다", , , , , , , 8)
Application.DisplayAlerts = True

Count = myRng.Count

selectionCount = Selection.Count

For y = 1 To selectionCount
    Selection(y).Offset(0, 1).ClearContents
    For x = 1 To Count
        If myRng(x) = Selection(y) Then
            Selection(y).Offset(0, 1) = Selection(y).Offset(0, 1) & " " & myRng(x).Offset(0, 1)
        End If
    Next x
Next y

DD:

End Sub




