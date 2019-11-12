Sub delRowEven()

Dim 데이터표의마지막행 As Integer

    데이터표의마지막행 = ActiveCell.End(xlDown).Row
    
    For K = 데이터표의마지막행 To 1 Step -1
        
        If K > ActiveCell.Row Then
            If K Mod 2 = 0 Then
                'Cells(K, 2).EntireRow.Delete
                Cells(K, 2).Interior.ColorIndex = 3
                Cells(K, 3).Interior.ColorIndex = 4
                Cells(K, 4).Interior.ColorIndex = 5
                Cells(K, 5).Interior.ColorIndex = 6
            Else
                Cells(K, 2).EntireRow.Hidden = True
            End If
            ElseIf K = ActiveCell.Row Then
                Cells(K, 2).Interior.ColorIndex = 15
                Cells(K, 3).Interior.ColorIndex = 15
                Cells(K, 4).Interior.ColorIndex = 15
                Cells(K, 5).Interior.ColorIndex = 15
        End If
    Next K
        
    
    
End Sub